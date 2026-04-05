import hashlib
import os
import re
import sqlite3
from typing import Dict, List, Optional, Set, Tuple

from config.settings import MAX_CHUNK_LINES


class VBNetParser:
    # All recognised VB.NET/VB6 source extensions
    VB_EXTENSIONS = ('.vb', '.vbs', '.bas', '.cls', '.frm')

    # Common non-business call tokens that create noisy false-positive edges
    EXCLUDED_CALLEE_TOKENS = {
        "and", "array", "as", "boolean", "byref", "byval", "call", "case", "cbool", "cbyte",
        "cdate", "cdbl", "cint", "clng", "const", "csng", "cstr", "date", "dim", "do", "each",
        "else", "elseif", "end", "error", "exit", "false", "fields", "for", "format", "friend",
        "function", "get", "if", "integer", "is", "item", "len", "let", "loop", "me", "mod",
        "msgbox", "new", "next", "nothing", "on", "option", "or", "private", "property", "public",
        "rem", "resume", "select", "set", "shared", "step", "string", "sub", "then", "to", "trim",
        "true", "wend", "while", "with", "xor"
    }

    def __init__(self, db_path: str):
        self.db_path = db_path
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(
                '''
                CREATE TABLE IF NOT EXISTS CodeChunks (
                    id TEXT PRIMARY KEY,
                    file_path TEXT,
                    block_type TEXT,
                    name TEXT,
                    content TEXT,
                    content_hash TEXT,
                    manifest TEXT,
                    processed_rules BOOLEAN DEFAULT 0,
                    processed_manifest BOOLEAN DEFAULT 0,
                    processed_uml BOOLEAN DEFAULT 0
                )
                '''
            )
            cursor.execute(
                '''
                CREATE TABLE IF NOT EXISTS CallGraph (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    caller_file TEXT,
                    caller_name TEXT,
                    caller_type TEXT,
                    callee_name TEXT,
                    call_line TEXT
                )
                '''
            )

            # Lightweight schema migrations for existing metadata.db files
            existing_cols = {row[1] for row in cursor.execute("PRAGMA table_info(CodeChunks)")}
            if "content_hash" not in existing_cols:
                cursor.execute("ALTER TABLE CodeChunks ADD COLUMN content_hash TEXT")
            if "manifest" not in existing_cols:
                cursor.execute("ALTER TABLE CodeChunks ADD COLUMN manifest TEXT")
            if "processed_manifest" not in existing_cols:
                cursor.execute("ALTER TABLE CodeChunks ADD COLUMN processed_manifest BOOLEAN DEFAULT 0")

            cursor.execute("CREATE INDEX IF NOT EXISTS idx_codechunks_file ON CodeChunks(file_path)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_codechunks_rules ON CodeChunks(processed_rules)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_codechunks_manifest ON CodeChunks(processed_manifest)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_codechunks_uml ON CodeChunks(processed_uml)")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_callgraph_caller_file ON CallGraph(caller_file)")
            conn.commit()

    def generate_id(
        self,
        file_path: str,
        block_type: str,
        name: str,
        start_line: Optional[int] = None,
    ) -> str:
        key = f"{file_path}::{block_type}::{name}"
        if start_line is not None:
            key += f"::{start_line}"
        return hashlib.md5(key.encode("utf-8")).hexdigest()

    def _hash_content(self, content: str) -> str:
        return hashlib.sha1(content.encode("utf-8", errors="ignore")).hexdigest()

    def parse_directory(self, root_dir: str):
        total_files = 0
        for root, _, files in os.walk(root_dir):
            rel_dir = os.path.relpath(root, root_dir)
            vb_files = [
                f for f in files
                if f.lower().endswith(self.VB_EXTENSIONS)
                and "designer.vb" not in f.lower()
            ]
            if vb_files:
                print(f"  [DIR] {rel_dir} -> {len(vb_files)} file(s)")
            for file in vb_files:
                file_path = os.path.join(root, file)
                print(f"    Parsing: {file}")
                self.parse_file(file_path)
                total_files += 1
        print(f"Total VB.NET files parsed: {total_files}")

    def parse_file(self, file_path: str):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                content = f.read()
        except UnicodeDecodeError:
            with open(file_path, "r", encoding="latin-1") as f:
                content = f.read()

        lines = content.splitlines()
        method_chunks = self._extract_blocks(lines, file_path)

        # Build call graph from original blocks (before chunk splitting) to avoid overlap duplicates.
        self.extract_call_graph(file_path, method_chunks)

        split_chunks = self.split_large_chunks(method_chunks)
        self.save_chunks(file_path, split_chunks)

    def _extract_blocks(self, lines: List[str], file_path: str) -> List[Dict]:
        chunks: List[Dict] = []
        stack: List[Dict] = []

        modifier = (
            r"(?:Public|Private|Protected|Friend|Shared|Partial|MustInherit|NotInheritable|"
            r"Shadows|Overloads|Overrides|Overridable|Static|Default)\s+"
        )
        attributes = r"(?:<[^>]+>\s*)*"

        type_start = re.compile(
            rf"^\s*{attributes}(?:{modifier})*(Class|Module|Structure)\s+([A-Za-z_][A-Za-z0-9_]*)\b",
            re.IGNORECASE,
        )
        method_start = re.compile(
            rf"^\s*{attributes}(?:{modifier})*(Sub|Function)\s+([A-Za-z_][A-Za-z0-9_]*)\b",
            re.IGNORECASE,
        )
        property_start = re.compile(
            rf"^\s*{attributes}(?:{modifier})*Property\s+(Get|Let|Set)\s+([A-Za-z_][A-Za-z0-9_]*)\b",
            re.IGNORECASE,
        )
        end_pattern = re.compile(
            r"^\s*End\s+(Class|Module|Structure|Sub|Function|Property)\b",
            re.IGNORECASE,
        )

        for line_num, line in enumerate(lines):
            end_match = end_pattern.search(line)
            if end_match:
                end_type = end_match.group(1).title()
                for i in range(len(stack) - 1, -1, -1):
                    if stack[i]["type"] == end_type:
                        block = stack.pop(i)
                        block["end_line"] = line_num
                        block["content"] = "\n".join(lines[block["start_line"] : line_num + 1])
                        chunks.append(block)
                        break
                continue

            property_match = property_start.search(line)
            if property_match:
                accessor = property_match.group(1).lower()
                prop_name = property_match.group(2)
                stack.append(
                    {
                        "type": "Property",
                        "name": f"{prop_name}_{accessor}",
                        "start_line": line_num,
                        "file_path": file_path,
                    }
                )
                continue

            method_match = method_start.search(line)
            if method_match:
                stack.append(
                    {
                        "type": method_match.group(1).title(),
                        "name": method_match.group(2),
                        "start_line": line_num,
                        "file_path": file_path,
                    }
                )
                continue

            type_match = type_start.search(line)
            if type_match:
                stack.append(
                    {
                        "type": type_match.group(1).title(),
                        "name": type_match.group(2),
                        "start_line": line_num,
                        "file_path": file_path,
                    }
                )

        return chunks

    def split_large_chunks(self, chunks: List[Dict]) -> List[Dict]:
        """Split chunks over MAX_CHUNK_LINES using overlap, without changing the logical method name."""
        result: List[Dict] = []
        overlap = 20
        for chunk in chunks:
            lines = chunk["content"].splitlines()
            if len(lines) <= MAX_CHUNK_LINES:
                result.append(chunk)
                continue

            start = 0
            parts = 0
            while start < len(lines):
                end = min(start + MAX_CHUNK_LINES, len(lines))
                result.append(
                    {
                        "type": chunk["type"],
                        "name": chunk["name"],
                        "start_line": chunk["start_line"] + start,
                        "end_line": chunk["start_line"] + end - 1,
                        "file_path": chunk["file_path"],
                        "content": "\n".join(lines[start:end]),
                    }
                )
                start = end - overlap if end < len(lines) else len(lines)
                parts += 1
            print(f"    [Split] '{chunk['name']}' ({len(lines)} lines) -> {parts} sub-chunks")
        return result

    def save_chunks(self, file_path: str, chunks: List[Dict]):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT id, content_hash FROM CodeChunks WHERE file_path = ?", (file_path,))
            existing = {row[0]: row[1] for row in cursor.fetchall()}
            incoming_ids: Set[str] = set()

            for chunk in chunks:
                chunk_id = self.generate_id(
                    chunk["file_path"],
                    chunk["type"],
                    chunk["name"],
                    chunk.get("start_line"),
                )
                incoming_ids.add(chunk_id)
                content_hash = self._hash_content(chunk["content"])
                prev_hash = existing.get(chunk_id)

                if prev_hash is None:
                    cursor.execute(
                        '''
                        INSERT INTO CodeChunks
                        (id, file_path, block_type, name, content, content_hash, manifest, processed_rules, processed_manifest, processed_uml)
                        VALUES (?, ?, ?, ?, ?, ?, NULL, 0, 0, 0)
                        ''',
                        (chunk_id, chunk["file_path"], chunk["type"], chunk["name"], chunk["content"], content_hash),
                    )
                elif prev_hash != content_hash:
                    # Content changed: invalidate downstream LLM outputs for this chunk.
                    cursor.execute(
                        '''
                        UPDATE CodeChunks
                        SET block_type = ?, name = ?, content = ?, content_hash = ?,
                            manifest = NULL, processed_rules = 0, processed_manifest = 0, processed_uml = 0
                        WHERE id = ?
                        ''',
                        (chunk["type"], chunk["name"], chunk["content"], content_hash, chunk_id),
                    )

            stale_ids = [chunk_id for chunk_id in existing.keys() if chunk_id not in incoming_ids]
            if stale_ids:
                cursor.executemany("DELETE FROM CodeChunks WHERE id = ?", [(chunk_id,) for chunk_id in stale_ids])

            conn.commit()

    def _is_valid_callee(self, callee: str) -> bool:
        tokens = [t.strip().lower() for t in callee.split(".") if t.strip()]
        if not tokens:
            return False
        return not any(token in self.EXCLUDED_CALLEE_TOKENS for token in tokens)

    def extract_call_graph(self, file_path: str, chunks: List[Dict]):
        """Scan each method/property chunk for caller->callee edges and replace this file's old call-graph rows."""
        call_patterns = [
            re.compile(r"\bCall\s+([A-Za-z_][A-Za-z0-9_\.]*)\b", re.IGNORECASE),
            re.compile(r"\b(?:Me|MyClass)\.([A-Za-z_][A-Za-z0-9_]*)\s*\(", re.IGNORECASE),
            re.compile(r"\b([A-Za-z_][A-Za-z0-9_]*)\.([A-Za-z_][A-Za-z0-9_]*)\s*\(", re.IGNORECASE),
            re.compile(r"^\s*(?:Set\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*\(", re.IGNORECASE),
        ]
        direct_call_control_prefixes = (
            "if ",
            "elseif ",
            "for ",
            "while ",
            "do ",
            "loop",
            "select ",
            "case ",
            "with ",
            "dim ",
            "const ",
            "public ",
            "private ",
            "friend ",
            "sub ",
            "function ",
            "property ",
            "end ",
            "exit ",
            "on error",
            "#",
        )

        rows: Set[Tuple[str, str, str, str, str]] = set()
        for chunk in chunks:
            if chunk["type"] in ("Class", "Module", "Structure"):
                continue

            for line in chunk["content"].splitlines():
                stripped = line.strip()
                if not stripped:
                    continue

                # Skip whole-line comments and strip inline comments.
                if stripped.startswith("'") or stripped.upper().startswith("REM "):
                    continue
                code_line = line.split("'", 1)[0].strip()
                if not code_line:
                    continue

                for pattern_index, pat in enumerate(call_patterns):
                    for match in pat.finditer(code_line):
                        if pattern_index == 0:
                            callee = match.group(1)
                        elif pattern_index == 1:
                            callee = match.group(1)
                        elif pattern_index == 2:
                            callee = f"{match.group(1)}.{match.group(2)}"
                        else:
                            lowered = code_line.lower()
                            if lowered.startswith(direct_call_control_prefixes):
                                continue
                            callee = match.group(1)

                        if not self._is_valid_callee(callee):
                            continue

                        rows.add((file_path, chunk["name"], chunk["type"], callee, code_line))

        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM CallGraph WHERE caller_file = ?", (file_path,))
            if rows:
                cursor.executemany(
                    '''
                    INSERT INTO CallGraph (caller_file, caller_name, caller_type, callee_name, call_line)
                    VALUES (?, ?, ?, ?, ?)
                    ''',
                    sorted(rows),
                )
            conn.commit()


if __name__ == "__main__":
    # Test script usage
    import sys

    sys.path.append(os.path.join(os.path.dirname(__file__), ".."))
    from config.settings import INPUT_DIR, METADATA_DB_PATH

    parser = VBNetParser(METADATA_DB_PATH)
    print("Parsing input directory...")
    parser.parse_directory(INPUT_DIR)
    print("Parsing complete.")
