import os
import re
import sqlite3
import hashlib
from typing import List, Dict, Optional
from config.settings import MAX_CHUNK_LINES

class VBNetParser:
    def __init__(self, db_path: str):
        self.db_path = db_path
        self._init_db()

    def _init_db(self):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS CodeChunks (
                    id TEXT PRIMARY KEY,
                    file_path TEXT,
                    block_type TEXT,
                    name TEXT,
                    content TEXT,
                    processed_rules BOOLEAN DEFAULT 0,
                    processed_uml BOOLEAN DEFAULT 0
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS CallGraph (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    caller_file TEXT,
                    caller_name TEXT,
                    caller_type TEXT,
                    callee_name TEXT,
                    call_line TEXT
                )
            ''')
            conn.commit()

    def generate_id(self, file_path: str, block_type: str, name: str) -> str:
        key = f"{file_path}::{block_type}::{name}"
        return hashlib.md5(key.encode('utf-8')).hexdigest()

    # All recognised VB.NET source extensions
    VB_EXTENSIONS = ('.vb', '.vbs', '.bas', '.cls', '.frm')

    def parse_directory(self, root_dir: str):
        total_files = 0
        for root, dirs, files in os.walk(root_dir):
            rel_dir = os.path.relpath(root, root_dir)
            vb_files = [
                f for f in files
                if f.lower().endswith(self.VB_EXTENSIONS)
                and "designer.vb" not in f.lower()
            ]
            if vb_files:
                print(f"  [DIR] {rel_dir} → {len(vb_files)} file(s)")
            for file in vb_files:
                file_path = os.path.join(root, file)
                print(f"    Parsing: {file}")
                self.parse_file(file_path)
                total_files += 1
        print(f"Total VB.NET files parsed: {total_files}")

    def parse_file(self, file_path: str):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except UnicodeDecodeError:
            with open(file_path, 'r', encoding='latin-1') as f:
                content = f.read()

        lines = content.splitlines()
        
        # We will use a line-by-line stack approach to identify Classes, Modules, Subs, Functions
        # This regex matches the start of significant blocks
        start_pattern = re.compile(
            r'^\s*(?:Public\s+|Private\s+|Protected\s+|Friend\s+|Overrides\s+|Overridable\s+|Shared\s+|Partial\s+|MustInherit\s+)*(Class|Module|Structure|Sub|Function|Property)\s+([A-Za-z0-9_]+)', 
            re.IGNORECASE
        )
        # Match ends
        end_pattern = re.compile(
            r'^\s*End\s+(Class|Module|Structure|Sub|Function|Property)', 
            re.IGNORECASE
        )

        stack = []
        chunks = []

        for line_num, line in enumerate(lines):
            end_match = end_pattern.search(line)
            if end_match:
                block_type = end_match.group(1).title()
                # Pop from stack until we match the type
                for i in range(len(stack)-1, -1, -1):
                    if stack[i]['type'].title() == block_type:
                        block = stack.pop(i)
                        block['end_line'] = line_num
                        block['content'] = "\n".join(lines[block['start_line']:block['end_line']+1])
                        chunks.append(block)
                        break
                continue

            start_match = start_pattern.search(line)
            if start_match:
                block_type = start_match.group(1).title()
                name = start_match.group(2)
                # Ignore single line subs maybe? Assuming standard formatting.
                stack.append({
                    'type': block_type,
                    'name': name,
                    'start_line': line_num,
                    'file_path': file_path
                })
        
        chunks = self.split_large_chunks(chunks)
        self.save_chunks(chunks)
        self.extract_call_graph(chunks)

    def split_large_chunks(self, chunks: List[Dict]) -> List[Dict]:
        """Split any chunk that exceeds MAX_CHUNK_LINES into smaller sub-chunks
        with a 20-line overlap for context continuity."""
        result = []
        overlap = 20
        for chunk in chunks:
            lines = chunk['content'].splitlines()
            if len(lines) <= MAX_CHUNK_LINES:
                result.append(chunk)
                continue
            
            part = 1
            start = 0
            while start < len(lines):
                end = min(start + MAX_CHUNK_LINES, len(lines))
                sub_content = "\n".join(lines[start:end])
                result.append({
                    'type': chunk['type'],
                    'name': f"{chunk['name']}_part{part}",
                    'start_line': chunk['start_line'] + start,
                    'end_line': chunk['start_line'] + end - 1,
                    'file_path': chunk['file_path'],
                    'content': sub_content
                })
                start = end - overlap if end < len(lines) else len(lines)
                part += 1
            print(f"    ⚠ Split '{chunk['name']}' ({len(lines)} lines) into {part-1} sub-chunks")
        return result

    def save_chunks(self, chunks: List[Dict]):
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            for chunk in chunks:
                chunk_id = self.generate_id(chunk['file_path'], chunk['type'], chunk['name'])
                cursor.execute(
                    '''INSERT OR IGNORE INTO CodeChunks 
                       (id, file_path, block_type, name, content) 
                       VALUES (?, ?, ?, ?, ?)''',
                    (chunk_id, chunk['file_path'], chunk['type'], chunk['name'], chunk['content'])
                )
            conn.commit()

    def extract_call_graph(self, chunks: List[Dict]):
        """Scan each chunk for method/function calls and store caller->callee relationships."""
        # Patterns to detect VB method calls
        call_patterns = [
            re.compile(r'(?:Call\s+)(\w+)', re.IGNORECASE),                               # Call MethodName
            re.compile(r'(?:Me\.)(\w+)\s*\(', re.IGNORECASE),                             # Me.MethodName(
            re.compile(r'(\w+)\.([A-Z]\w+)\s*\(', re.IGNORECASE),                         # obj.MethodName(
            re.compile(r'^\s*(\w+)\s*(?:\(|\s+=\s+\w+\.)', re.IGNORECASE | re.MULTILINE),  # Direct call
        ]
        
        rows = []
        for chunk in chunks:
            if chunk['type'] in ('Class', 'Module', 'Structure'):
                continue  # Only extract calls from methods, not class-level blocks
            
            for line in chunk['content'].splitlines():
                stripped = line.strip()
                if not stripped or stripped.startswith("'") or stripped.upper().startswith("REM "):
                    continue
                
                for pat in call_patterns:
                    for m in pat.finditer(line):
                        if len(m.groups()) == 2 and m.group(2):
                            callee = f"{m.group(1)}.{m.group(2)}"
                        else:
                            callee = m.group(1)
                        # Filter basic keywords and types
                        if callee.lower() not in ('string', 'integer', 'boolean', 'true', 'false', 'msgbox'):
                            rows.append((chunk['file_path'], chunk['name'], chunk['type'], callee, stripped))
        
        if rows:
            with sqlite3.connect(self.db_path) as conn:
                conn.executemany(
                    'INSERT INTO CallGraph (caller_file, caller_name, caller_type, callee_name, call_line) VALUES (?, ?, ?, ?, ?)',
                    rows
                )
                conn.commit()

if __name__ == "__main__":
    # Test script usage
    import sys
    sys.path.append(os.path.join(os.path.dirname(__file__), ".."))
    from config.settings import METADATA_DB_PATH, INPUT_DIR
    
    parser = VBNetParser(METADATA_DB_PATH)
    print("Parsing input directory...")
    parser.parse_directory(INPUT_DIR)
    print("Parsing complete.")
