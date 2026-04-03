import os
import re
import sqlite3
import hashlib
from typing import List, Dict, Optional

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
            conn.commit()

    def generate_id(self, file_path: str, block_type: str, name: str) -> str:
        key = f"{file_path}::{block_type}::{name}"
        return hashlib.md5(key.encode('utf-8')).hexdigest()

    def parse_directory(self, root_dir: str):
        for root, _, files in os.walk(root_dir):
            for file in files:
                if file.lower().endswith(".vb") and not "designer.vb" in file.lower():
                    file_path = os.path.join(root, file)
                    self.parse_file(file_path)

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
        
        self.save_chunks(chunks)

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

if __name__ == "__main__":
    # Test script usage
    import sys
    sys.path.append(os.path.join(os.path.dirname(__file__), ".."))
    from config.settings import METADATA_DB_PATH, INPUT_DIR
    
    parser = VBNetParser(METADATA_DB_PATH)
    print("Parsing input directory...")
    parser.parse_directory(INPUT_DIR)
    print("Parsing complete.")
