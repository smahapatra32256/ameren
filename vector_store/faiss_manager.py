import os
import sqlite3
import numpy as np
import faiss
import pickle
from typing import List, Dict, Tuple
from langchain_google_vertexai import VertexAIEmbeddings

class FAISSManager:
    def __init__(self, db_path: str, index_path: str, project_id: str, region: str, model_name: str="text-embedding-004"):
        self.db_path = db_path
        self.index_path = index_path
        self.index = None
        self.id_mapping = [] # Maps faiss index (int) to chunk id (str)
        self.embeddings = VertexAIEmbeddings(
            project=project_id,
            location=region,
            model_name=model_name
        )

    def _truncate(self, text: str, max_chars: int = 15000) -> str:
        """Truncate text to stay safely under the embedding model's token limit."""
        if len(text) <= max_chars:
            return text
        return text[:max_chars]

    def build_index(self):
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("SELECT id, name, content FROM CodeChunks")
            rows = cursor.fetchall()

        if not rows:
            print("No chunks found in DB. Run parser first.")
            return

        print(f"Building FAISS index for {len(rows)} chunks...")
        
        all_embeddings = []
        # Dynamic batching: accumulate until we approach the token limit
        MAX_BATCH_CHARS = 30000  # ~10k tokens, safely under 20k API limit
        batch_texts = []
        batch_ids = []
        batch_chars = 0
        processed = 0

        def flush_batch():
            nonlocal batch_texts, batch_ids, batch_chars, processed
            if not batch_texts:
                return
            try:
                embs = self.embeddings.embed_documents(batch_texts)
                all_embeddings.extend(embs)
                self.id_mapping.extend(batch_ids)
            except Exception as e:
                # Fallback: embed one at a time
                print(f"  Batch failed ({e}), falling back to single-doc mode...")
                for t, doc_id in zip(batch_texts, batch_ids):
                    try:
                        emb = self.embeddings.embed_documents([t])
                        all_embeddings.extend(emb)
                        self.id_mapping.append(doc_id)
                    except Exception as e2:
                        print(f"  ⚠ Skipped chunk {doc_id}: {e2}")
            processed += len(batch_texts)
            print(f"Processed {processed} / {len(rows)}")
            batch_texts = []
            batch_ids = []
            batch_chars = 0

        for row in rows:
            text = self._truncate(f"Name: {row['name']}\nCode:\n{row['content']}")
            text_len = len(text)
            
            # Flush if adding this doc would exceed our char budget
            if batch_chars + text_len > MAX_BATCH_CHARS and batch_texts:
                flush_batch()
            
            batch_texts.append(text)
            batch_ids.append(row['id'])
            batch_chars += text_len

        flush_batch()  # Final flush

        if all_embeddings:
            dimension = len(all_embeddings[0])
            self.index = faiss.IndexFlatL2(dimension)
            self.index.add(np.array(all_embeddings).astype('float32'))
            
            self.save_index()
            print("FAISS index built successfully.")

    def save_index(self):
        faiss.write_index(self.index, self.index_path)
        with open(self.index_path + ".map", 'wb') as f:
            pickle.dump(self.id_mapping, f)

    def load_index(self):
        if os.path.exists(self.index_path):
            self.index = faiss.read_index(self.index_path)
            with open(self.index_path + ".map", 'rb') as f:
                self.id_mapping = pickle.load(f)
            return True
        return False

    def retrieve_context(self, query: str, k: int = 3) -> List[Dict]:
        if not self.index:
            if not self.load_index():
                print("Index not found. Build it first.")
                return []
                
        query_vector = self.embeddings.embed_query(self._truncate(query))
        D, I = self.index.search(np.array([query_vector]).astype('float32'), k)
        
        results = []
        with sqlite3.connect(self.db_path) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            for idx in I[0]:
                if idx < len(self.id_mapping) and idx != -1:
                    chunk_id = self.id_mapping[idx]
                    cursor.execute("SELECT * FROM CodeChunks WHERE id=?", (chunk_id,))
                    row = cursor.fetchone()
                    if row:
                        results.append(dict(row))
        return results

if __name__ == "__main__":
    import sys
    sys.path.append(os.path.join(os.path.dirname(__file__), ".."))
    from config.settings import METADATA_DB_PATH, FAISS_INDEX_PATH, PROJECT_ID, REGION, EMBEDDING_MODEL_NAME
    
    manager = FAISSManager(METADATA_DB_PATH, FAISS_INDEX_PATH, PROJECT_ID, REGION, EMBEDDING_MODEL_NAME)
    manager.build_index()
