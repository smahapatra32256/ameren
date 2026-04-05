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
        MAX_BATCH_CHARS = 30000  # ~10k tokens, safely under 20k API limit
        batches = []
        current_batch_texts = []
        current_batch_ids = []
        current_chars = 0
        
        for row in rows:
            text = self._truncate(f"Name: {row['name']}\nCode:\n{row['content']}")
            text_len = len(text)
            
            if current_chars + text_len > MAX_BATCH_CHARS and current_batch_texts:
                batches.append((current_batch_texts, current_batch_ids))
                current_batch_texts = []
                current_batch_ids = []
                current_chars = 0
                
            current_batch_texts.append(text)
            current_batch_ids.append(row['id'])
            current_chars += text_len
            
        if current_batch_texts:
            batches.append((current_batch_texts, current_batch_ids))

        id_mapping = []
        
        from concurrent.futures import ThreadPoolExecutor, as_completed
        from tqdm import tqdm
        import time

        def embed_batch(batch_tuple):
            texts, ids = batch_tuple
            for attempt in range(3):
                try:
                    embs = self.embeddings.embed_documents(texts)
                    return embs, ids
                except Exception as e:
                    if '429' in str(e) and attempt < 2:
                        time.sleep(5 * (2 ** attempt))
                    else:
                        return None, ids
            return None, ids

        print(f"Processing {len(batches)} batches in parallel...")
        # Use up to 20 threads to not overwhelm embedding API
        with ThreadPoolExecutor(max_workers=min(20, max(1, len(batches)))) as executor:
            futures = {executor.submit(embed_batch, batch): batch for batch in batches}
            pbar = tqdm(total=len(batches), desc="Embedding Chunks")
            for future in as_completed(futures):
                embs, ids = future.result()
                if embs is not None:
                    all_embeddings.extend(embs)
                    id_mapping.extend(ids)
                else:
                    # Single doc fallback on failure
                    texts, ids = futures[future]
                    for t, doc_id in zip(texts, ids):
                        try:
                            time.sleep(1) # Small delay to prevent rapid 429
                            emb = self.embeddings.embed_documents([t])
                            all_embeddings.extend(emb)
                            id_mapping.append(doc_id)
                        except Exception as e2:
                            print(f"  ⚠ Skipped chunk {doc_id}: {e2}")
                pbar.update(1)
            pbar.close()

        self.id_mapping = id_mapping

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
