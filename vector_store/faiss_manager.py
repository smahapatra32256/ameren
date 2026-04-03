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

    def build_index(self, batch_size=100):
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
        for i in range(0, len(rows), batch_size):
            batch = rows[i:i+batch_size]
            texts = [f"Name: {row['name']}\nCode:\n{row['content']}" for row in batch]
            ids = [row['id'] for row in batch]
            
            # Request embeddings
            batch_embeddings = self.embeddings.embed_documents(texts)
            all_embeddings.extend(batch_embeddings)
            self.id_mapping.extend(ids)
            print(f"Processed {i + len(batch)} / {len(rows)}")

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
                
        query_vector = self.embeddings.embed_query(query)
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
