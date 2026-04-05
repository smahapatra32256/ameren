import os
import pickle
import sqlite3
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, List, Optional, Tuple

import faiss
import numpy as np
from langchain_google_vertexai import VertexAIEmbeddings
from tqdm import tqdm

from config.settings import MAX_EMBEDDING_CHARS


class FAISSManager:
    def __init__(
        self,
        db_path: str,
        index_path: str,
        project_id: str,
        region: str,
        model_name: str = "text-embedding-004",
    ):
        self.db_path = db_path
        self.index_path = index_path
        self.index = None
        self.id_mapping: List[str] = []  # Maps faiss index (int) -> chunk id (str)
        self.embeddings = VertexAIEmbeddings(
            project=project_id,
            location=region,
            model_name=model_name,
        )

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.execute("PRAGMA journal_mode=MEMORY")
        conn.execute("PRAGMA temp_store=MEMORY")
        return conn

    def _truncate(self, text: str, max_chars: int = MAX_EMBEDDING_CHARS) -> str:
        if len(text) <= max_chars:
            return text
        return text[:max_chars]

    def build_index(self):
        with self._connect() as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute("SELECT id, name, content FROM CodeChunks").fetchall()

        if not rows:
            print("No chunks found in DB. Run parser first.")
            return

        print(f"Building FAISS index for {len(rows)} chunks...")

        all_embeddings: List[List[float]] = []
        id_mapping: List[str] = []
        skipped_docs: List[Tuple[str, str]] = []

        max_batch_chars = 30000  # Keep under embedding request limits with margin.
        batches: List[Tuple[List[str], List[str]]] = []
        current_batch_texts: List[str] = []
        current_batch_ids: List[str] = []
        current_chars = 0

        for row in rows:
            text = self._truncate(f"Name: {row['name']}\nCode:\n{row['content']}")
            text_len = len(text)
            if current_chars + text_len > max_batch_chars and current_batch_texts:
                batches.append((current_batch_texts, current_batch_ids))
                current_batch_texts = []
                current_batch_ids = []
                current_chars = 0

            current_batch_texts.append(text)
            current_batch_ids.append(row["id"])
            current_chars += text_len

        if current_batch_texts:
            batches.append((current_batch_texts, current_batch_ids))

        def embed_batch(batch_tuple: Tuple[List[str], List[str]]) -> Tuple[List[List[float]], List[str], Optional[str]]:
            texts, ids = batch_tuple
            last_error = ""
            for attempt in range(3):
                try:
                    embs = self.embeddings.embed_documents(texts)
                    return embs, ids, None
                except Exception as e:
                    last_error = str(e)
                    if "429" in last_error and attempt < 2:
                        time.sleep(5 * (2 ** attempt))
                    else:
                        break
            return [], ids, last_error or "Unknown embedding error"

        print(f"Processing {len(batches)} embedding batches in parallel...")
        with ThreadPoolExecutor(max_workers=min(20, max(1, len(batches)))) as executor:
            futures = {executor.submit(embed_batch, batch): batch for batch in batches}
            pbar = tqdm(total=len(batches), desc="Embedding Chunks")
            for future in as_completed(futures):
                embs, ids, err = future.result()
                if embs:
                    all_embeddings.extend(embs)
                    id_mapping.extend(ids)
                else:
                    texts, batch_ids = futures[future]
                    for t, doc_id in zip(texts, batch_ids):
                        try:
                            time.sleep(1)
                            emb = self.embeddings.embed_documents([t])
                            all_embeddings.extend(emb)
                            id_mapping.append(doc_id)
                        except Exception as e2:
                            skipped_docs.append((doc_id, str(e2)))
                pbar.update(1)
            pbar.close()

        self.id_mapping = id_mapping

        if not all_embeddings:
            raise RuntimeError("FAISS index build failed: no embeddings were generated.")

        dimension = len(all_embeddings[0])
        self.index = faiss.IndexFlatL2(dimension)
        self.index.add(np.array(all_embeddings).astype("float32"))
        self.save_index()

        print(f"FAISS index built successfully with {len(self.id_mapping)} vectors.")
        if skipped_docs:
            print(f"WARNING: Skipped {len(skipped_docs)} chunks during embedding. See first 5 below:")
            for doc_id, reason in skipped_docs[:5]:
                print(f"  - {doc_id}: {reason}")

    def save_index(self):
        faiss.write_index(self.index, self.index_path)
        with open(self.index_path + ".map", "wb") as f:
            pickle.dump(self.id_mapping, f)

    def load_index(self):
        if not os.path.exists(self.index_path):
            return False
        self.index = faiss.read_index(self.index_path)
        with open(self.index_path + ".map", "rb") as f:
            self.id_mapping = pickle.load(f)
        return True

    def retrieve_context(self, query: str, k: int = 3) -> List[Dict]:
        if self.index is None:
            if not self.load_index():
                print("Index not found. Build it first.")
                return []

        query_vector = self.embeddings.embed_query(self._truncate(query))
        _, idx_matrix = self.index.search(np.array([query_vector]).astype("float32"), k)

        valid_idx = [idx for idx in idx_matrix[0] if idx != -1 and idx < len(self.id_mapping)]
        if not valid_idx:
            return []

        chunk_ids = [self.id_mapping[idx] for idx in valid_idx]
        placeholders = ",".join(["?"] * len(chunk_ids))
        query_sql = f"SELECT * FROM CodeChunks WHERE id IN ({placeholders})"

        with self._connect() as conn:
            conn.row_factory = sqlite3.Row
            rows = conn.execute(query_sql, chunk_ids).fetchall()

        by_id = {row["id"]: dict(row) for row in rows}
        # Preserve FAISS rank order
        return [by_id[cid] for cid in chunk_ids if cid in by_id]


if __name__ == "__main__":
    import sys

    sys.path.append(os.path.join(os.path.dirname(__file__), ".."))
    from config.settings import EMBEDDING_MODEL_NAME, METADATA_DB_PATH, FAISS_INDEX_PATH, PROJECT_ID, REGION

    manager = FAISSManager(METADATA_DB_PATH, FAISS_INDEX_PATH, PROJECT_ID, REGION, EMBEDDING_MODEL_NAME)
    manager.build_index()
