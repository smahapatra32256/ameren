import os
import sqlite3
import pandas as pd
from tqdm import tqdm
import vertexai

from config.settings import (
    PROJECT_ID, REGION, METADATA_DB_PATH, FAISS_INDEX_PATH, 
    INPUT_DIR, OUTPUT_DIR, RULES_OUTPUT_DIR, UML_OUTPUT_DIR, UML_COMPONENTS_DIR, MODEL_NAME
)
from parsers.vbnet_parser import VBNetParser
from vector_store.faiss_manager import FAISSManager
from llm_agents.rule_extractor import RuleExtractor
from llm_agents.uml_generator import UMLGenerator
import urllib.request

original_urlopen = urllib.request.urlopen
def patched_urlopen(url, data=None, timeout=None, **kwargs):
    if isinstance(url, urllib.request.Request):
        url.add_header('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36')
    elif isinstance(url, str):
        url = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
    
    if timeout is None:
        return original_urlopen(url, data, **kwargs)
    else:
        return original_urlopen(url, data, timeout, **kwargs)
urllib.request.urlopen = patched_urlopen

class PipelineOrchestrator:
    def __init__(self):
        # Initialize Google Cloud Vertex AI
        vertexai.init(project=PROJECT_ID, location=REGION)
        
        self.parser = VBNetParser(METADATA_DB_PATH)
        self.faiss_manager = FAISSManager(METADATA_DB_PATH, FAISS_INDEX_PATH, PROJECT_ID, REGION)
        self.rule_extractor = RuleExtractor(MODEL_NAME)
        self.uml_generator = UMLGenerator(MODEL_NAME)

    def run_parser_and_indexer(self):
        print(f"--- 1. Parsing input code from {INPUT_DIR} ---")
        self.parser.parse_directory(INPUT_DIR)
        
        print("--- 2. Building FAISS index ---")
        self.faiss_manager.build_index()

    def process_workload(self):
        print("--- 3. Starting Map-Reduce LLM Extraction ---")
        
        with sqlite3.connect(METADATA_DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            # Fetch unprocessed chunks
            cursor.execute("SELECT * FROM CodeChunks WHERE processed_rules = 0 OR processed_uml = 0")
            rows = cursor.fetchall()
            
        if not rows:
            print("No tasks pending.")
            return

        print(f"Total chunks to process: {len(rows)}")
        
        all_rules_df = pd.DataFrame()

        for row in tqdm(rows, desc="Processing chunks"):
            chunk_dict = dict(row)
            
            # Fetch context using RAG
            context_chunks = self.faiss_manager.retrieve_context(chunk_dict['content'], k=3)
            # Remove self from context
            context_chunks = [c for c in context_chunks if c['id'] != chunk_dict['id']]

            # Rule Extraction
            if not chunk_dict['processed_rules']:
                rules_df = self.rule_extractor.process_chunk(chunk_dict, context_chunks)
                if not rules_df.empty:
                    rules_path = os.path.join(RULES_OUTPUT_DIR, "consolidated_rules.csv")
                    # If exists, append to avoid data loss
                    if os.path.exists(rules_path):
                        rules_df.to_csv(rules_path, mode='a', header=False, index=False, sep='|')
                    else:
                        rules_df.to_csv(rules_path, index=False, sep='|')
                
                # Update DB
                with sqlite3.connect(METADATA_DB_PATH) as conn:
                    conn.cursor().execute("UPDATE CodeChunks SET processed_rules = 1 WHERE id = ?", (chunk_dict['id'],))
                    conn.commit()

            # UML Extraction
            if not chunk_dict['processed_uml']:
                uml_str = self.uml_generator.process_chunk(chunk_dict, context_chunks)
                if uml_str:
                    file_safe_name = "".join([c if c.isalnum() else "_" for c in chunk_dict['name']])
                    uml_file_path = os.path.join(UML_COMPONENTS_DIR, f"{file_safe_name}_{chunk_dict['id'][:6]}.puml")
                    with open(uml_file_path, "w", encoding="utf-8") as f:
                        f.write(uml_str)
                        
                    try:
                        from plantuml import PlantUML
                        plantuml_server = PlantUML(url="http://www.plantuml.com/plantuml/img/")
                        png_bytes = plantuml_server.processes(uml_str)
                        png_path = os.path.join(UML_COMPONENTS_DIR, f"{file_safe_name}_{chunk_dict['id'][:6]}.png")
                        if png_bytes:
                            with open(png_path, "wb") as f:
                                f.write(png_bytes)
                    except Exception as e:
                        print(f"Failed to generate PNG for {file_safe_name}: {e}")

                # Update DB
                with sqlite3.connect(METADATA_DB_PATH) as conn:
                    conn.cursor().execute("UPDATE CodeChunks SET processed_uml = 1 WHERE id = ?", (chunk_dict['id'],))
                    conn.commit()
                    
        print("--- 4. Batch logic complete ---")
        
    def generate_macro_architecture(self):
        print("--- 5. Generating Complete End-To-End Architecture UML ---")
        try:
            with sqlite3.connect(METADATA_DB_PATH) as conn:
                conn.row_factory = sqlite3.Row
                c = conn.cursor()
                c.execute("SELECT block_type, name FROM CodeChunks")
                all_chunks = c.fetchall()
                names = [f"{r['block_type']} {r['name']}" for r in all_chunks]

            if names:
                macro_uml = self.uml_generator.generate_macro_uml(names)
                if macro_uml:
                    macro_path = os.path.join(UML_OUTPUT_DIR, "End_To_End_Flow.puml")
                    with open(macro_path, "w", encoding="utf-8") as f:
                        f.write(macro_uml)
                    from plantuml import PlantUML
                    import requests
                    p_server = PlantUML(url="http://www.plantuml.com/plantuml/png/")
                    req_url = p_server.get_url(macro_uml)
                    r = requests.get(req_url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
                    if r.status_code == 200:
                        with open(os.path.join(UML_OUTPUT_DIR, "End_To_End_Flow.png"), "wb") as f:
                            f.write(r.content)
                    else:
                        print("PlantUML Server Denied Request: ", r.status_code)

        except Exception as e:
            print(f"Failed to generate End To End Flow UML: {e}")
            
        print(f"Extraction tasks successfully wrapped. Please check '{OUTPUT_DIR}' for all generated schemas and rules.")

if __name__ == "__main__":
    orchestrator = PipelineOrchestrator()
    orchestrator.run_parser_and_indexer()
    orchestrator.process_workload()
    orchestrator.generate_macro_architecture()
    print("Pipeline Execution Complete!")
