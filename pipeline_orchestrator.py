import os
import sqlite3
import threading
import warnings
import pandas as pd
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor, as_completed

warnings.filterwarnings("ignore", category=UserWarning, module="vertexai")
warnings.filterwarnings("ignore", category=DeprecationWarning)

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

# Patch urllib to bypass PlantUML 403
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

# Thread-safe locks
csv_lock = threading.Lock()
db_lock = threading.Lock()

# Max parallel LLM workers
MAX_WORKERS = 50


class PipelineOrchestrator:
    def __init__(self):
        vertexai.init(project=PROJECT_ID, location=REGION)
        self.parser = VBNetParser(METADATA_DB_PATH)
        self.faiss_manager = FAISSManager(METADATA_DB_PATH, FAISS_INDEX_PATH, PROJECT_ID, REGION)

    def run_parser_and_indexer(self):
        print(f"--- 1. Parsing input code from {INPUT_DIR} ---")
        self.parser.parse_directory(INPUT_DIR)
        print("--- 2. Building FAISS index ---")
        self.faiss_manager.build_index()

    def _process_single_chunk(self, chunk_dict, rule_extractor, uml_generator):
        """Process a single chunk: extract rules + generate UML. Thread-safe."""
        try:
            context_chunks = self.faiss_manager.retrieve_context(chunk_dict['content'], k=3)
            context_chunks = [c for c in context_chunks if c['id'] != chunk_dict['id']]
        except Exception:
            context_chunks = []

        # Rule Extraction
        if not chunk_dict['processed_rules']:
            try:
                rules_df = rule_extractor.process_chunk(chunk_dict, context_chunks)
                if not rules_df.empty:
                    rules_path = os.path.join(RULES_OUTPUT_DIR, "consolidated_rules.csv")
                    with csv_lock:
                        if os.path.exists(rules_path):
                            rules_df.to_csv(rules_path, mode='a', header=False, index=False, sep='|')
                        else:
                            rules_df.to_csv(rules_path, index=False, sep='|')
            except Exception as e:
                pass  # Skip silently, don't crash the thread

            with db_lock:
                with sqlite3.connect(METADATA_DB_PATH) as conn:
                    conn.cursor().execute("UPDATE CodeChunks SET processed_rules = 1 WHERE id = ?", (chunk_dict['id'],))
                    conn.commit()

        # UML Extraction
        if not chunk_dict['processed_uml']:
            try:
                uml_str = uml_generator.process_chunk(chunk_dict, context_chunks)
                if uml_str:
                    # Clean name: SourceFile_MethodName (no hash)
                    src_file = os.path.splitext(os.path.basename(chunk_dict['file_path']))[0]
                    method_name = "".join([c if c.isalnum() else "_" for c in chunk_dict['name']])
                    clean_name = f"{src_file}_{method_name}"
                    
                    puml_path = os.path.join(UML_COMPONENTS_DIR, f"{clean_name}.puml")
                    with open(puml_path, "w", encoding="utf-8") as f:
                        f.write(uml_str)

                    try:
                        from plantuml import PlantUML
                        import requests as req_lib
                        p = PlantUML(url="http://www.plantuml.com/plantuml/png/")
                        png_url = p.get_url(uml_str)
                        r = req_lib.get(png_url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=15)
                        if r.status_code == 200 and len(r.content) > 100:
                            png_path = os.path.join(UML_COMPONENTS_DIR, f"{clean_name}.png")
                            with open(png_path, "wb") as f:
                                f.write(r.content)
                            # Delete .puml after successful PNG
                            try:
                                os.remove(puml_path)
                            except OSError:
                                pass
                    except Exception:
                        pass  # Keep .puml as fallback if PNG fails
            except Exception:
                pass

            with db_lock:
                with sqlite3.connect(METADATA_DB_PATH) as conn:
                    conn.cursor().execute("UPDATE CodeChunks SET processed_uml = 1 WHERE id = ?", (chunk_dict['id'],))
                    conn.commit()

        return chunk_dict['id']

    def process_workload(self):
        print("--- 3. Starting Parallel Map-Reduce LLM Extraction ---")

        with sqlite3.connect(METADATA_DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM CodeChunks WHERE processed_rules = 0 OR processed_uml = 0")
            rows = cursor.fetchall()

        if not rows:
            print("No tasks pending.")
            return

        print(f"Total chunks to process: {len(rows)} (using {MAX_WORKERS} parallel workers)")

        # Each thread gets its own LLM model instances to avoid contention
        def make_worker_fn(chunk_dict):
            re = RuleExtractor(MODEL_NAME)
            ug = UMLGenerator(MODEL_NAME)
            return self._process_single_chunk(chunk_dict, re, ug)

        completed = 0
        failed = 0
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {
                executor.submit(make_worker_fn, dict(row)): dict(row)['id']
                for row in rows
            }
            pbar = tqdm(total=len(rows), desc="Processing chunks")
            for future in as_completed(futures):
                try:
                    future.result()
                    completed += 1
                except Exception:
                    failed += 1
                pbar.update(1)
            pbar.close()

        print(f"--- 4. Batch logic complete (✓ {completed} | ✗ {failed}) ---")

    def generate_end_to_end_flow(self):
        """Build a readable End-To-End component diagram from REAL call graph data."""
        print("--- 5. Generating End-To-End Flow from Call Graph ---")

        try:
            with sqlite3.connect(METADATA_DB_PATH) as conn:
                conn.row_factory = sqlite3.Row
                c = conn.cursor()

                # Get all call relationships with source file info
                c.execute("SELECT DISTINCT caller_file, caller_name, callee_name FROM CallGraph")
                edges = c.fetchall()

                # Get all chunks with file info for grouping
                c.execute("SELECT DISTINCT file_path, block_type, name FROM CodeChunks")
                all_chunks = c.fetchall()

            if not edges:
                print("  No call graph edges found. Falling back to LLM-based generation...")
                self._fallback_llm_macro()
                return

            # ── Group methods by their source file (module/class) ──
            from collections import defaultdict
            file_to_methods = defaultdict(set)
            for chunk in all_chunks:
                src = os.path.splitext(os.path.basename(chunk['file_path']))[0]
                file_to_methods[src].add(chunk['name'])

            # ── Build callee targets (external objects being called) ──
            external_targets = set()
            for edge in edges:
                callee = edge['callee_name']
                if '.' in callee:
                    external_targets.add(callee.split('.')[0])

            # ── Build PlantUML Component Diagram ──
            puml = []
            puml.append("@startuml")
            puml.append("!theme cerulean")
            puml.append("skinparam componentFontSize 11")
            puml.append("skinparam packageFontSize 13")
            puml.append("skinparam arrowFontSize 9")
            puml.append("skinparam defaultFontName Segoe UI")
            puml.append("skinparam linetype ortho")
            puml.append("left to right direction")
            puml.append("")

            # Create packages for each source file with their methods
            sanitize = lambda s: "".join(c if c.isalnum() else "_" for c in s)
            
            for src_file, methods in sorted(file_to_methods.items()):
                safe_pkg = sanitize(src_file)
                puml.append(f'package "{src_file}" as pkg_{safe_pkg} {{')
                for method in sorted(methods):
                    safe_m = sanitize(method)
                    puml.append(f'  [{method}] as {safe_pkg}_{safe_m}')
                puml.append("}")
                puml.append("")

            # External dependencies (objects referenced but not in our codebase)
            if external_targets:
                puml.append('package "External Dependencies" as pkg_external {')
                for ext in sorted(external_targets):
                    safe_ext = sanitize(ext)
                    puml.append(f'  [{ext}] as ext_{safe_ext}')
                puml.append("}")
                puml.append("")

            # Add edges
            for edge in edges:
                caller_src = os.path.splitext(os.path.basename(edge['caller_file']))[0]
                caller_name = edge['caller_name']
                callee = edge['callee_name']
                
                safe_src = sanitize(caller_src)
                safe_caller = sanitize(caller_name)
                from_id = f"{safe_src}_{safe_caller}"

                if '.' in callee:
                    callee_obj = callee.split('.')[0]
                    callee_method = callee.split('.')[1]
                    to_id = f"ext_{sanitize(callee_obj)}"
                    puml.append(f'{from_id} --> {to_id} : {callee_method}()')
                else:
                    # Internal call — find which package it belongs to
                    safe_callee = sanitize(callee)
                    found = False
                    for s, ms in file_to_methods.items():
                        if callee in ms:
                            to_id = f"{sanitize(s)}_{safe_callee}"
                            puml.append(f'{from_id} --> {to_id}')
                            found = True
                            break
                    if not found:
                        puml.append(f'{from_id} --> ext_{safe_callee} : call')

            puml.append("@enduml")
            puml_content = "\n".join(puml)

            # Save .puml
            puml_path = os.path.join(UML_OUTPUT_DIR, "End_To_End_Flow.puml")
            with open(puml_path, "w", encoding="utf-8") as f:
                f.write(puml_content)

            # Generate PNG / SVG via PlantUML
            self._render_puml_to_image(puml_content, UML_OUTPUT_DIR, "End_To_End_Flow")

        except Exception as e:
            print(f"  Failed to generate End-To-End flow: {e}")
            import traceback; traceback.print_exc()

        # Also export call graph as CSV for reference
        try:
            with sqlite3.connect(METADATA_DB_PATH) as conn:
                df = pd.read_sql_query("SELECT caller_file, caller_name, caller_type, callee_name, call_line FROM CallGraph", conn)
                cg_path = os.path.join(RULES_OUTPUT_DIR, "call_graph.csv")
                df.to_csv(cg_path, index=False, sep='|')
                print(f"  ✓ Call graph exported: {len(df)} edges → {cg_path}")
        except Exception:
            pass

        print(f"Extraction complete. Check '{OUTPUT_DIR}' for all outputs.")

    def _render_puml_to_image(self, puml_content, output_dir, base_name):
        """Render PlantUML content to PNG (or SVG fallback) with retry backoff. Deletes .puml on success."""
        import requests
        import time
        from plantuml import PlantUML
        puml_path = os.path.join(output_dir, f"{base_name}.puml")

        max_retries = 5
        for attempt in range(max_retries):
            try:
                # Try PNG
                p = PlantUML(url="http://www.plantuml.com/plantuml/png/")
                req_url = p.get_url(puml_content)
                r = requests.get(req_url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}, timeout=60)

                if r.status_code == 200 and len(r.content) > 100:
                    with open(os.path.join(output_dir, f"{base_name}.png"), "wb") as f:
                        f.write(r.content)
                    print(f"  ✓ {base_name}.png generated successfully")
                    try: os.remove(puml_path)
                    except OSError: pass
                    return

                if r.status_code in (429, 500, 502, 503, 509):
                    wait = min(10 * (2 ** attempt), 60)
                    print(f"  ⏳ PlantUML rate-limited ({r.status_code}), retrying in {wait}s... (attempt {attempt+1}/{max_retries})")
                    time.sleep(wait)
                    continue

                # Try SVG fallback (non-retryable status)
                p_svg = PlantUML(url="http://www.plantuml.com/plantuml/svg/")
                svg_url = p_svg.get_url(puml_content)
                r2 = requests.get(svg_url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=60)
                if r2.status_code == 200:
                    with open(os.path.join(output_dir, f"{base_name}.svg"), "wb") as f:
                        f.write(r2.content)
                    print(f"  ✓ {base_name}.svg generated (SVG fallback)")
                    try: os.remove(puml_path)
                    except OSError: pass
                    return

                print(f"  PlantUML server returned {r.status_code} — .puml file kept for manual rendering")
                return

            except requests.exceptions.Timeout:
                wait = min(10 * (2 ** attempt), 60)
                print(f"  ⏳ PlantUML timed out, retrying in {wait}s... (attempt {attempt+1}/{max_retries})")
                time.sleep(wait)
            except Exception as e:
                print(f"  Image generation failed ({e}) — .puml file kept for manual rendering")
                return

        print(f"  ✗ All {max_retries} attempts failed — .puml file kept for manual rendering")

    def _fallback_llm_macro(self):
        """Fallback: use LLM to generate macro diagram if no call graph data exists."""
        try:
            uml_gen = UMLGenerator(MODEL_NAME)
            with sqlite3.connect(METADATA_DB_PATH) as conn:
                conn.row_factory = sqlite3.Row
                c = conn.cursor()
                c.execute("SELECT block_type, name FROM CodeChunks")
                all_chunks = c.fetchall()
                names = [f"{r['block_type']} {r['name']}" for r in all_chunks]

            if names:
                macro_uml = uml_gen.generate_macro_uml(names)
                if macro_uml:
                    with open(os.path.join(UML_OUTPUT_DIR, "End_To_End_Flow.puml"), "w", encoding="utf-8") as f:
                        f.write(macro_uml)

                    from plantuml import PlantUML
                    import requests
                    p_server = PlantUML(url="http://www.plantuml.com/plantuml/png/")
                    req_url = p_server.get_url(macro_uml)
                    r = requests.get(req_url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
                    if r.status_code == 200:
                        with open(os.path.join(UML_OUTPUT_DIR, "End_To_End_Flow.png"), "wb") as f:
                            f.write(r.content)
        except Exception as e:
            print(f"  Fallback LLM macro also failed: {e}")


if __name__ == "__main__":
    orchestrator = PipelineOrchestrator()
    orchestrator.run_parser_and_indexer()
    orchestrator.process_workload()
    orchestrator.generate_end_to_end_flow()
    print("Pipeline Execution Complete!")
