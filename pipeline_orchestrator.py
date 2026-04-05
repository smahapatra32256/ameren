import os
import sys
import sqlite3
import threading
import warnings
import pandas as pd
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor, as_completed

# Force utf-8 encoding for stdout on Windows to prevent UnicodeEncodeError
sys.stdout.reconfigure(encoding='utf-8')

warnings.filterwarnings("ignore", category=UserWarning, module="vertexai")
warnings.filterwarnings("ignore", category=DeprecationWarning)

import vertexai

from config.settings import (
    PROJECT_ID, REGION, METADATA_DB_PATH, FAISS_INDEX_PATH, 
    INPUT_DIR, OUTPUT_DIR, RULES_OUTPUT_DIR, UML_OUTPUT_DIR, UML_COMPONENTS_DIR, MODEL_NAME,
    PLANTUML_JAR_PATH, GRAPHVIZ_DOT_PATH
)
from parsers.vbnet_parser import VBNetParser
from vector_store.faiss_manager import FAISSManager
from llm_agents.rule_extractor import RuleExtractor
from llm_agents.uml_generator import UMLGenerator

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

    def process_workload(self):
        print("--- 3. Starting Parallel Map-Reduce LLM Extraction ---")

        with sqlite3.connect(METADATA_DB_PATH) as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            
            # Get chunks for rule extraction
            cursor.execute("SELECT * FROM CodeChunks WHERE processed_rules = 0 AND block_type NOT IN ('Class', 'Module', 'Structure')")
            rule_chunks = cursor.fetchall()
            
            # Get unique files for UML generation
            cursor.execute("SELECT DISTINCT file_path FROM CodeChunks WHERE processed_uml = 0")
            uml_files = cursor.fetchall()

        if not rule_chunks and not uml_files:
            print("No tasks pending.")
            return

        print(f"Total rule chunks to process: {len(rule_chunks)} (using {MAX_WORKERS} parallel workers)")

        # Phase 1: Rule & Manifest Extraction (Chunk Level)
        if rule_chunks:
            def make_rule_worker_fn(chunk_dict):
                re = RuleExtractor(MODEL_NAME)
                ug = UMLGenerator(MODEL_NAME)
                try:
                    context_chunks = self.faiss_manager.retrieve_context(chunk_dict['content'], k=3)
                    context_chunks = [c for c in context_chunks if c['id'] != chunk_dict['id']]
                except Exception:
                    context_chunks = []

                # Extract Rules
                try:
                    rules_df = re.process_chunk(chunk_dict, context_chunks)
                    if not rules_df.empty:
                        rules_path = os.path.join(RULES_OUTPUT_DIR, "consolidated_rules.csv")
                        with csv_lock:
                            if os.path.exists(rules_path):
                                rules_df.to_csv(rules_path, mode='a', header=False, index=False, sep='|')
                            else:
                                rules_df.to_csv(rules_path, index=False, sep='|')
                except Exception:
                    pass

                # Extract Structural Manifest (Map Step for UML)
                manifest = None
                try:
                    manifest = ug.extract_manifest(chunk_dict)
                except Exception as e:
                    print(f"Failed manifest extraction for {chunk_dict['name']}: {e}")

                with db_lock:
                    with sqlite3.connect(METADATA_DB_PATH) as conn:
                        if manifest is not None and manifest != "":
                            conn.cursor().execute(
                                "UPDATE CodeChunks SET processed_rules = 1, manifest = ? WHERE id = ?", 
                                (manifest, chunk_dict['id'])
                            )
                        else:
                            conn.cursor().execute(
                                "UPDATE CodeChunks SET processed_rules = 1 WHERE id = ?", 
                                (chunk_dict['id'],)
                            )
                        conn.commit()
                return chunk_dict['id']

            completed_rules = 0
            failed_rules = 0
            with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
                futures = {executor.submit(make_rule_worker_fn, dict(row)): dict(row)['id'] for row in rule_chunks}
                pbar = tqdm(total=len(rule_chunks), desc="Extracting Rules & Manifests")
                for future in as_completed(futures):
                    try:
                        future.result()
                        completed_rules += 1
                    except Exception:
                        failed_rules += 1
                    pbar.update(1)
                pbar.close()
            print(f"  ✓ Map Phase complete (✓ {completed_rules} | ✗ {failed_rules})")

        # Phase 2: UML Reduction (File Level)
        if uml_files:
            print(f"\nTotal files for UML reduction: {len(uml_files)} (using {MAX_WORKERS} parallel workers)")
            
            def make_uml_worker_fn(file_path):
                ug = UMLGenerator(MODEL_NAME)
                
                # Fetch all manifests for this file
                with sqlite3.connect(METADATA_DB_PATH) as conn:
                    cursor = conn.cursor()
                    cursor.execute("SELECT manifest FROM CodeChunks WHERE file_path = ?", (file_path,))
                    manifests = [row[0] for row in cursor.fetchall() if row[0]]
                
                src_file = os.path.splitext(os.path.basename(file_path))[0]
                
                if manifests:
                    try:
                        uml_str = ug.reduce_manifests_to_uml(src_file, manifests)
                        if uml_str:
                            puml_path = os.path.join(UML_COMPONENTS_DIR, f"{src_file}.puml")
                            with open(puml_path, "w", encoding="utf-8") as f:
                                f.write(ug._sanitize_puml(uml_str))
                    except Exception:
                        pass
                
                with db_lock:
                    with sqlite3.connect(METADATA_DB_PATH) as conn:
                        conn.cursor().execute("UPDATE CodeChunks SET processed_uml = 1 WHERE file_path = ?", (file_path,))
                        conn.commit()
                return src_file

            completed_uml = 0
            failed_uml = 0
            with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
                futures = {executor.submit(make_uml_worker_fn, row['file_path']): row['file_path'] for row in uml_files}
                pbar = tqdm(total=len(uml_files), desc="Generating Component UMLs")
                for future in as_completed(futures):
                    try:
                        future.result()
                        completed_uml += 1
                    except Exception:
                        failed_uml += 1
                    pbar.update(1)
                pbar.close()
            
            # Track generated files
            puml_count = len([f for f in os.listdir(UML_COMPONENTS_DIR) if f.endswith('.puml')])
            print(f"  ✓ UML Reduction complete (✓ {completed_uml} | ✗ {failed_uml})")
            print(f"  ✓ Total Component .puml files generated: {puml_count}")

        print("--- 4. Batch logic complete ---")

    def generate_end_to_end_flow(self):
        """Build an Architectural End-To-End component diagram by analyzing the Call Graph."""
        print("--- 5. Generating End-To-End Flow from Call Graph ---")

        try:
            with sqlite3.connect(METADATA_DB_PATH) as conn:
                df = pd.read_sql_query("SELECT caller_file, caller_name, caller_type, callee_name, call_line FROM CallGraph", conn)
                
            if df.empty:
                print("  ⚠ No call graph edges found.")
                return
                
            # Export call graph as CSV
            cg_path = os.path.join(RULES_OUTPUT_DIR, "call_graph.csv")
            df.to_csv(cg_path, index=False, sep='|')
            print(f"  ✓ Raw call graph exported: {len(df)} method edges → {cg_path}")

            # --- AGGREGATE TO FILE/COMPONENT LEVEL ---
            df['caller_comp'] = df['caller_file'].apply(lambda x: os.path.splitext(os.path.basename(x))[0])
            df['callee_comp'] = df['callee_name'].apply(lambda x: x.split('.')[0] if isinstance(x, str) and '.' in x else x)
            
            # Filter self-calls and get unique edges
            unique_edges = df[df['caller_comp'] != df['callee_comp']][['caller_comp', 'callee_comp']].drop_duplicates()
            
            print(f"  ✓ Aggregated to {len(unique_edges)} unique component-to-component edges.")

            # Send CSV to LLM to generate semantic flow diagram
            print("  Analyzing Call Graph to generate End-To-End flow...")
            ug = UMLGenerator(MODEL_NAME)
            
            if len(unique_edges) > 500:
                top_callees = unique_edges['callee_comp'].value_counts().nlargest(150).index
                csv_content = unique_edges[unique_edges['callee_comp'].isin(top_callees)].head(500).to_csv(index=False)
            else:
                csv_content = unique_edges.to_csv(index=False)
                
            puml_content = ug.generate_e2e_business_flow(csv_content)
            
            if puml_content:
                puml_path = os.path.join(UML_OUTPUT_DIR, "End_To_End_Flow.puml")
                with open(puml_path, "w", encoding="utf-8") as f:
                    f.write(puml_content)
                print(f"  ✓ Semantic End-To-End Architectural Flow diagram generated -> {puml_path}")
            else:
                print("  ✗ Failed to generate End-To-End Flow diagram from LLM.")

        except Exception as e:
            print(f"  Failed to generate End-To-End flow: {e}")
            import traceback; traceback.print_exc()

    def render_all_puml_locally(self):
        """Render all .puml files locally using plantuml.jar and Graphviz."""
        print("--- 6. Rendering PlantUML Diagrams Locally ---")
        if not os.path.exists(PLANTUML_JAR_PATH):
            print(f"  ⚠ PlantUML JAR not found at {PLANTUML_JAR_PATH}. Skipping local render.")
            return

        import subprocess
        
        # Render end-to-end flow
        e2e_puml = os.path.join(UML_OUTPUT_DIR, "End_To_End_Flow.puml")
        if os.path.exists(e2e_puml):
            cmd = ['java', '-jar', PLANTUML_JAR_PATH]
            if GRAPHVIZ_DOT_PATH and os.path.exists(GRAPHVIZ_DOT_PATH):
                cmd.extend(['-graphvizdot', GRAPHVIZ_DOT_PATH])
            cmd.append(e2e_puml)
            
            print(f"  Rendering End-To-End Flow...")
            subprocess.run(cmd, shell=True)
            print(f"  ✓ End-To-End Flow rendered as image.")

        # Render component diagrams
        print(f"  Rendering Component Diagrams...")
        cmd = ['java', '-jar', PLANTUML_JAR_PATH]
        if GRAPHVIZ_DOT_PATH and os.path.exists(GRAPHVIZ_DOT_PATH):
            cmd.extend(['-graphvizdot', GRAPHVIZ_DOT_PATH])
        cmd.append(os.path.join(UML_COMPONENTS_DIR, "*.puml"))
        
        subprocess.run(cmd, shell=True)
        
        # Track generated images
        png_count = len([f for f in os.listdir(UML_COMPONENTS_DIR) if f.endswith('.png') or f.endswith('.svg')])
        print(f"  ✓ Local rendering complete. Generated {png_count} Component images.")
        
        print(f"\nExtraction complete. Check '{OUTPUT_DIR}' for all outputs.")

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
        except Exception as e:
            print(f"  Fallback LLM macro also failed: {e}")


    def _generate_global_sequence_trace(self):
        call_graph_path = os.path.join(RULES_OUTPUT_DIR, "call_graph.csv")
        rules_path = os.path.join(RULES_OUTPUT_DIR, "consolidated_rules.csv")
        out_path = os.path.join(RULES_OUTPUT_DIR, "global_sequence_trace.csv")
        
        if not os.path.exists(call_graph_path) or not os.path.exists(rules_path):
            return

        cg = pd.read_csv(call_graph_path, sep='|')
        rules = pd.read_csv(rules_path, sep='|')
        
        # Build adjacency list
        graph = {}
        in_degree = {}
        for _, row in cg.iterrows():
            caller = row['caller_name']
            callee = row['callee_name']
            if caller not in graph: graph[caller] = []
            graph[caller].append(callee)
            if caller not in in_degree: in_degree[caller] = 0
            if callee not in in_degree: in_degree[callee] = 0
            in_degree[callee] += 1
            
        # Find entry points (in-degree == 0)
        entry_points = [node for node, degree in in_degree.items() if degree == 0]
        
        # DFS Trace
        visited = set()
        trace = []
        global_seq = 1
        
        def dfs(node, depth):
            nonlocal global_seq
            if node in visited: return
            visited.add(node)
            
            # Find rules for this node
            node_rules = rules[rules['Source_Method'] == node]
            if not node_rules.empty:
                for _, r in node_rules.sort_values('Sequence_Order').iterrows():
                    trace.append({
                        'Global_Sequence': global_seq,
                        'Call_Depth': depth,
                        'Method_Name': node,
                        'Rule_Name': r['Rule_Name'],
                        'Rule': r['Rule'],
                        'Actual_Code': r['Actual_Code'],
                        'Source_File': r['Source_File']
                    })
                    global_seq += 1
            else:
                trace.append({
                    'Global_Sequence': global_seq,
                    'Call_Depth': depth,
                    'Method_Name': node,
                    'Rule_Name': 'Call Flow Node',
                    'Rule': 'Invoked ' + node,
                    'Actual_Code': '',
                    'Source_File': ''
                })
                global_seq += 1
                
            if node in graph:
                for child in graph[node]:
                    dfs(child, depth + 1)
                    
        for ep in entry_points:
            dfs(ep, 0)
            
        pd.DataFrame(trace).to_csv(out_path, index=False, sep='|')
        print(f"  \u2714 Global sequence trace generated -> {out_path}")


if __name__ == "__main__":
    orchestrator = PipelineOrchestrator()
    orchestrator.run_parser_and_indexer()
    orchestrator.process_workload()
    orchestrator.generate_end_to_end_flow()
    orchestrator.render_all_puml_locally()
    print("--- 7. Generating Global Sequence Trace ---")
    orchestrator._generate_global_sequence_trace()
    print("Pipeline Execution Complete!")
