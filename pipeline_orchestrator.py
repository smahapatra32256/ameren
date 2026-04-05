import csv
import hashlib
import os
import sqlite3
import subprocess
import sys
import threading
import warnings
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd
from tqdm import tqdm

# Force utf-8 encoding for stdout on Windows to prevent UnicodeEncodeError
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

warnings.filterwarnings("ignore", category=UserWarning, module="vertexai")
warnings.filterwarnings("ignore", category=DeprecationWarning)

import vertexai

from config.settings import (
    FAISS_INDEX_PATH,
    GRAPHVIZ_DOT_PATH,
    INPUT_DIR,
    METADATA_DB_PATH,
    MODEL_NAME,
    OUTPUT_DIR,
    PLANTUML_JAR_PATH,
    PROJECT_ID,
    REGION,
    RULES_OUTPUT_DIR,
    UML_COMPONENTS_DIR,
    UML_OUTPUT_DIR,
)
from llm_agents.rule_extractor import RuleExtractionError, RuleExtractor
from llm_agents.uml_generator import UMLGenerationError, UMLGenerator
from parsers.vbnet_parser import VBNetParser
from vector_store.faiss_manager import FAISSManager

# Thread-safe locks
csv_lock = threading.Lock()
db_lock = threading.Lock()
agent_local = threading.local()

# Max parallel LLM workers
MAX_WORKERS = 50


class PipelineOrchestrator:
    def __init__(self):
        vertexai.init(project=PROJECT_ID, location=REGION)
        self.parser = VBNetParser(METADATA_DB_PATH)
        self.faiss_manager = FAISSManager(METADATA_DB_PATH, FAISS_INDEX_PATH, PROJECT_ID, REGION)

    def _db_connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(METADATA_DB_PATH)
        conn.execute("PRAGMA journal_mode=MEMORY")
        conn.execute("PRAGMA temp_store=MEMORY")
        return conn

    def _get_rule_extractor(self) -> RuleExtractor:
        if not hasattr(agent_local, "rule_extractor"):
            agent_local.rule_extractor = RuleExtractor(MODEL_NAME)
        return agent_local.rule_extractor

    def _get_uml_generator(self) -> UMLGenerator:
        if not hasattr(agent_local, "uml_generator"):
            agent_local.uml_generator = UMLGenerator(MODEL_NAME)
        return agent_local.uml_generator

    def _append_failure_log(self, log_path: str, header: List[str], row_values: List[str]) -> None:
        os.makedirs(os.path.dirname(log_path), exist_ok=True)
        with csv_lock:
            file_exists = os.path.exists(log_path)
            with open(log_path, "a", encoding="utf-8", newline="") as f:
                writer = csv.writer(f, delimiter="|")
                if not file_exists:
                    writer.writerow(header)
                writer.writerow(row_values)

    def _safe_component_filename(self, file_path: str) -> str:
        base = os.path.splitext(os.path.basename(file_path))[0]
        suffix = hashlib.md5(file_path.encode("utf-8")).hexdigest()[:8]
        safe_base = "".join(c if c.isalnum() or c in ("_", "-") else "_" for c in base)
        return f"{safe_base}_{suffix}"

    def run_parser_and_indexer(self):
        print(f"--- 1. Parsing input code from {INPUT_DIR} ---")
        self.parser.parse_directory(INPUT_DIR)
        print("--- 2. Building FAISS index ---")
        self.faiss_manager.build_index()

    def process_workload(self):
        print("--- 3. Starting Parallel Map-Reduce LLM Extraction ---")

        with self._db_connect() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()

            # Process method/property chunks that still need rules and/or manifests.
            cursor.execute(
                """
                SELECT *
                FROM CodeChunks
                WHERE (processed_rules = 0 OR processed_manifest = 0)
                  AND block_type NOT IN ('Class', 'Module', 'Structure')
                """
            )
            rule_chunks = cursor.fetchall()

            # Only reduce UML when all chunk manifests for that file are ready.
            cursor.execute(
                """
                SELECT DISTINCT c.file_path
                FROM CodeChunks c
                WHERE c.processed_uml = 0
                  AND NOT EXISTS (
                      SELECT 1
                      FROM CodeChunks pending
                      WHERE pending.file_path = c.file_path
                        AND pending.block_type NOT IN ('Class', 'Module', 'Structure')
                        AND pending.processed_manifest = 0
                  )
                """
            )
            uml_files = [r["file_path"] for r in cursor.fetchall()]

        if not rule_chunks and not uml_files:
            print("No tasks pending.")
            return

        print(f"Rule/Manifest chunks pending: {len(rule_chunks)} (workers={MAX_WORKERS})")

        # Phase 1: Rule + Manifest extraction at chunk level
        if rule_chunks:

            def map_worker(chunk_dict: Dict) -> Dict[str, int]:
                rules_done = 0
                rules_failed = 0
                manifests_done = 0
                manifests_failed = 0
                manifest_payload: Optional[str] = None

                rule_extractor = self._get_rule_extractor()
                uml_generator = self._get_uml_generator()

                try:
                    context_chunks = self.faiss_manager.retrieve_context(chunk_dict["content"], k=3)
                    context_chunks = [c for c in context_chunks if c.get("id") != chunk_dict.get("id")]
                except Exception:
                    context_chunks = []

                # Rules extraction
                if int(chunk_dict.get("processed_rules", 0)) == 0:
                    try:
                        rules_df = rule_extractor.process_chunk(chunk_dict, context_chunks)
                        if not rules_df.empty:
                            rules_path = os.path.join(RULES_OUTPUT_DIR, "consolidated_rules.csv")
                            with csv_lock:
                                if os.path.exists(rules_path):
                                    rules_df.to_csv(rules_path, mode="a", header=False, index=False, sep="|")
                                else:
                                    rules_df.to_csv(rules_path, index=False, sep="|")
                        rules_done = 1
                    except RuleExtractionError as e:
                        rules_failed = 1
                        self._append_failure_log(
                            os.path.join(RULES_OUTPUT_DIR, "failed_rule_chunks.csv"),
                            ["chunk_id", "file_path", "block_type", "name", "error"],
                            [chunk_dict["id"], chunk_dict["file_path"], chunk_dict["block_type"], chunk_dict["name"], str(e)],
                        )
                    except Exception as e:
                        rules_failed = 1
                        self._append_failure_log(
                            os.path.join(RULES_OUTPUT_DIR, "failed_rule_chunks.csv"),
                            ["chunk_id", "file_path", "block_type", "name", "error"],
                            [chunk_dict["id"], chunk_dict["file_path"], chunk_dict["block_type"], chunk_dict["name"], str(e)],
                        )

                # Manifest extraction
                if int(chunk_dict.get("processed_manifest", 0)) == 0:
                    try:
                        manifest_payload = uml_generator.extract_manifest(chunk_dict)
                        manifests_done = 1
                    except UMLGenerationError as e:
                        manifests_failed = 1
                        self._append_failure_log(
                            os.path.join(RULES_OUTPUT_DIR, "failed_manifest_chunks.csv"),
                            ["chunk_id", "file_path", "block_type", "name", "error"],
                            [chunk_dict["id"], chunk_dict["file_path"], chunk_dict["block_type"], chunk_dict["name"], str(e)],
                        )
                    except Exception as e:
                        manifests_failed = 1
                        self._append_failure_log(
                            os.path.join(RULES_OUTPUT_DIR, "failed_manifest_chunks.csv"),
                            ["chunk_id", "file_path", "block_type", "name", "error"],
                            [chunk_dict["id"], chunk_dict["file_path"], chunk_dict["block_type"], chunk_dict["name"], str(e)],
                        )

                # Persist only successful stage flags.
                if rules_done or manifests_done:
                    with db_lock:
                        with self._db_connect() as conn:
                            cursor = conn.cursor()
                            set_clauses: List[str] = []
                            params: List[object] = []
                            if rules_done:
                                set_clauses.append("processed_rules = 1")
                            if manifests_done:
                                set_clauses.append("processed_manifest = 1")
                                set_clauses.append("manifest = ?")
                                params.append(manifest_payload)
                            params.append(chunk_dict["id"])
                            cursor.execute(
                                f"UPDATE CodeChunks SET {', '.join(set_clauses)} WHERE id = ?",
                                params,
                            )
                            conn.commit()

                return {
                    "rules_done": rules_done,
                    "rules_failed": rules_failed,
                    "manifests_done": manifests_done,
                    "manifests_failed": manifests_failed,
                }

            stats = Counter()
            with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
                futures = [executor.submit(map_worker, dict(row)) for row in rule_chunks]
                pbar = tqdm(total=len(rule_chunks), desc="Extracting Rules + Manifests")
                for future in as_completed(futures):
                    try:
                        result = future.result()
                        stats.update(result)
                    except Exception as e:
                        stats["rules_failed"] += 1
                        stats["manifests_failed"] += 1
                        self._append_failure_log(
                            os.path.join(RULES_OUTPUT_DIR, "failed_map_futures.csv"),
                            ["error"],
                            [str(e)],
                        )
                    pbar.update(1)
                pbar.close()

            print(
                "  Map complete: "
                f"rules_done={stats['rules_done']}, rules_failed={stats['rules_failed']}, "
                f"manifests_done={stats['manifests_done']}, manifests_failed={stats['manifests_failed']}"
            )

        # Refresh UML-ready files after map phase updates.
        with self._db_connect() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.cursor()
            cursor.execute(
                """
                SELECT DISTINCT c.file_path
                FROM CodeChunks c
                WHERE c.processed_uml = 0
                  AND NOT EXISTS (
                      SELECT 1
                      FROM CodeChunks pending
                      WHERE pending.file_path = c.file_path
                        AND pending.block_type NOT IN ('Class', 'Module', 'Structure')
                        AND pending.processed_manifest = 0
                  )
                """
            )
            uml_files = [r["file_path"] for r in cursor.fetchall()]

        # Phase 2: UML file-level reduction
        if uml_files:
            print(f"UML reductions pending: {len(uml_files)} (workers={MAX_WORKERS})")

            def reduce_worker(file_path: str) -> Tuple[bool, str]:
                uml_generator = self._get_uml_generator()
                src_file = os.path.splitext(os.path.basename(file_path))[0]
                safe_name = self._safe_component_filename(file_path)

                with self._db_connect() as conn:
                    cursor = conn.cursor()
                    cursor.execute(
                        """
                        SELECT manifest
                        FROM CodeChunks
                        WHERE file_path = ?
                          AND processed_manifest = 1
                          AND manifest IS NOT NULL
                          AND TRIM(manifest) <> ''
                        """,
                        (file_path,),
                    )
                    manifests = [row[0] for row in cursor.fetchall() if row[0]]

                if not manifests:
                    self._append_failure_log(
                        os.path.join(RULES_OUTPUT_DIR, "failed_uml_files.csv"),
                        ["file_path", "error"],
                        [file_path, "No manifests available for reduction"],
                    )
                    return False, src_file

                try:
                    uml_str = uml_generator.reduce_manifests_to_uml(src_file, manifests)
                    puml_path = os.path.join(UML_COMPONENTS_DIR, f"{safe_name}.puml")
                    with open(puml_path, "w", encoding="utf-8") as f:
                        f.write(uml_str)

                    with db_lock:
                        with self._db_connect() as conn:
                            conn.execute("UPDATE CodeChunks SET processed_uml = 1 WHERE file_path = ?", (file_path,))
                            conn.commit()
                    return True, src_file
                except Exception as e:
                    self._append_failure_log(
                        os.path.join(RULES_OUTPUT_DIR, "failed_uml_files.csv"),
                        ["file_path", "error"],
                        [file_path, str(e)],
                    )
                    return False, src_file

            completed_uml = 0
            failed_uml = 0
            with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
                futures = [executor.submit(reduce_worker, file_path) for file_path in uml_files]
                pbar = tqdm(total=len(uml_files), desc="Generating File-Level UML")
                for future in as_completed(futures):
                    ok, _ = future.result()
                    if ok:
                        completed_uml += 1
                    else:
                        failed_uml += 1
                    pbar.update(1)
                pbar.close()

            puml_count = len([f for f in os.listdir(UML_COMPONENTS_DIR) if f.endswith(".puml")])
            print(f"  UML reduction complete: success={completed_uml}, failed={failed_uml}, files={puml_count}")

        print("--- 4. Batch logic complete ---")

    def _alias_map(self, nodes: List[str]) -> Dict[str, str]:
        alias_map: Dict[str, str] = {}
        used: Set[str] = set()
        for node in nodes:
            base = "".join(ch if ch.isalnum() else "_" for ch in node)
            if not base:
                base = "Node"
            candidate = base
            suffix = 1
            while candidate in used:
                suffix += 1
                candidate = f"{base}_{suffix}"
            alias_map[node] = candidate
            used.add(candidate)
        return alias_map

    def _build_component_flow_puml(
        self,
        edges_df: pd.DataFrame,
        source_col: str,
        target_col: str,
        title: str,
    ) -> str:
        if edges_df.empty:
            return ""

        nodes = sorted(set(edges_df[source_col].astype(str)) | set(edges_df[target_col].astype(str)))
        aliases = self._alias_map(nodes)

        lines: List[str] = [
            "@startuml",
            f"title {title}",
            "left to right direction",
            "skinparam componentStyle rectangle",
        ]

        for node in nodes:
            lines.append(f'component "{node}" as {aliases[node]}')

        lines.append("")
        for _, row in edges_df.iterrows():
            src = str(row[source_col])
            dst = str(row[target_col])
            if src == dst:
                continue
            lines.append(f"{aliases[src]} --> {aliases[dst]}")

        lines.append("@enduml")
        return "\n".join(lines)

    def _classify_component(self, name: str) -> str:
        n = (name or "").lower()
        if n.startswith("bus"):
            return "Business"
        if n.startswith("db"):
            return "DataAccess"
        if n.startswith("util"):
            return "Utility"
        if n.startswith("bo"):
            return "BusinessObject"
        if n.startswith("rpt"):
            return "Reporting"
        if n.startswith("cd") or n.startswith("cbs") or n.startswith("cu"):
            return "ServiceClass"
        return "ExternalOrShared"

    def generate_end_to_end_flow(self):
        """Generate deterministic end-to-end file/component flow diagrams from the call graph."""
        print("--- 5. Generating End-To-End Flow from Call Graph ---")

        try:
            with self._db_connect() as conn:
                call_graph_df = pd.read_sql_query(
                    "SELECT caller_file, caller_name, caller_type, callee_name, call_line FROM CallGraph",
                    conn,
                )
                chunks_df = pd.read_sql_query(
                    "SELECT file_path, name, block_type FROM CodeChunks",
                    conn,
                )

            if call_graph_df.empty:
                print("  No call graph edges found.")
                return

            # Full-fidelity call graph export (no truncation)
            call_graph_path = os.path.join(RULES_OUTPUT_DIR, "call_graph.csv")
            call_graph_df.to_csv(call_graph_path, index=False, sep="|")
            print(f"  Call graph exported: {len(call_graph_df)} rows -> {call_graph_path}")

            # Build method->file resolver for internal file-level mapping
            method_chunks = chunks_df[~chunks_df["block_type"].isin(["Class", "Module", "Structure"])].copy()
            method_chunks["file_component"] = method_chunks["file_path"].apply(lambda p: os.path.splitext(os.path.basename(p))[0])
            method_to_files: Dict[str, Set[str]] = defaultdict(set)
            for _, row in method_chunks.iterrows():
                method_to_files[str(row["name"])].add(str(row["file_component"]))

            file_edges: List[Dict[str, str]] = []
            for _, row in call_graph_df.iterrows():
                caller_component = os.path.splitext(os.path.basename(str(row["caller_file"])))[0]
                callee_name = str(row["callee_name"])

                if "." in callee_name:
                    callee_component = callee_name.split(".", 1)[0]
                    relation_type = "object_call"
                else:
                    candidate_files = sorted(method_to_files.get(callee_name, set()))
                    if len(candidate_files) == 1:
                        callee_component = candidate_files[0]
                        relation_type = "internal_method"
                    elif len(candidate_files) > 1:
                        callee_component = f"AMBIG::{callee_name}"
                        relation_type = "ambiguous_method"
                    else:
                        callee_component = callee_name
                        relation_type = "external_or_builtin"

                file_edges.append(
                    {
                        "caller_component": caller_component,
                        "callee_component": callee_component,
                        "relation_type": relation_type,
                    }
                )

            file_edges_df = pd.DataFrame(file_edges)
            file_unique_edges = (
                file_edges_df[["caller_component", "callee_component"]]
                .dropna()
                .drop_duplicates()
                .reset_index(drop=True)
            )

            file_edges_path = os.path.join(RULES_OUTPUT_DIR, "file_level_edges.csv")
            file_edges_df.to_csv(file_edges_path, index=False, sep="|")

            file_puml = self._build_component_flow_puml(
                file_unique_edges,
                source_col="caller_component",
                target_col="callee_component",
                title="End To End File Flow",
            )
            if file_puml:
                e2e_path = os.path.join(UML_OUTPUT_DIR, "End_To_End_Flow.puml")
                with open(e2e_path, "w", encoding="utf-8") as f:
                    f.write(file_puml)
                print(f"  File-level E2E diagram generated: {e2e_path}")

            layer_edges_df = file_unique_edges.copy()
            layer_edges_df["caller_layer"] = layer_edges_df["caller_component"].apply(self._classify_component)
            layer_edges_df["callee_layer"] = layer_edges_df["callee_component"].apply(self._classify_component)
            layer_unique = layer_edges_df[["caller_layer", "callee_layer"]].drop_duplicates()

            component_edges_path = os.path.join(RULES_OUTPUT_DIR, "component_level_edges.csv")
            layer_edges_df.to_csv(component_edges_path, index=False, sep="|")

            component_puml = self._build_component_flow_puml(
                layer_unique,
                source_col="caller_layer",
                target_col="callee_layer",
                title="Component Level Flow",
            )
            if component_puml:
                comp_path = os.path.join(UML_OUTPUT_DIR, "Component_Level_Flow.puml")
                with open(comp_path, "w", encoding="utf-8") as f:
                    f.write(component_puml)
                print(f"  Component-level diagram generated: {comp_path}")

            print(f"  File-level edges: {len(file_unique_edges)} | Layer-level edges: {len(layer_unique)}")

        except Exception as e:
            print(f"  Failed to generate End-To-End flow: {e}")

    def render_all_puml_locally(self):
        """Render all .puml files locally using plantuml.jar and Graphviz."""
        print("--- 6. Rendering PlantUML Diagrams Locally ---")
        if not os.path.exists(PLANTUML_JAR_PATH):
            print(f"  PlantUML JAR not found at {PLANTUML_JAR_PATH}. Skipping local render.")
            return

        puml_files: List[str] = []
        for fixed_name in ["End_To_End_Flow.puml", "Component_Level_Flow.puml"]:
            p = os.path.join(UML_OUTPUT_DIR, fixed_name)
            if os.path.exists(p):
                puml_files.append(p)

        if os.path.exists(UML_COMPONENTS_DIR):
            puml_files.extend(
                [os.path.join(UML_COMPONENTS_DIR, f) for f in os.listdir(UML_COMPONENTS_DIR) if f.endswith(".puml")]
            )

        if not puml_files:
            print("  No .puml files found to render.")
            return

        base_cmd = ["java", "-jar", PLANTUML_JAR_PATH]
        if GRAPHVIZ_DOT_PATH and os.path.exists(GRAPHVIZ_DOT_PATH):
            base_cmd.extend(["-graphvizdot", GRAPHVIZ_DOT_PATH])

        rendered = 0
        failed = 0
        for puml_path in puml_files:
            cmd = base_cmd + [puml_path]
            proc = subprocess.run(cmd, capture_output=True, text=True)
            if proc.returncode == 0:
                rendered += 1
            else:
                failed += 1
                self._append_failure_log(
                    os.path.join(RULES_OUTPUT_DIR, "failed_render_jobs.csv"),
                    ["puml_path", "return_code", "stderr"],
                    [puml_path, str(proc.returncode), (proc.stderr or "").strip()],
                )

        png_count = 0
        if os.path.exists(UML_COMPONENTS_DIR):
            png_count = len([f for f in os.listdir(UML_COMPONENTS_DIR) if f.endswith(".png") or f.endswith(".svg")])

        print(f"  Render jobs: success={rendered}, failed={failed}")
        print(f"  Component images present: {png_count}")
        print(f"\nExtraction complete. Check '{OUTPUT_DIR}' for outputs.")

    def _generate_global_sequence_trace(self):
        call_graph_path = os.path.join(RULES_OUTPUT_DIR, "call_graph.csv")
        rules_path = os.path.join(RULES_OUTPUT_DIR, "consolidated_rules.csv")
        out_path = os.path.join(RULES_OUTPUT_DIR, "global_sequence_trace.csv")

        if not os.path.exists(call_graph_path) or not os.path.exists(rules_path):
            print("  Skipping global sequence trace (missing call graph or rules CSV).")
            return

        cg = pd.read_csv(call_graph_path, sep="|", dtype=str, keep_default_na=False)
        rules = pd.read_csv(rules_path, sep="|", dtype=str, keep_default_na=False)

        if cg.empty:
            print("  Skipping global sequence trace (call graph is empty).")
            return

        graph: Dict[str, List[str]] = defaultdict(list)
        in_degree: Dict[str, int] = defaultdict(int)
        nodes: Set[str] = set()

        for _, row in cg.iterrows():
            caller = row.get("caller_name", "")
            callee = row.get("callee_name", "")
            if not caller or not callee:
                continue
            graph[caller].append(callee)
            in_degree[callee] += 1
            _ = in_degree[caller]
            nodes.add(caller)
            nodes.add(callee)

        entry_points = sorted([node for node in nodes if in_degree[node] == 0])
        if not entry_points:
            # Cyclic graph fallback: start from top fan-out callers
            caller_counts = Counter(cg["caller_name"].tolist())
            entry_points = [name for name, _ in caller_counts.most_common(10)]

        trace = []
        global_seq = 1
        edge_visits: Dict[Tuple[str, str], int] = defaultdict(int)
        max_edge_visits = 2

        def sorted_rules_for_node(node: str) -> pd.DataFrame:
            node_rules = rules[rules["Source_Method"] == node].copy()
            if node_rules.empty:
                return node_rules
            node_rules["_seq"] = pd.to_numeric(node_rules["Sequence_Order"], errors="coerce").fillna(999999)
            node_rules = node_rules.sort_values(["_seq", "Rule_Name"], ascending=[True, True])
            return node_rules

        def dfs(node: str, depth: int):
            nonlocal global_seq

            node_rules = sorted_rules_for_node(node)
            if not node_rules.empty:
                for _, r in node_rules.iterrows():
                    trace.append(
                        {
                            "Global_Sequence": global_seq,
                            "Call_Depth": depth,
                            "Method_Name": node,
                            "Rule_Name": r.get("Rule_Name", ""),
                            "Rule": r.get("Rule", ""),
                            "Actual_Code": r.get("Actual_Code", ""),
                            "Source_File": r.get("Source_File", ""),
                        }
                    )
                    global_seq += 1
            else:
                trace.append(
                    {
                        "Global_Sequence": global_seq,
                        "Call_Depth": depth,
                        "Method_Name": node,
                        "Rule_Name": "Call Flow Node",
                        "Rule": f"Invoked {node}",
                        "Actual_Code": "",
                        "Source_File": "",
                    }
                )
                global_seq += 1

            for child in graph.get(node, []):
                edge = (node, child)
                if edge_visits[edge] >= max_edge_visits:
                    continue
                edge_visits[edge] += 1
                dfs(child, depth + 1)

        for ep in entry_points:
            dfs(ep, 0)

        pd.DataFrame(trace).to_csv(out_path, index=False, sep="|")
        print(f"  Global sequence trace generated -> {out_path}")


if __name__ == "__main__":
    orchestrator = PipelineOrchestrator()
    orchestrator.run_parser_and_indexer()
    orchestrator.process_workload()
    orchestrator.generate_end_to_end_flow()
    orchestrator.render_all_puml_locally()
    print("--- 7. Generating Global Sequence Trace ---")
    orchestrator._generate_global_sequence_trace()
    print("Pipeline Execution Complete!")
