# Pipeline Architecture & System Design

This document details the internal architecture, design patterns, and workflows of the VB.NET Extraction and UML Generation pipeline. It is intended for developers maintaining or extending the extraction toolset.

## System Overview

The pipeline is a parallelized, multi-stage application written in Python. It acts as an orchestrator that bridges static code analysis (regex parsing, AST-like block detection) with advanced generative AI (Google Vertex AI / Gemini Flash) to extract deep semantic meaning from legacy VB.NET applications. 

The architecture follows a distributed **Map-Reduce** paradigm, leveraging local SQLite for state management and FAISS for vector-based Retrieval-Augmented Generation (RAG).

## Core Components

### 1. The Orchestrator (`pipeline_orchestrator.py`)
Acts as the central nervous system of the pipeline.
* **Concurrency:** Uses Python's `ThreadPoolExecutor` to manage up to 50 concurrent worker threads.
* **State Management:** Connects to a local SQLite database to track the processing state of individual code chunks, allowing for resumability and preventing duplicate work if the pipeline fails or is interrupted.
* **Flow Control:** Sequences the pipeline into strict phases: Parsing → Vector Indexing → Parallel Map-Reduce (Rules & UML Manifests) → End-to-End Analysis → Local Rendering → Global Trace Generation.

### 2. Static Analysis & Chunking (`parsers/vbnet_parser.py`)
Responsible for reading raw source files and breaking them down into manageable, context-rich chunks for the LLM.
* **Heuristic Block Detection:** Uses regular expressions to identify logical block boundaries (`Class`, `Module`, `Sub`, `Function`, `Property`). It builds a stack to track nested structures and extracts the content between the declaration and the corresponding `End` statement.
* **Chunk Splitting:** If a logical block exceeds the maximum allowed length (to prevent exceeding LLM context windows), it dynamically splits the chunk, enforcing a 20-line overlap to maintain code continuity.
* **Call Graph Extraction:** Performs a static scan over the chunks to identify method invocations (`Call X`, `Me.X`, `Obj.X`). It writes caller-to-callee edges to the SQLite database, which is later critical for building architectural diagrams and the global chronological trace.

### 3. Vector Database & RAG (`vector_store/faiss_manager.py`)
Provides context awareness to the LLM.
* **Embedding:** Uses `VertexAIEmbeddings` (`text-embedding-004`) to convert raw VB.NET code chunks into high-dimensional vector representations. It includes a dynamic batching mechanism that calculates token/character counts on the fly to maximize throughput without exceeding API token limits.
* **Storage:** Leverages Facebook AI Similarity Search (FAISS) for lightning-fast similarity lookups. 
* **Context Provision:** When the LLM analyzes a specific chunk, the FAISS manager queries the index for the Top-K most similar code chunks elsewhere in the codebase, injecting them into the LLM prompt as contextual background.

### 4. AI Agents (`llm_agents/`)
Wrappers around the Vertex AI SDK, heavily customized with strict prompting, safety settings, and exponential backoff retry logic to handle rate limiting (HTTP 429).

#### A. Rule Extractor (`rule_extractor.py`)
* **Purpose:** Identifies business logic, conditions, and operations within a specific code chunk.
* **Format Enforcement:** Uses few-shot prompting and explicit instructions to force the LLM to output valid CSV data wrapped in markdown. It includes post-processing logic to strip markdown, sanitize internal pipes (`|`), and clean up newlines.
* **Integration:** Outputs are written to a consolidated `consolidated_rules.csv`.

#### B. UML Generator (`uml_generator.py`)
Operates using a Map-Reduce strategy to ensure architectural accuracy without blowing out the token context window.
* **Map Phase (`extract_manifest`):** Asks the LLM to read a single method chunk and output raw PlantUML Sequence Diagram syntax documenting *only* what that method does. Uses a highly restricted token configuration (`safe_config = 4096`) to prevent silent payload truncation.
* **Reduce Phase (`reduce_manifests_to_uml`):** Collects all manifests generated for a single `.cls` or `.bas` file and asks the LLM to stitch them together into a unified Component Sequence Diagram.
* **End-to-End Phase:** Takes the entire extracted static Call Graph (as a CSV string) and asks the LLM to generate a high-level System Component Diagram, enforcing directional data flow arrows (`-->`).

## Data Flow & Execution Phases

1. **Initialization (`_init_db`):** The local SQLite `metadata.db` is created with tables for `CodeChunks` and `CallGraph`.
2. **Phase 1: Parse & Index:** The VB.NET parser reads the target directory, populates the DB, and `FAISSManager` embeds all chunks.
3. **Phase 2: Parallel Chunk Processing (Map):** The orchestrator spins up 50 threads. For each chunk:
   * It asks FAISS for related context.
   * It calls the `RuleExtractor` to generate CSV rules.
   * It calls the `UMLGenerator` to extract a localized PlantUML manifest.
   * Marks the chunk as `processed_rules = 1` in SQLite.
4. **Phase 3: File-Level Reduction (Reduce):** For every unique file, the orchestrator retrieves all associated manifests from SQLite and passes them to the `UMLGenerator` to synthesize a complete file-level `.puml` Sequence Diagram.
5. **Phase 4: Global Architectural Analysis:** The orchestrator reads the `CallGraph` table, exports it as a CSV, and passes it to the `UMLGenerator` to map the End-to-End flow.
6. **Phase 5: Local Rendering:** The orchestrator invokes a local Java subprocess running `plantuml.jar` (and Graphviz) to convert all generated `.puml` files into final `.png` images, bypassing public PlantUML API bottlenecks.
7. **Phase 6: Trace Construction:** The pipeline runs a Depth-First Search (DFS) algorithm over the Call Graph edges, cross-referencing them with the extracted business rules to build `global_sequence_trace.csv`. This provides a flattened, chronological timeline of application execution.

## Resilience Mechanisms
* **API Rate Limiting:** All Vertex AI calls are wrapped in robust retry blocks utilizing exponential backoff (`wait = 5 * (2 ** attempt)`).
* **Token Cutoffs:** The manifest extraction phase actively checks `if not response_obj.candidates[0].content.parts:` to detect and recover from silent `MAX_TOKENS` finish reasons provided by the Vertex API.
* **Platform Encoding:** Explicitly forces `sys.stdout.reconfigure(encoding='utf-8')` to prevent `cp1252` encoding crashes on Windows consoles when printing unicode shapes (like `→`).