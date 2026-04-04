# VB.NET Business Rule Extractor & UML Generator

A production-ready, parallel Map-Reduce pipeline that extracts business rules and generates end-to-end PlantUML flow diagrams from large-scale (~100MB) VB.NET/VB6 codebases using Google Cloud Vertex AI (Gemini Flash).

---

## Prerequisites

1. **Python 3.10+**
2. **Google Cloud Account** with Vertex AI API enabled
3. **Java Runtime Environment (JRE)** (Required for local PlantUML rendering)
   - Download from Adoptium: [OpenJDK 25 (x64 Windows)](https://adoptium.net/download?link=https%3A%2F%2Fgithub.com%2Fadoptium%2Ftemurin25-binaries%2Freleases%2Fdownload%2Fjdk-25.0.2%252B10%2FOpenJDK25U-jdk_x64_windows_hotspot_25.0.2_10.msi&vendor=Adoptium)
4. **Graphviz** (Required for complex UML layout generation)
   - *Windows:* `winget install graphviz` OR use the direct installer: [Graphviz 14.1.4 Installer](https://gitlab.com/api/v4/projects/4207231/packages/generic/graphviz-releases/14.1.4/windows_10_cmake_Release_graphviz-install-14.1.4-win64.exe)
   - **Important:** Make sure to select the option to add Graphviz to your system PATH during installation.
5. Authenticate locally:
   ```bash
   gcloud auth application-default login
   ```

---

## Setup & Installation

1. **Clone the repo and create a virtual environment:**
   ```bash
   git clone https://github.com/smahapatra32256/ameren.git
   cd ameren
   python -m venv venv
   ```

2. **Activate the virtual environment:**
   ```powershell
   # Windows PowerShell (run this ONCE if venv activation is blocked)
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

   # Then activate
   .\venv\Scripts\activate
   ```
   ```bash
   # Linux/Mac
   source venv/bin/activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Download PlantUML:**
   - Download the JAR file: [plantuml-1.2026.2.jar](https://github.com/plantuml/plantuml/releases/download/v1.2026.2/plantuml-1.2026.2.jar)
   - Rename it to `plantuml.jar` and place it directly in the root directory of this project (`VbNet_Extractor/`).

5. **Configure your Project & Paths:**
   - Edit `config/settings.py`:
     - **GCP Project:** Set your `PROJECT_ID` or export as environment variable `$env:GCP_PROJECT_ID="your-project-id"`.
     - **Graphviz Path:** If Graphviz was not added to your PATH (e.g. via winget), explicitly set its location:
       ```python
       GRAPHVIZ_DOT_PATH = "C:\\Program Files\\Graphviz\\bin\\dot.exe"
       ```
     - **Target Codebase:** By default, it scans everything in `input_code`. To narrow it down, edit `INPUT_DIR`:
       ```python
       INPUT_DIR = os.path.join(BASE_DIR, "input_code", "SpecificFolder")
       ```

---

## Google Cloud Services Required

| Service | Purpose | How to Enable |
|---|---|---|
| **Vertex AI API** | LLM calls (Gemini Flash) for rule extraction & UML generation | `gcloud services enable aiplatform.googleapis.com` |
| **Vertex AI Embeddings** | Text embeddings for FAISS-based context retrieval | Included with Vertex AI API |

> **Note:** Ensure your GCP project has billing enabled and sufficient quota for Vertex AI API calls. The pipeline uses `gemini-2.5-flash` and `text-embedding-004`.

---

## How to Run

### 1. Place Your Codebase

Drop your VB.NET/VB6 source files (`.vb`, `.vbs`, `.bas`, `.cls`, `.frm`) into the `input_code/` folder. **All subdirectories are scanned recursively.**

```
VbNet_Extractor/
├── input_code/
│   ├── Module1/
│   │   ├── Service1.cls
│   │   └── Helper.bas
│   ├── Module2/
│   │   └── DataAccess.vb
│   └── ...
```

### 2. Run the Pipeline

```bash
python pipeline_orchestrator.py
```

The pipeline runs in 6 stages:
1. **Parse** — Recursively scans all VB files, extracts Classes/Modules/Subs/Functions, splits large blocks into chunks
2. **Index** — Builds FAISS vector index for context retrieval (RAG)
3. **Extract** — Runs 50 parallel LLM workers for rule extraction + individual UML `.puml` generation
4. **Complete** — Saves all rules and component diagrams
5. **End-to-End Flow** — Builds deterministic call graph diagram from real code analysis
6. **Local Rendering** — Automatically runs `plantuml.jar` with Graphviz to convert all `.puml` files to `.png` images locally.

### 3. View Outputs

```
output/
├── metadata.db                    # SQLite tracking DB (resume support)
├── faiss_index.bin                # FAISS vector index
├── faiss_index.bin.map            # FAISS ID mapping
├── rules/
│   ├── consolidated_rules.csv     # All business rules (pipe-delimited)
│   └── call_graph.csv             # Every caller→callee relationship
├── uml/
│   ├── End_To_End_Flow.puml       # System-wide flow from REAL call graph
│   ├── End_To_End_Flow.png        # Rendered diagram
│   └── components/                # Individual method/class UML diagrams
│       ├── MethodName_abc123.puml
│       └── MethodName_abc123.png
```

---

## Cleanup & Reset

### Full Cleanup (removes ALL generated outputs — run before re-processing)

```bash
python -c "import os, shutil; [os.remove(f) for f in ['output/metadata.db','output/faiss_index.bin','output/faiss_index.bin.map'] if os.path.exists(f)]; [shutil.rmtree(d) for d in ['output/rules','output/uml'] if os.path.exists(d)]; print('Full cleanup done!')"
```

### Quick Cleanup (keeps UML components, resets rules and tracking DB only)

```bash
python -c "import os, glob; [os.remove(f) for f in glob.glob('output/metadata.db') + glob.glob('output/faiss_index.*') + glob.glob('output/rules/*') if os.path.exists(f)]; print('Quick cleanup done!')"
```

> **Important:** Always clean the database (`metadata.db`) when you change the input directory. Otherwise old chunks accumulate and inflate the count.

---

## Resuming Interrupted Runs

The pipeline tracks processing state in `output/metadata.db`. If the process crashes, the network drops, or the computer shuts down:

```bash
python pipeline_orchestrator.py
```

It will **automatically skip all already-processed chunks** and resume from where it left off. Zero data loss.

---

## Configuration Reference (`config/settings.py`)

| Setting | Default | Description |
|---|---|---|
| `PROJECT_ID` | `quiet-grail-431414-e7` | Your GCP Project ID |
| `REGION` | `us-central1` | Vertex AI region |
| `MODEL_NAME` | `gemini-2.5-flash` | LLM model for extraction |
| `EMBEDDING_MODEL_NAME` | `text-embedding-004` | Embedding model for FAISS |
| `MAX_OUTPUT_TOKENS` | `8192` | Max LLM output tokens |
| `MAX_CHUNK_LINES` | `300` | Split code blocks larger than this |
| `MAX_WORKERS` | `50` | Parallel LLM threads (in `pipeline_orchestrator.py`) |
| `PLANTUML_JAR_PATH` | `BASE_DIR/plantuml.jar` | Path to the local PlantUML JAR file |
| `GRAPHVIZ_DOT_PATH` | `C:\Program Files\...` | Path to the Graphviz dot.exe file |

---

## Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| `429 Resource exhausted` | Too many parallel API calls | Reduce `MAX_WORKERS` in `pipeline_orchestrator.py` (try 20-30). Built-in retry with backoff handles occasional 429s automatically. |
| `input token count exceeds 20000` | Code chunk too large for embeddings | Reduce `MAX_CHUNK_LINES` in `config/settings.py` (try 200) |
| `Dot Executable Not Found` | PlantUML cannot locate Graphviz | Ensure Graphviz is installed and update `GRAPHVIZ_DOT_PATH` in `config/settings.py` to point to `dot.exe`. |
| `Error tokenizing data` | LLM returned malformed CSV | Non-fatal — chunk is skipped. Rules from other chunks are unaffected. |
| `ModuleNotFoundError` | Missing Python package | Run `pip install -r requirements.txt` |