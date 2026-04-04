# VB.NET Business Rule Extractor & UML Generator

A production-ready, parallel Map-Reduce pipeline that extracts business rules and generates end-to-end PlantUML flow diagrams from large-scale (~100MB) VB.NET/VB6 codebases using Google Cloud Vertex AI (Gemini Flash).

---

## Prerequisites

1. **Python 3.10+**
2. **Google Cloud Account** with Vertex AI API enabled
3. Authenticate locally:
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

4. **Configure your GCP Project:**
   - Edit `config/settings.py` and set your `PROJECT_ID`:
     ```python
     PROJECT_ID = os.getenv("GCP_PROJECT_ID", "your-project-id-here")
     ```
   - Or export as an environment variable:
     ```powershell
     $env:GCP_PROJECT_ID="your-project-id"
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

**To process a specific subfolder only**, edit `config/settings.py`:
```python
INPUT_DIR = os.path.join(BASE_DIR, "input_code", "COM Server", "busServer")
```

### 2. Run the Pipeline

```bash
python pipeline_orchestrator.py
```

The pipeline runs in 5 stages:
1. **Parse** — Recursively scans all VB files, extracts Classes/Modules/Subs/Functions, splits large blocks into chunks
2. **Index** — Builds FAISS vector index for context retrieval (RAG)
3. **Extract** — Runs 50 parallel LLM workers for rule extraction + individual UML generation
4. **Complete** — Saves all rules and component diagrams
5. **End-to-End Flow** — Builds deterministic call graph diagram from real code analysis

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

---

## Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| `429 Resource exhausted` | Too many parallel API calls | Reduce `MAX_WORKERS` in `pipeline_orchestrator.py` (try 20-30). Built-in retry with backoff handles occasional 429s automatically. |
| `input token count exceeds 20000` | Code chunk too large for embeddings | Reduce `MAX_CHUNK_LINES` in `config/settings.py` (try 200) |
| `PlantUML 509 Bandwidth Exceeded` | Too many PNG requests to public server | Wait 30s and re-run. SVG fallback is automatic. For production, use a local PlantUML install. |
| `Error tokenizing data` | LLM returned malformed CSV | Non-fatal — chunk is skipped. Rules from other chunks are unaffected. |
| `ModuleNotFoundError` | Missing Python package | Run `pip install -r requirements.txt` |