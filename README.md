# VB.NET Business Rule Extractor & UML Generator

This is a production-ready, Map-Reduce architecture pipeline designed to process large-scale (~100MB) VB.NET codebases without data loss. It extracts business logic line-by-line into comprehensive sequential CSV rules and generates full PlantUML activity and system sequence diagrams natively using Google Cloud Vertex AI (Gemini Flash).

## Prerequisites
1. **Python 3.10+**
2. **Google Cloud Account** with Vertex AI API enabled.
3. Authenticate locally with your Google Cloud Credentials:
   ```bash
   gcloud auth application-default login
   ```

## Setup & Installation
1. Install all dependencies from `requirements.txt`:
   ```bash
   pip install -r requirements.txt
   ```
2. Configuration bindings are defined in `config/settings.py`. Ensure your `GCP_PROJECT_ID` and `GCP_REGION` bindings match your Google Cloud project (by default, it binds to `PROJECT_ID` and `us-central1`).
   - If you need to change your Project ID, you can either update it directly in `config/settings.py` or export it in your terminal as `$env:GCP_PROJECT_ID="your_project_id"`.

## Executing the Pipeline

1. **Place Your Codebase**
   Dump your entire `.vb` codebase (e.g., 100MB of files) directly into the `input_code/` folder:
   ```plaintext
   VbNet_Extractor/
   ├── input_code/
   │   ├── ApplicationFile1.vb
   │   ├── ModuleFile2.vb
   │   └── ...
   ```

2. **Run the Orchestrator**
   Execute the central orchestrator to begin Map-Reduce processing, RAG indexing, Rule extraction, and UML rendering linearly.
   ```bash
   python pipeline_orchestrator.py
   ```

### Resuming from Interruptions
The system tracks all completions inside an embedded SQLite state database (`output/metadata.db`). If the process crashes, the network disconnects, or the computer shuts down, simply run `python pipeline_orchestrator.py` again. It will skip all processed chunks and guarantee 100% data preservation seamlessly.

## Generated Outputs
All artifacts are routed cleanly into the `output/` directory:
- **`output/rules/consolidated_rules.csv`**: Contains the complete serialized tracking dataframe of all business logic rules globally, properly formatted with the actual sequence order, rule extraction type, and raw VB.NET code mapped.
- **`output/uml/End_To_End_Flow.png`** (and `.puml`): The master sequence architecture mapped end-to-end for the entire system based on all module interactions.
- **`output/uml/components/`**: A subdirectory populated with the individual method-by-method computational UML flow arrays `.puml` and generated `.png` renders for granular visual code review.