import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Google Cloud Vertex AI settings
PROJECT_ID = os.getenv("GCP_PROJECT_ID", "quiet-grail-431414-e7")
REGION = os.getenv("GCP_REGION", "us-central1")
MODEL_NAME = "gemini-2.5-flash"
EMBEDDING_MODEL_NAME = "text-embedding-004"

# FAISS Configuration
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
FAISS_INDEX_PATH = os.path.join(OUTPUT_DIR, "faiss_index.bin")
METADATA_DB_PATH = os.path.join(OUTPUT_DIR, "metadata.db")

# Chunking and generation limits
BATCH_SIZE = 100
MAX_OUTPUT_TOKENS = 8192
TEMPERATURE = 0.0
MAX_EMBEDDING_CHARS = 40000   # ~10k tokens safe limit for text-embedding-004 (max 20k tokens)
MAX_CHUNK_LINES = 300         # Split code blocks larger than this into sub-chunks

# Paths
INPUT_DIR = os.path.join(BASE_DIR, "input_code", "COM Server", "dbServer")
RULES_OUTPUT_DIR = os.path.join(OUTPUT_DIR, "rules")
UML_OUTPUT_DIR = os.path.join(OUTPUT_DIR, "uml")

# Local Rendering
PLANTUML_JAR_PATH = os.path.join(BASE_DIR, "plantuml.jar")
GRAPHVIZ_DOT_PATH = "C:\\Program Files\\Graphviz\\bin\\dot.exe"

# Ensure directories exist
os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(RULES_OUTPUT_DIR, exist_ok=True)
os.makedirs(UML_OUTPUT_DIR, exist_ok=True)
UML_COMPONENTS_DIR = os.path.join(UML_OUTPUT_DIR, "components")
os.makedirs(UML_COMPONENTS_DIR, exist_ok=True)
