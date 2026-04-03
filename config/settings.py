import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Google Cloud Vertex AI settings
PROJECT_ID = os.getenv("GCP_PROJECT_ID", "PROJECT_ID")
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

# Paths
INPUT_DIR = os.path.join(BASE_DIR, "input_code")
RULES_OUTPUT_DIR = os.path.join(OUTPUT_DIR, "rules")
UML_OUTPUT_DIR = os.path.join(OUTPUT_DIR, "uml")

# Ensure directories exist
os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(RULES_OUTPUT_DIR, exist_ok=True)
os.makedirs(UML_OUTPUT_DIR, exist_ok=True)
UML_COMPONENTS_DIR = os.path.join(UML_OUTPUT_DIR, "components")
os.makedirs(UML_COMPONENTS_DIR, exist_ok=True)
