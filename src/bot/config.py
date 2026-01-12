import os
from dotenv import load_dotenv

load_dotenv()

class DefaultConfig:
    """ Bot Configuration """

    PORT = 3978
    APP_ID = os.environ.get("MicrosoftAppId", "")
    APP_PASSWORD = os.environ.get("MicrosoftAppPassword", "")
    
    # RAG Config
    OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
    QDRANT_URL = os.environ.get("QDRANT_URL", "http://localhost:6333")
