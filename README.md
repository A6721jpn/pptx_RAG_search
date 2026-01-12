# PPTX RAG Search - Local PoC

Local Retrieval-Augmented Generation (RAG) system for PowerPoint files.
Runs entirely on your local machine using Windows COM for rendering and local embeddings.

## Features

- **Local Ingestion**: Scans a local directory for `.pptx` files.
- **Data Extraction**: Extracts text and speaker notes from slides.
- **Slide Rendering**: Converts slides to high-quality PNG images using PowerPoint.
- **Incremental Processing**: Only processes new or modified files.

## Prerequisites

- **OS**: Windows 10/11
- **PowerPoint**: Installed (Office 365 or 2019+)
- **Python**: 3.9+

## Setup

1. **Clone Repository**:
   ```bash
   git clone <repo_url>
   cd pptx_RAG_search
   ```

2. **Create Virtual Environment**:
   ```bash
   python -m venv .venv
   .venv\Scripts\activate
   ```

3. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Configuration**:
   Check `configs/local_config.yaml`. Default settings:
   ```yaml
   ingest:
     data_dir: "data/pptx"       # Put your .pptx files here
     output_dir: "data"          # Output directory for images/logs
     db_path: "data/processed_files.db"
   ```

## Usage

### Ingestion (Extract & Render)

1. **Prepare Data**:
   Place your PowerPoint files in `data/pptx` (or valid path in config).

2. **Run Pipeline**:
   ```bash
   python src/main.py --mode ingest
   ```

   This process will:
   - Identify new or modified `.pptx` files.
   - Extract text and speaker notes.
   - Render each slide as a PNG image in `data/rendered/`.
   - Update processing status in `data/processed_files.db`.

### Teams Bot Interface

The system includes a Microsoft Teams-compatible bot interface.

1. **Set Environment Variables** (in `.env`):
   ```
   MicrosoftAppId=""         # Leave empty for local emulator
   MicrosoftAppPassword=""   # Leave empty for local emulator
   OPENAI_API_KEY="sk-..."   # Required for LLM answer generation
   ```

2. **Run Bot Server**:
   ```bash
   python src/bot/app.py
   ```
   Server runs on `http://localhost:3978`.

3. **Test with Bot Framework Emulator**:
   - Open Bot Framework Emulator.
   - Connect to `http://localhost:3978/api/messages`.
   - Send a message in English. (Non-English inputs will be rejected).

### Check Logs

Logs are written to `data/logs/app.log` and printed to console.

## Project Structure

```
pptx_RAG_search/
├── configs/                # Configuration files
├── data/                   # Data directory (ignored by git)
│   ├── pptx/               # Input PPTX files
│   ├── rendered/           # Output PNG images
│   └── processed_files.db  # Status database
├── src/
│   ├── ingest/             # Ingestion modules (Extract/Render)
│   ├── utils/              # Utilities
│   └── main.py             # Entry point
└── README.md
```
