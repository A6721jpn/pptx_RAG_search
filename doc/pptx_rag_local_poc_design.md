# PPTX Mechanical Design Guide RAG — Local-Only PoC Design (Windows COM)

**Author:** PK / ChatGPT  
**Date:** 2026-01-11 (Asia/Tokyo)  
**Status:** Draft v1 (PoC-ready)

---

## 1. Objectives

Build a **fully local** Proof of Concept (PoC) for searching large volumes of **mechanical design guide** PowerPoint decks where a slide typically contains:

- A **figure/diagram** (often the key evidence)
- Explanatory **English text** (bullets, captions, notes)

PoC must support:

- **English text queries**
- **Image queries** (optional in PoC; supported by design)
- Display to users:
  - **Slide thumbnail** (whole-slide is OK)
  - **Text excerpt** (from slide text and/or speaker notes)
- If search confidence is low or multiple candidates are plausible:
  - Ask **clarifying questions** and refine search interactively

Constraints:

- **Windows** environment
- PPTX slide rendering via **Windows COM automation (PowerPoint)**
- **No discrete GPU** (CPU-only inference is acceptable; model choices optimized accordingly)
- **No external API** in PoC (later: answer generation via external LLM is planned)

---

## 2. Non-goals (PoC)

- Perfect slide-region cropping (thumbnail is sufficient)
- Complex OCR (we rely primarily on PPT text extraction and slide rendering)
- Full conversational “assistant” UX (CLI is acceptable; minimal web UI can be added later)
- Enterprise-grade security hardening (documented as future work)

---

## 3. High-Level Architecture

### 3.1 Data flow

1. **Ingest PPTX**  
2. **Extract text** (shapes + speaker notes)  
3. **Render slide thumbnails** (PowerPoint COM → PNG)  
4. **Compute embeddings** (local):
   - Text embeddings (semantic search)
   - Image-text joint embeddings (CLIP) for cross-modal search
5. **Store in vector DB** (local Qdrant) with payload for retrieval/UI
6. **Query**:
   - Compute query embedding(s)
   - Search both indexes (text, clip)
   - **Fuse** results (RRF)
7. **Decision**:
   - If confident: show results
   - If ambiguous/low-confidence: ask clarifying question and re-run search with refinement

### 3.2 Retrieval model

- Two vectors per slide (or per chunk):
  - `text_vec`: semantic vector for slide text + notes
  - `clip_vec`: image vector of slide thumbnail (OpenCLIP)  
- Use **rank fusion** so either modality can surface relevant slides.

---

## 4. Technology Stack (Local)

### 4.1 Slide parsing & rendering

- **python-pptx** for extracting:
  - Text from shapes (text boxes, placeholders)
  - Speaker notes (if present)
- **PowerPoint COM automation (win32com)** for:
  - Rendering slides to PNG thumbnails (stable, accurate rendering)

### 4.2 Embeddings (CPU)

Text embedding (local, CPU-friendly):

- Baseline: `sentence-transformers` with a strong English retrieval model.  
  Examples (choose one based on speed/quality tradeoff):
  - `intfloat/e5-base-v2` (good quality, moderate speed)
  - `BAAI/bge-base-en-v1.5` (good quality, moderate speed)
  - If CPU is tight: smaller “mini” variants

Image-text embedding (local, CPU-friendly):

- `open_clip_torch` with a small CLIP model (CPU acceptable):
  - e.g., ViT-B/32 equivalents are common choices for PoC
  - This supports:
    - text query → image space retrieval
    - image query → image space retrieval

### 4.3 Vector storage (local)

- **Qdrant** local instance (Docker recommended)
- Use **named vectors**:
  - `text_vec`
  - `clip_vec`

---

## 5. Data Model

### 5.1 Index unit

**Default PoC unit:** 1 slide = 1 item

Later experiments can introduce chunking. The design supports multiple chunks per slide by adding `chunk_id`.

### 5.2 Qdrant point schema

**Point ID** (string or integer) — recommended deterministic:
- `"{doc_id}:{slide_no}:{chunk_id}"`

**Vectors:**
- `text_vec` : float32[d_text]
- `clip_vec` : float32[d_clip]

**Payload (minimum for UI):**
- `doc_id` (hash or stable name)
- `file_path` (local path)
- `file_name`
- `slide_no` (1-based)
- `title` (optional; inferred)
- `text_raw` (shape text)
- `notes_raw` (speaker notes)
- `thumb_path` (PNG thumbnail)
- `chunk_id` (default `"0"`)
- `chunk_mode` (`"slide_full"`, `"text_blocks"`, etc.)
- `created_at`, `updated_at`

**Security note:** For PoC local-only, payload can include `text_raw`. If you later externalize the vector DB, reconsider what belongs in payload.

---

## 6. File & Folder Layout

Example working directory structure:

```
rag_pptx_poc/
  data/
    pptx/
      <decks...>.pptx
    rendered/
      <doc_id>/
        slide_0001.png
        slide_0002.png
        ...
    extracted/
      <doc_id>.jsonl
  index/
    qdrant_storage/   (if using local volume)
  src/
    ingest/
      pptx_extract.py
      pptx_render_com.py
      build_index.py
    query/
      search.py
      clarify.py
    utils/
      hashing.py
      text_clean.py
      rrf.py
      config.py
  configs/
    default.yaml
  README.md
```

---

## 7. Ingestion Pipeline Details

### 7.1 Document ID (doc_id)

Compute a stable `doc_id` to avoid reindexing everything when paths change:

- `doc_id = sha1(file_bytes)[:12]` or `sha1(normalized_path + mtime)`

Recommendation for PoC: use **file bytes hash** (stable, content-derived).

### 7.2 Text extraction rules

For each slide:

1. Extract text from all shapes:
   - Text frames (including placeholders)
   - Tables if accessible via python-pptx (extract cell text)
2. Extract speaker notes (if any):
   - Notes often contain “why” and are very useful for retrieval
3. Clean text:
   - Normalize whitespace
   - Remove repeated headers/footers if they appear on every slide
   - Preserve bullet ordering (helps comprehension)

Output per slide (JSONL record):

```json
{
  "doc_id": "...",
  "file_path": "...",
  "slide_no": 12,
  "text_raw": "...",
  "notes_raw": "...",
  "thumb_path": "data/rendered/.../slide_0012.png",
  "chunk_id": "0",
  "chunk_mode": "slide_full"
}
```

### 7.3 Slide rendering via PowerPoint COM

Use PowerPoint COM to export slides to PNG. Conceptually:

- Open PPTX in PowerPoint
- Export all slides to a folder (PowerPoint supports `Export` on Presentation or Slide)
- Ensure deterministic filenames (`slide_0001.png`, etc.)
- Close PowerPoint properly to avoid orphaned processes

**Key requirements:**
- Run on Windows with Office installed
- COM is not thread-safe; do rendering in a single process, single thread

### 7.4 Embedding computation

For each record:

- `text_input = text_raw + "\n\n" + notes_raw`
- Compute `text_vec = TextEmbedder.embed(text_input)`
- Load `thumb_path`, compute `clip_vec = ClipEmbedder.embed_image(image)`

Batching:
- Text: batch embeds for speed
- Images: batch if possible; CPU may be slower; keep batch size small

---

## 8. Query Pipeline

### 8.1 Query types

- **Text query** (English)
- **Image query** (optional for PoC; supported by same ClipEmbedder)

### 8.2 Retrieval steps

Given a text query `q`:

1. `q_text_vec = TextEmbedder.embed(q)`
2. `q_clip_text_vec = ClipEmbedder.embed_text(q)`  *(OpenCLIP supports text embeddings)*

Perform two searches:

- `S_text = search(collection, vector="text_vec", query=q_text_vec, top_k=K_text)`
- `S_clip = search(collection, vector="clip_vec", query=q_clip_text_vec, top_k=K_clip)`

### 8.3 Rank fusion (RRF)

Use **Reciprocal Rank Fusion** to merge rankings without score calibration:

For each candidate `d`:

\[
RRF(d)=\sum_{r\in\{text,clip\}} \frac{1}{k + rank_r(d)}
\]

- Typical `k = 60` (tunable)
- Result: unified ranked list `S_fused`

### 8.4 Output formatting

For the top N results, show:

- Slide thumbnail (path or rendered in UI)
- File name + slide number
- A short excerpt:
  - First ~200–400 characters from `text_raw + notes_raw`
  - Optionally highlight query terms (simple string match PoC)

---

## 9. Interactive Clarification (Reverse Questioning)

### 9.1 Why clarification is needed

In design guide retrieval:
- Many slides reuse similar vocabulary (“tolerance”, “clearance”, “stiffness”)
- Multiple slides may be relevant but for different subsystems or failure modes
- Some queries are under-specified (“hinge issue”)

### 9.2 Trigger conditions

Let `scores` be similarity scores from the fused ranking or from the best modality (implementation-dependent).

Recommended triggers:

**Low relevance (no good match):**
- `top1_score < T_low`
- Or `avg(topK_scores) < T_avg`

**Ambiguous (multiple plausible):**
- `top1_score - top2_score < Δ`
- Or diversity indicates multiple intents (optional clustering)

Because RRF is rank-based, you may also compute modality-specific triggers:

- Low relevance if both modalities have low top scores
- Ambiguous if both modalities produce different top candidates

**Practical PoC approach:**
- Use Qdrant similarity score from `text_vec` search as primary confidence indicator for text queries
- Use `clip_vec` score for image queries
- Use Δ and thresholds tuned on a small labeled evaluation set

### 9.3 Clarification strategies

**Strategy A: Choice question (best for ambiguity)**  
Show 3–5 candidates and ask:

- “Which is closer to what you mean? (1/2/3/4/5)”
- Optionally allow “none of these”

**Strategy B: Facet question (best for low relevance)**  
Ask 1–2 questions that add discriminative constraints:

Suggested facets for mechanical design guides:
- Subsystem: hinge / chassis / keyboard / touchpad / thermal / speaker / connector
- Objective: stiffness / strength / tolerance / assembly / reliability / acoustics
- Process: injection molding / stamping / die-cast / fastening / adhesive
- Failure mode: crack / creep / loosen / wear / noise / deformation

In PoC (no external LLM), provide fixed options to avoid free-text ambiguity.

### 9.4 How clarification refines search

**Option 1 — Metadata filter (best long-term)**  
If you create lightweight tags per slide, clarification can apply filters:

- `payload.subsystem == "hinge"`

PoC tagging can start simple:
- Keyword rules over `text_raw + notes_raw` (fast, transparent)

**Option 2 — Query expansion (fast to implement)**  
Append chosen facet terms to the query:

- `q' = q + " " + facet_terms`

**Option 3 — Relevance feedback (Rocchio)**
If user chooses candidate `d*` as “closest”, update query vector:

\[
q_{new}=\alpha q_{old} + \beta \cdot vec(d^*)
\]

For PoC: use only positive feedback (no negatives). Works well to steer results.

---

## 10. Chunking Experiments (Planned in PoC)

Even though PoC starts with 1 slide = 1 chunk, design supports alternates.

### 10.1 Chunk modes

- `slide_full` (baseline)
- `text_blocks`:
  - chunk per text box / per bullet group
  - still reference `thumb_path` at slide-level
- `notes_only`:
  - chunk only speaker notes (good if notes are rich)
- `hybrid_text`:
  - (text boxes) + (notes) but weighted differently (e.g., duplicate notes to emphasize)

### 10.2 Evaluation approach

Create a small “gold set”:

- 30–50 real queries from designers
- For each query, label 1–3 correct slides

Metrics:
- Recall@5, Recall@10
- MRR (Mean Reciprocal Rank)

Use these to tune:
- embedding model choice
- chunk mode
- RRF `k`
- clarification thresholds (`T_low`, `Δ`)

---

## 11. Performance & Scaling Notes (CPU-only)

Expected bottlenecks:

1. Slide rendering via COM (I/O bound, PowerPoint overhead)
2. Image embedding computation (CPU heavy)

Mitigations:
- Cache rendered PNGs by `doc_id` and skip if already present
- Incremental indexing (only new/changed decks)
- Use smaller OpenCLIP model for PoC
- Limit image embedding to thumbnails (e.g., 1024 px max dimension)

---

## 12. Implementation Plan (Milestones)

### Milestone 1 — Baseline (1–2 days)
- PPTX text extraction (python-pptx)
- COM thumbnail export
- Text-only embedding + Qdrant indexing
- Text query → top results with excerpt + thumbnail path

### Milestone 2 — Cross-modal retrieval (1–2 days)
- OpenCLIP image embeddings for thumbnails
- Text query also searched in CLIP space
- RRF fusion

### Milestone 3 — Clarification loop (1–2 days)
- Confidence/ambiguity triggers
- Choice-based clarification
- Query expansion or positive-feedback vector update

### Milestone 4 — Chunking experiments (ongoing)
- Add `text_blocks` chunking
- Compare metrics on labeled set

---

## 13. Future Extension: External Answer Generation

Once local retrieval is robust, external LLM can generate final narrative answers.

Recommended pattern:
- Keep retrieval and excerpt extraction **local**
- Send only **minimal excerpts** to external LLM
- Display:
  - LLM narrative answer
  - Local evidence thumbnails + excerpts

This preserves “PPTX never leaves local machine” while enabling high-quality explanations.

---

## 14. Open Questions / Decisions Needed (for implementation)

1. **Office version** availability (PowerPoint COM behavior can vary)
2. Desired maximum index size (number of decks / slides)
3. Whether speaker notes are consistently present
4. Whether to pre-tag slides with simple dictionaries for faceted clarification

---

## 15. Appendix — Pseudocode Snippets

### 15.1 RRF fusion

```python
def rrf_fuse(rankings, k=60):
    # rankings: list of lists of doc_ids in rank order
    scores = {}
    for r in rankings:
        for i, doc_id in enumerate(r, start=1):
            scores[doc_id] = scores.get(doc_id, 0.0) + 1.0 / (k + i)
    return sorted(scores.items(), key=lambda x: x[1], reverse=True)
```

### 15.2 Clarification trigger (simple)

```python
def need_clarification(top_scores, T_low=0.20, delta=0.02):
    if not top_scores:
        return "facet"
    if top_scores[0] < T_low:
        return "facet"
    if len(top_scores) >= 2 and (top_scores[0] - top_scores[1]) < delta:
        return "choice"
    return None
```

---

## 16. Appendix — Minimal PoC CLI UX

- `index build --pptx_dir data/pptx --out data/extracted --render data/rendered`
- `search "hinge clearance tolerance stackup" --top 5`
- If ambiguous:
  - show candidates 1..5
  - prompt: `Select [1-5], or 0 for none:`
- If low relevance:
  - prompt: `Select subsystem: [hinge, chassis, keyboard, thermal, other]`

---

End of document.
