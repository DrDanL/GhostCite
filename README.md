# 👻 GhostCite

> **This project was vibe coded.** Built rapidly with AI-assisted development — expect rough edges, have fun, and contribute if you spot improvements!

GhostCite is a local-first web app that checks whether the references in an academic PDF or Word document are real. Upload a document, and GhostCite will extract every reference it finds, look each one up on [Crossref](https://www.crossref.org/), and tell you which ones appear to be genuine and which might be fabricated or "hallucinated".

## ✨ Features

- **PDF & DOCX upload** — drag-and-drop or select a file (up to 20 MB) through a simple web UI.
- **Document metadata and provenance signals** — reports embedded properties plus neutral, evidence-backed signals such as text below 2 pt, explicitly white text, GenAI product markers, C2PA/Content Credentials declarations, generator/conversion tools, tracked changes, comments, digital-signature parts, PDF revisions, Office add-ins, embedded attachments, and reference-manager indicators. PDF font-size checks use the effective rendered size after text/page transformations, preventing scaled `1 pt` storage values from producing false positives. Every detected sub-2-pt or white-text occurrence is listed with its location, formatting and full extracted text in both the web result and summary PDF. For PDFs it also surfaces trailer/document IDs, raw XMP edit history and custom fields, IPTC digital-source declarations, page geometry, annotations and reviewers, optional layers (including layers off by default), page-piece application data, form/signature metadata, viewer settings, document actions/scripts, and attachment properties. GhostCite never turns these signals into an AI verdict; human interpretation is always required.
- **Structure-aware reference extraction** — scores candidate reference sections by citation evidence rather than fixed document position; handles numbered, bracketed, bulleted, Harvard/APA, MLA-style, Unicode and institutional authors, large appendices, and Word references stored in tables or text boxes.
- **Crossref verification** — queries the Crossref API for each reference and classifies it as ✅ Verified, ⚠️ Partial Match, or ❌ Not Found.
- **Summary report** — download a standalone summary PDF listing every reference and its verification status.
- **Fully local & private** — runs entirely on your machine via Docker. Your documents are never sent to a third-party service (only individual reference strings are queried against Crossref's public API).

## 🚀 Quick Start

### Prerequisites

- [Docker](https://docs.docker.com/get-docker/) and [Docker Compose](https://docs.docker.com/compose/install/) installed on your machine.

### Run with one command

```bash
./start.sh
```

This builds the Docker image and starts the container. Once ready, open your browser at:

**👉 http://localhost:8000**

### How to use

1. Open **http://localhost:8000** in your browser.
2. Upload a PDF or DOCX file containing academic references.
3. Wait for GhostCite to extract and verify the references (this may take a moment depending on the number of references).
4. Review the results table showing each reference and its verification status.
5. Review the document metadata and any embedded reference-manager indicators shown above the verification results.
6. Download the **summary report** PDF, which includes the same metadata header and lists each reference and its status.

### Manual Docker commands

```bash
# Build and start in background
docker compose up --build -d

# View logs
docker compose logs -f

# Stop the app
docker compose down
```

## 🛠 Tech Stack

- **Python 3.12** with **Flask** and **Gunicorn**
- **pypdf** and **ReportLab** for PDF reading and report generation
- **python-docx** for DOCX support
- **Crossref REST API** for reference verification
- **Docker** for containerised deployment

## Contact

For questions or queries about GhostCite, contact Dan at [daniel.leightley@kcl.ac.uk](mailto:daniel.leightley@kcl.ac.uk).

## 📄 License

This project is provided as-is.
