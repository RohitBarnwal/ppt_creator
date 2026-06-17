# Paytm & PPSL Organization Chart Creator Utility

A modern, standalone utility that dynamically reverse-engineers and generates beautifully styled, hierarchical organization chart PowerPoint decks from standard Excel databases. It can be run either as a local Desktop Application or hosted as a FastAPI Web Application (locally or containerized).

## 🚀 Features
- **FastAPI Web Service:** Modern, async API that processes uploaded Excel files and returns styled PPTX presentations.
- **Visual Color Gradient Coding:** Automatically applies hierarchical executive color pallets (VSS Navy $\rightarrow$ HOD Medium Blue $\rightarrow$ Manager Light Blue $\rightarrow$ IC Very Light Blue) matching professional design standards.
- **Dual-Level Visualization Limit:** Restricts each sub-org sheet to 2 levels (HOD + direct reports) to maintain a highly readable layout.
- **Rollup Summaries:** Automatically detects and displays total rollup counts (recursive child counts) inside manager cards (e.g., `DR- 13`) rather than cluttering slides.
- **Programmatic Elbow Routing:** Utilizes native PowerPoint Elbow Connectors that auto-snap to standard boundary ports.
- **Multi-Sheet Support:** Automatically concatenates divided worksheets (e.g., `Leads` and `Team`) into a single, unified tree database.

---

## 🌐 FastAPI Web Application (Recommended)

### 1. Local Setup & Execution
Install python dependencies (including FastAPI and testing frameworks):
```bash
pip install -r requirements.txt
```

Run the FastAPI backend:
```bash
uvicorn main:app --reload --host 127.0.0.1 --port 8000
```
Open [http://localhost:8000](http://localhost:8000) in your browser to access the drag-and-drop web UI.

### 2. Run via Docker (Containerized)
Build the Docker image:
```bash
docker build -t org-chart-utility .
```

Run the container:
```bash
docker run -p 8000:8000 org-chart-utility
```
The application will be accessible at [http://localhost:8000](http://localhost:8000).

### 3. Running Automated Tests
Run unit and integration tests using pytest:
```bash
pytest test_main.py
```

---

## 💻 Standalone Desktop Application (Fallback)
If you prefer to run the standalone customtkinter desktop interface:
1. Ensure `customtkinter` is installed (included in `requirements.txt`).
2. Run the desktop application:
   ```bash
   python org_chart_app.py
   ```

