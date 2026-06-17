# Background & Motivation
The current Organization Chart Generator is a Python desktop application built with `customtkinter`. To improve accessibility, enable remote usage, and prepare for future cloud deployment, the application is being migrated to a FastAPI-based web application.

# Scope & Impact
- **Backend:** Replace `customtkinter` with **FastAPI** to serve HTTP requests.
- **Frontend:** Implement a simple HTML/JS frontend to mimic the current desktop UI (file upload, generation status, and file download).
- **Core Logic:** Retain the existing `build_org_structure.py` for the core processing, modifying it slightly to support temporary file paths handling concurrent web requests.
- **Deployment:** Containerize the application using **Docker** for seamless local testing and future web deployment.

# Proposed Solution
1. **FastAPI Application (`main.py`):**
   - Implement an endpoint (e.g., `POST /generate`) that accepts a file upload (`UploadFile`).
   - Save the uploaded file to a temporary directory.
   - Invoke `build_org_structure.build_org_presentation()`.
   - Return the generated `.pptx` file using FastAPI's `FileResponse`.
   - Use `BackgroundTasks` to clean up temporary files after the response is sent.

2. **Web Frontend (`static/index.html`):**
   - Create a clean, simple web interface.
   - Include a file input for the Excel file.
   - Provide visual feedback (processing spinner/status) during generation.
   - Automatically trigger a file download once the API responds.

3. **Dockerization (`Dockerfile` & `.dockerignore`):**
   - Create a standard Python Dockerfile using `uvicorn` to run the FastAPI app.
   - Expose port 8000.

# Phased Implementation Plan
- **Phase 1: Setup & API Construction**
  - Add `fastapi`, `uvicorn`, and `python-multipart` to `requirements.txt`.
  - Create `main.py` and implement the `POST /generate` endpoint.
- **Phase 2: Frontend Integration**
  - Create the `static/index.html` interface.
  - Configure FastAPI to serve static files.
- **Phase 3: Dockerization**
  - Write `Dockerfile` and `.dockerignore`.
  - Add instructions to `README.md` for running locally via Docker and direct Uvicorn.

# Verification
- Start the server locally via `uvicorn`.
- Upload a sample Excel file via the HTML frontend.
- Verify that the UI shows processing status and successfully downloads the PPTX.
- Verify the generated PPTX is valid.
- Build and run the Docker image to ensure it works containerized.

# Migration & Rollback
- The existing `org_chart_app.py` (Desktop app) will remain in the repository as a fallback during the transition. Once the web app is fully verified and deployed, it can be optionally deprecated.
