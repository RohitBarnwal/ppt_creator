import os
import shutil
import tempfile
import logging
from fastapi import FastAPI, UploadFile, File, BackgroundTasks, HTTPException
from fastapi.responses import HTMLResponse, FileResponse
import build_org_structure

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("org-chart-api")

app = FastAPI(title="Paytm Org Chart Creator API")

def cleanup_files(*filepaths):
    """Clean up files after response is sent."""
    for filepath in filepaths:
        try:
            if filepath and os.path.exists(filepath):
                os.remove(filepath)
                logger.info(f"Successfully deleted temp file: {filepath}")
        except Exception as e:
            logger.error(f"Error cleaning up file {filepath}: {str(e)}")

@app.post("/generate")
async def generate_org_chart(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...)
):
    # Validate file extension
    if not (file.filename.endswith(".xlsx") or file.filename.endswith(".xls")):
        raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are supported.")
        
    # Create temporary files
    try:
        temp_excel = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file.filename)[1])
        temp_excel_path = temp_excel.name
        
        # Write uploaded content
        shutil.copyfileobj(file.file, temp_excel)
        temp_excel.close() # Close so build_org_presentation can read it
    except Exception as e:
        logger.error(f"Failed to save uploaded file: {str(e)}")
        raise HTTPException(status_code=500, detail="Failed to save uploaded file.")
        
    temp_pptx_path = temp_excel_path.replace(".xlsx", ".pptx").replace(".xls", ".pptx")
    
    # Run the core generation logic
    try:
        # Register cleanup tasks
        background_tasks.add_task(cleanup_files, temp_excel_path, temp_pptx_path)
        
        success = build_org_structure.build_org_presentation(temp_excel_path, temp_pptx_path)
        if not success:
            raise HTTPException(status_code=400, detail="Generation failed. Please verify that the Excel sheet is correctly formatted.")
            
        if not os.path.exists(temp_pptx_path):
            raise HTTPException(status_code=500, detail="Output presentation file was not created.")
            
        # Propose custom response filename based on input filename
        output_filename = "Automated_Org_Structure_" + file.filename.replace(".xlsx", ".pptx").replace(".xls", ".pptx")
        
        return FileResponse(
            path=temp_pptx_path,
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Unexpected error during presentation build: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Internal Server Error: {str(e)}")

@app.get("/", response_class=HTMLResponse)
async def serve_index():
    index_path = os.path.join(os.path.dirname(__file__), "static", "index.html")
    if os.path.exists(index_path):
        with open(index_path, "r", encoding="utf-8") as f:
            return f.read()
    else:
        return "<html><body><h3>index.html not found</h3></body></html>"
