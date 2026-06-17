import io
import os
import tempfile
import pytest
import pandas as pd
from fastapi.testclient import TestClient
from main import app

client = TestClient(app)

def test_read_root():
    """Verify that the home page is served successfully with HTML content."""
    response = client.get("/")
    assert response.status_code == 200
    assert "text/html" in response.headers["content-type"]
    assert "Org Chart Creator Utility" in response.text

def test_generate_invalid_file_type():
    """Verify that files other than Excel are rejected with a 400 Bad Request."""
    files = {"file": ("test.txt", b"invalid file content", "text/plain")}
    response = client.post("/generate", files=files)
    assert response.status_code == 400
    assert "Only Excel files (.xlsx, .xls) are supported." in response.json()["detail"]

def test_generate_valid_org_chart():
    """Verify that a valid Excel file produces a downloadable PPTX presentation."""
    # Create a minimal valid organization dataset using pandas
    data = {
        "Employee Name": [
            "Vijay Shekhar Sharma", 
            "Deependra Singh Rathore", 
            "Gagan", 
            "HOD Compliance", 
            "IC Reportee"
        ],
        "Manager Name": [
            "None", 
            "Vijay Shekhar Sharma", 
            "Deependra Singh Rathore", 
            "Gagan", 
            "HOD Compliance"
        ],
        "Designation": [
            "MD & CEO", 
            "COO - Tech, Product and Ops", 
            "Director", 
            "Compliance Head", 
            "Analyst"
        ],
        "Sub Department": [
            "CEO Office", 
            "Operations", 
            "PG Tech", 
            "Compliance", 
            "Compliance"
        ]
    }
    df = pd.DataFrame(data)
    
    # Save to a temporary Excel file in bytes
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Leads", index=False)
    excel_buffer.seek(0)
    
    # Send post request to generate
    files = {"file": ("test_employees.xlsx", excel_buffer, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
    response = client.post("/generate", files=files)
    
    assert response.status_code == 200
    assert response.headers["content-type"] == "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    assert "Automated_Org_Structure_test_employees.pptx" in response.headers["content-disposition"]
    
    # Check that the binary data returned is indeed a PPTX / ZIP structure (starts with PK)
    assert response.content.startswith(b"PK")
