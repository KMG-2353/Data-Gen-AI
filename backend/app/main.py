from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from openpyxl import load_workbook, Workbook
from io import BytesIO
import uuid
import os
from dotenv import load_dotenv

from llm_service import generate_test_data

load_dotenv()

app = FastAPI()

# Configure CORS
origins = [
    "http://localhost:5173",  # Allow your Vite app's origin
]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Store uploaded file data temporarily
sessions = {}

@app.get("/api/hello")
def read_hello():
    return {"message": "Test123"}

@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    try:
        # Validate file type
        if not file.filename.endswith(('.xlsx', '.xls')):
            raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are allowed")
        
        # Read file content
        content = await file.read()
        
        # Load workbook from bytes
        workbook = load_workbook(BytesIO(content), read_only=True)
        
        # Extract headers from first row of all sheets
        headers_by_sheet = {}
        for sheet in workbook.worksheets:
            sheet_headers = []
            for cell in sheet[1]:
                if cell.value:
                    sheet_headers.append(str(cell.value))
            if sheet_headers:
                headers_by_sheet[sheet.title] = sheet_headers
        
        workbook.close()
        
        if not headers_by_sheet:
            raise HTTPException(status_code=400, detail="No headers found in any sheet")
        
        # Generate session ID
        session_id = str(uuid.uuid4())
        
        # Store session data
        sessions[session_id] = {
            "headers_by_sheet": headers_by_sheet,
            "filename": file.filename
        }
        
        return {
            "session_id": session_id,
            "headers_by_sheet": headers_by_sheet,
            "sheet_names": list(headers_by_sheet.keys()),
            "sheet_count": len(headers_by_sheet),
            "filename": file.filename
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing file: {str(e)}")

@app.post("/api/generate")
async def generate_data(request: dict):
    """Generate test data using LLM"""
    session_id = request.get("session_id")
    row_count = request.get("row_count", 10)
    selected_sheet = request.get("sheet_name")  # Optional: specific sheet
    
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    headers_by_sheet = sessions[session_id]["headers_by_sheet"]
    
    # If no specific sheet selected, use the first one
    if selected_sheet and selected_sheet in headers_by_sheet:
        headers = headers_by_sheet[selected_sheet]
    else:
        # Use first sheet's headers
        first_sheet = list(headers_by_sheet.keys())[0]
        headers = headers_by_sheet[first_sheet]
    
    try:
        # Generate data using Gemini
        print(f"Generating {row_count} rows for headers: {headers}")
        data = generate_test_data(headers, row_count)
        print(f"Generated {len(data)} rows successfully")
        
        # Store generated data in session
        sessions[session_id]["generated_data"] = data
        sessions[session_id]["generated_headers"] = headers
        
        return {
            "session_id": session_id,
            "data": data,
            "row_count": len(data),
            "status": "complete"
        }
    except Exception as e:
        import traceback
        print(f"Error generating data: {traceback.format_exc()}")
        raise HTTPException(status_code=500, detail=f"Error generating data: {str(e)}")


@app.get("/api/download/{session_id}")
async def download_excel(session_id: str):
    """Download generated data as Excel file"""
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    session = sessions[session_id]
    
    if "generated_data" not in session:
        raise HTTPException(status_code=400, detail="No generated data found")
    
    headers = session.get("generated_headers", list(session["headers_by_sheet"].values())[0])
    data = session["generated_data"]
    
    # Create Excel workbook
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Test Data"
    
    # Write headers
    for col, header in enumerate(headers, 1):
        sheet.cell(row=1, column=col, value=header)
    
    # Write data rows
    for row_idx, row_data in enumerate(data, 2):
        for col_idx, header in enumerate(headers, 1):
            value = row_data.get(header, "")
            sheet.cell(row=row_idx, column=col_idx, value=value)
    
    # Auto-adjust column widths
    for column in sheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        sheet.column_dimensions[column_letter].width = min(max_length + 2, 50)
    
    # Save to BytesIO
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    
    filename = f"test_data_{session_id[:8]}.xlsx"
    
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )
