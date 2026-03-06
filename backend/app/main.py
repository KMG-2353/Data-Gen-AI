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

# @app.get("/api/hello")
# def read_hello():
#     return {"message": "Test123"}

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
    """Generate test data using LLM for all sheets"""
    session_id = request.get("session_id")
    row_count = request.get("row_count", 10)
    special_instructions=request.get("special_inst","")
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    headers_by_sheet = sessions[session_id]["headers_by_sheet"]
    
    try:
        # Generate data for ALL sheets sequentially, passing previous data for consistency
        generated_data_by_sheet = {}
        previous_sheets_data = {}
        
        for sheet_name, headers in headers_by_sheet.items():
            print(f"Generating {row_count} rows for sheet '{sheet_name}' with headers: {headers}")
            
            # Pass previously generated sheets so LLM can maintain cross-sheet consistency
            data = generate_test_data(
                headers=headers,
                row_count=row_count,
                special_instruction=special_instructions,
                sheet_name=sheet_name,
                previous_sheets_data=previous_sheets_data if previous_sheets_data else None
            )
            
            generated_data_by_sheet[sheet_name] = {
                "headers": headers,
                "data": data
            }
            
            # Add this sheet's data to context for subsequent sheets
            previous_sheets_data[sheet_name] = data
            
            print(f"Generated {len(data)} rows for sheet '{sheet_name}' successfully")
        
        # Store generated data in session
        sessions[session_id]["generated_data_by_sheet"] = generated_data_by_sheet
        
        return {
            "session_id": session_id,
            "sheets_generated": list(generated_data_by_sheet.keys()),
            "row_count_per_sheet": row_count,
            "status": "complete"
        }
    except Exception as e:
        import traceback
        print(f"Error generating data: {traceback.format_exc()}")
        raise HTTPException(status_code=500, detail=f"Error generating data: {str(e)}")


@app.get("/api/download/{session_id}")
async def download_excel(session_id: str):
    """Download generated data as Excel file with all sheets"""
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    session = sessions[session_id]
    
    if "generated_data_by_sheet" not in session:
        raise HTTPException(status_code=400, detail="No generated data found")
    
    generated_data_by_sheet = session["generated_data_by_sheet"]
    
    # Create Excel workbook
    workbook = Workbook()
    # Remove default sheet
    workbook.remove(workbook.active)
    
    # Create a sheet for each generated dataset
    for sheet_name, sheet_data in generated_data_by_sheet.items():
        sheet = workbook.create_sheet(title=sheet_name[:31])  # Excel limits sheet names to 31 chars
        headers = sheet_data["headers"]
        data = sheet_data["data"]
        
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
