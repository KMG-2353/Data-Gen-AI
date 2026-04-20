from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from openpyxl import load_workbook, Workbook
from io import BytesIO
import uuid
import os
from dotenv import load_dotenv

from app.llm_service import generate_test_data, detect_policy_type
from app.policies import get_handler
from app.policies.ims import (
    canonical_sheet_name as _ims_canonical_sheet_name,
    extract_lob_flags as _ims_extract_lob_flags,
    filter_rows_by_lob as _ims_filter_rows_by_lob,
)

load_dotenv()

app = FastAPI()

# Configure CORS with environment-aware origins
origins = [
    "http://localhost:5173",  # Development
    "https://data-gen-ai-1.onrender.com",  # Production frontend
]

# Add additional frontend URL from environment if specified
frontend_url = os.getenv("FRONTEND_URL")
if frontend_url and frontend_url not in origins:
    origins.append(frontend_url)

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
        
        # Canonicalize sheet names when the upload is an IMS workbook
        # (corrects known misspellings like "LIBIALITY" → "LIABILITY").
        policy_type_for_upload = detect_policy_type(file.filename)

        # Extract headers from first row of all sheets
        headers_by_sheet = {}       # original names (for display & Excel output)
        unique_headers_by_sheet = {} # deduplicated names (for LLM)
        for sheet in workbook.worksheets:
            original_headers = []
            for cell in sheet[1]:
                if cell.value:
                    original_headers.append(str(cell.value))
            if not original_headers:
                continue
            
            # Make duplicate headers unique by appending column position
            seen = {}
            unique_headers = []
            for idx, h in enumerate(original_headers):
                if original_headers.count(h) > 1:
                    # Track occurrence number
                    seen[h] = seen.get(h, 0) + 1
                    unique_headers.append(f"{h} (Col {idx + 1})")
                else:
                    unique_headers.append(h)
            
            sheet_title = sheet.title
            if policy_type_for_upload == "IMS":
                sheet_title = _ims_canonical_sheet_name(sheet_title)
            headers_by_sheet[sheet_title] = original_headers
            unique_headers_by_sheet[sheet_title] = unique_headers
        
        workbook.close()
        
        if not headers_by_sheet:
            raise HTTPException(status_code=400, detail="No headers found in any sheet")
        
        # Generate session ID
        session_id = str(uuid.uuid4())
        
        # Store session data
        sessions[session_id] = {
            "headers_by_sheet": headers_by_sheet,
            "unique_headers_by_sheet": unique_headers_by_sheet,
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
    
    unique_headers_by_sheet = sessions[session_id]["unique_headers_by_sheet"]
    original_headers_by_sheet = sessions[session_id]["headers_by_sheet"]
    filename = sessions[session_id].get("filename", "")
    policy_type = detect_policy_type(filename)
    handler = get_handler(policy_type)
    print(f"Detected policy_type={policy_type} for filename='{filename}'")

    try:
        # Generate data for ALL sheets sequentially, passing previous data for consistency
        generated_data_by_sheet = {}
        previous_sheets_data = {}

        # Track policy, driver, and vehicle data for cross-sheet row expansion
        policy_data = None
        driver_data = None
        vehicle_data = None

        # IMS-only: LOB flags extracted from the Policy sheet so we can filter
        # sub-sheet rows where the LOB is toggled off.
        lob_flags: dict[int, dict[str, bool]] = {}

        # Two-pass approach: process non-assignment sheets first so that
        # policy/driver/vehicle data are always available when the assignment
        # sheet is processed, regardless of workbook sheet order.
        sheet_items = list(unique_headers_by_sheet.items())
        non_assignment_sheets = [
            (name, hdrs) for name, hdrs in sheet_items
            if handler.detect_sheet_type(name) != "assignment"
        ]
        assignment_sheets = [
            (name, hdrs) for name, hdrs in sheet_items
            if handler.detect_sheet_type(name) == "assignment"
        ]

        for sheet_name, unique_headers in non_assignment_sheets + assignment_sheets:
            original_headers = original_headers_by_sheet[sheet_name]

            # Handler decides row count and sheet-specific prompt augmentation
            sheet_type = handler.detect_sheet_type(sheet_name)
            effective_row_count, additional_rules = handler.build_sheet_context(
                sheet_name=sheet_name,
                policy_data=policy_data,
                driver_data=driver_data,
                original_row_count=row_count,
                special_instruction=special_instructions,
            )

            print(f"Generating {effective_row_count} rows for sheet '{sheet_name}' (type={sheet_type}) with headers: {unique_headers}")

            # Deterministic pre-generation path (e.g. PAP assignment sheet).
            pre_generated = handler.pre_generate(
                sheet_name=sheet_name,
                unique_headers=unique_headers,
                policy_data=policy_data,
                driver_data=driver_data,
                vehicle_data=vehicle_data,
            )

            if pre_generated is not None:
                print(f"Using deterministic pre-generation for sheet '{sheet_name}'")
                data = pre_generated
            elif effective_row_count == 0:
                # Zero-row sheets (e.g. infraction with no drivers flagged)
                data = []
            else:
                data = generate_test_data(
                    headers=unique_headers,
                    row_count=row_count,
                    special_instruction=special_instructions,
                    sheet_name=sheet_name,
                    previous_sheets_data=previous_sheets_data if previous_sheets_data else None,
                    row_count_override=effective_row_count,
                    additional_rules=additional_rules,
                )

            # Run handler-specific post-processing (IMS sheet enforcers +
            # cross-sheet consistency; generic/PAP are no-ops here).
            data = handler.post_process(
                rows=data,
                sheet_name=sheet_name,
                special_instruction=special_instructions,
                previous_sheets_data=previous_sheets_data if previous_sheets_data else None,
            )

            # IMS: once the Policy sheet is done, extract LOB flags so later
            # sub-sheets can be filtered (Defect #216: LOB=No → drop row).
            if policy_type == "IMS" and sheet_type == "policy":
                lob_flags = _ims_extract_lob_flags(data)
                print(f"Extracted LOB flags from '{sheet_name}': {len(lob_flags)} rows")

            # IMS: filter sub-sheet rows against the LOB toggles.
            if policy_type == "IMS" and lob_flags:
                original_count = len(data)
                data = _ims_filter_rows_by_lob(data, sheet_name, lob_flags)
                if len(data) < original_count:
                    print(
                        f"LOB filtering: removed {original_count - len(data)} "
                        f"rows from '{sheet_name}' (LOB=No)"
                    )

            generated_data_by_sheet[sheet_name] = {
                "original_headers": original_headers,
                "unique_headers": unique_headers,
                "data": data
            }

            # Add this sheet's data to context for subsequent sheets
            previous_sheets_data[sheet_name] = data

            # Track policy, driver, and vehicle data for downstream sheets
            if sheet_type == "policy":
                policy_data = data
            elif sheet_type == "driver":
                driver_data = data
            elif sheet_type == "vehicle":
                vehicle_data = data

            print(f"Generated {len(data)} rows for sheet '{sheet_name}' successfully")
        
        # Restore original workbook sheet order for the Excel output
        generated_data_by_sheet = {
            name: generated_data_by_sheet[name]
            for name in unique_headers_by_sheet
            if name in generated_data_by_sheet
        }

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
        original_headers = sheet_data["original_headers"]
        unique_headers = sheet_data["unique_headers"]
        data = sheet_data["data"]
        
        # Write original headers in row 1
        for col, header in enumerate(original_headers, 1):
            sheet.cell(row=1, column=col, value=header)
        
        # Write data rows using unique header keys to look up values
        for row_idx, row_data in enumerate(data, 2):
            for col_idx, unique_header in enumerate(unique_headers, 1):
                value = row_data.get(unique_header, "")
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
