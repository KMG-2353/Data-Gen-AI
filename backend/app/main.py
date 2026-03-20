from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from openpyxl import load_workbook, Workbook
from io import BytesIO
import uuid
import os
import json
import asyncio
from dotenv import load_dotenv
from typing import Any

from app.llm_service import generate_test_data, analyze_sheet_patterns, refine_rule

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

@app.get("/api/health")
async def health_check():
    """Health check endpoint for warming up the server."""
    return {"status": "ok"}


@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    try:
        # Validate file type
        if not file.filename.endswith(('.xlsx', '.xls')):
            raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are allowed")

        # Read file content
        content = await file.read()

        # Load workbook from bytes (read_only + data_only for faster parsing)
        workbook = load_workbook(BytesIO(content), read_only=True, data_only=True)

        # Extract headers AND sample data from all sheets
        sheets_data: dict[str, dict[str, Any]] = {}

        for sheet in workbook.worksheets:
            # Get all rows as list for easier access
            rows = list(sheet.iter_rows(values_only=True))
            if not rows:
                continue

            # First row is headers
            original_headers = [str(cell) if cell else "" for cell in rows[0]]
            original_headers = [h for h in original_headers if h]  # Remove empty

            if not original_headers:
                continue

            # Make duplicate headers unique by appending column position
            seen: dict[str, int] = {}
            unique_headers: list[str] = []
            for idx, h in enumerate(original_headers):
                if original_headers.count(h) > 1:
                    seen[h] = seen.get(h, 0) + 1
                    unique_headers.append(f"{h} (Col {idx + 1})")
                else:
                    unique_headers.append(h)

            # Extract sample data (up to 10 rows) for each column
            samples: dict[str, list[Any]] = {h: [] for h in unique_headers}
            data_rows = rows[1:11]  # Skip header, take up to 10 rows

            for row in data_rows:
                for col_idx, unique_header in enumerate(unique_headers):
                    if col_idx < len(row) and row[col_idx] is not None:
                        samples[unique_header].append(row[col_idx])

            sheets_data[sheet.title] = {
                "original_headers": original_headers,
                "unique_headers": unique_headers,
                "samples": samples,
                "row_count": len(rows) - 1  # Exclude header row
            }

        workbook.close()

        if not sheets_data:
            raise HTTPException(status_code=400, detail="No data found in any sheet")

        # Generate session ID
        session_id = str(uuid.uuid4())

        # Store session data with new structure
        sessions[session_id] = {
            "sheets_data": sheets_data,
            "rule_sets": {},  # Will be populated by /api/analyze
            "filename": file.filename
        }

        # Build response with sample preview
        sheets_preview = []
        for sheet_name, data in sheets_data.items():
            sheets_preview.append({
                "sheet_name": sheet_name,
                "headers": data["original_headers"],
                "unique_headers": data["unique_headers"],
                "sample_count": min(10, data["row_count"]),
                "total_rows": data["row_count"]
            })

        return {
            "session_id": session_id,
            "sheets": sheets_preview,
            "sheet_names": list(sheets_data.keys()),
            "sheet_count": len(sheets_data),
            "filename": file.filename
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing file: {str(e)}")


async def analyze_sheets_stream(session_id: str, sheets_to_analyze: list[str]):
    """Generator for SSE events during pattern analysis."""
    session = sessions[session_id]
    sheets_data = session["sheets_data"]
    total_sheets = len(sheets_to_analyze)

    for idx, sheet_name in enumerate(sheets_to_analyze):
        # Send sheet_start event
        yield f"data: {json.dumps({'event': 'sheet_start', 'sheet_name': sheet_name, 'progress': idx / total_sheets, 'message': f'Analyzing {sheet_name}...'})}\n\n"

        try:
            sheet_info = sheets_data[sheet_name]

            print(f"Analyzing sheet '{sheet_name}' with {len(sheet_info['unique_headers'])} headers")

            # Call LLM to analyze patterns
            rule_set = analyze_sheet_patterns(
                sheet_name=sheet_name,
                headers=sheet_info["unique_headers"],
                samples=sheet_info["samples"]
            )

            # Store rules in session
            session["rule_sets"][sheet_name] = rule_set

            print(f"Successfully analyzed sheet '{sheet_name}' - found {len(rule_set.get('rules', []))} rules")

            # Send sheet_complete event
            yield f"data: {json.dumps({'event': 'sheet_complete', 'sheet_name': sheet_name, 'progress': (idx + 1) / total_sheets, 'rules': rule_set})}\n\n"

        except Exception as e:
            import traceback
            error_msg = str(e)
            print(f"Error analyzing sheet '{sheet_name}': {error_msg}")
            print(traceback.format_exc())
            yield f"data: {json.dumps({'event': 'error', 'sheet_name': sheet_name, 'message': error_msg})}\n\n"

        # Small delay to allow frontend to process
        await asyncio.sleep(0.1)

    # Send complete event
    yield f"data: {json.dumps({'event': 'complete', 'progress': 1.0, 'message': 'Analysis complete'})}\n\n"


@app.post("/api/analyze")
async def analyze_patterns(request: dict):
    """Analyze patterns in uploaded data - returns SSE stream."""
    session_id = request.get("session_id")
    sheets_to_analyze = request.get("sheets_to_analyze")

    if not session_id or session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")

    session = sessions[session_id]
    sheets = sheets_to_analyze or list(session["sheets_data"].keys())

    return StreamingResponse(
        analyze_sheets_stream(session_id, sheets),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no"
        }
    )


@app.get("/api/rules/{session_id}")
async def get_rules(session_id: str):
    """Get current rule sets for a session."""
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")

    session = sessions[session_id]
    return {
        "session_id": session_id,
        "rule_sets": session.get("rule_sets", {}),
        "sheets": list(session["sheets_data"].keys())
    }


@app.put("/api/rules/update")
async def update_rule(request: dict):
    """Manually update a rule for a column."""
    session_id = request.get("session_id")
    sheet_name = request.get("sheet_name")
    column_name = request.get("column_name")
    updated_rule = request.get("updated_rule")

    if not session_id or session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")

    session = sessions[session_id]

    if sheet_name not in session.get("rule_sets", {}):
        raise HTTPException(status_code=404, detail="Sheet rules not found")

    # Find and update the rule
    rules = session["rule_sets"][sheet_name]["rules"]
    for i, rule in enumerate(rules):
        if rule["column_name"] == column_name:
            updated_rule["user_modified"] = True
            rules[i] = updated_rule
            return {
                "success": True,
                "sheet_name": sheet_name,
                "column_name": column_name,
                "updated_rule": updated_rule
            }

    raise HTTPException(status_code=404, detail="Column rule not found")


@app.post("/api/rules/reprompt")
async def reprompt_rule(request: dict):
    """Use LLM to refine a rule based on user feedback."""
    session_id = request.get("session_id")
    sheet_name = request.get("sheet_name")
    column_name = request.get("column_name")
    user_feedback = request.get("user_feedback")

    if not session_id or session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")

    session = sessions[session_id]

    if sheet_name not in session.get("rule_sets", {}):
        raise HTTPException(status_code=404, detail="Sheet rules not found")

    # Find the current rule
    rules = session["rule_sets"][sheet_name]["rules"]
    current_rule = None
    rule_index = -1

    for i, rule in enumerate(rules):
        if rule["column_name"] == column_name:
            current_rule = rule
            rule_index = i
            break

    if current_rule is None:
        raise HTTPException(status_code=404, detail="Column rule not found")

    # Get sample data for context
    samples = session["sheets_data"][sheet_name]["samples"].get(column_name, [])

    try:
        # Call LLM to refine the rule
        new_rule = refine_rule(
            current_rule=current_rule,
            user_feedback=user_feedback,
            sample_values=samples
        )

        new_rule["user_modified"] = True

        # Update the rule in session
        rules[rule_index] = new_rule

        return {
            "session_id": session_id,
            "sheet_name": sheet_name,
            "column_name": column_name,
            "old_rule": current_rule,
            "new_rule": new_rule
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error refining rule: {str(e)}")

@app.post("/api/generate")
async def generate_data(request: dict):
    """Generate test data using LLM based on verified rules."""
    session_id = request.get("session_id")
    row_count = request.get("row_count", 10)
    special_instructions = request.get("special_inst", "")
    skip_rules = request.get("skip_rules", False)  # If True, generate without rules (legacy mode)

    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")

    session = sessions[session_id]
    sheets_data = session["sheets_data"]
    rule_sets = session.get("rule_sets", {})

    try:
        generated_data_by_sheet = {}
        previous_sheets_data = {}

        for sheet_name, sheet_info in sheets_data.items():
            unique_headers = sheet_info["unique_headers"]
            original_headers = sheet_info["original_headers"]

            print(f"Generating {row_count} rows for sheet '{sheet_name}'")

            # Get rules for this sheet if available
            sheet_rules = rule_sets.get(sheet_name) if not skip_rules else None

            data = generate_test_data(
                headers=unique_headers,
                row_count=row_count,
                special_instruction=special_instructions,
                sheet_name=sheet_name,
                previous_sheets_data=previous_sheets_data if previous_sheets_data else None,
                rules=sheet_rules  # Pass verified rules
            )

            generated_data_by_sheet[sheet_name] = {
                "original_headers": original_headers,
                "unique_headers": unique_headers,
                "data": data
            }

            previous_sheets_data[sheet_name] = data
            print(f"Generated {len(data)} rows for sheet '{sheet_name}' successfully")

        # Store generated data in session
        session["generated_data_by_sheet"] = generated_data_by_sheet

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
