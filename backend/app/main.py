from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import load_workbook
from io import BytesIO
import uuid

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
        sheet = workbook.active
        
        # Extract headers from first row
        headers = []
        for cell in sheet[1]:
            if cell.value:
                headers.append(str(cell.value))
        
        workbook.close()
        
        if not headers:
            raise HTTPException(status_code=400, detail="No headers found in the first row")
        
        # Generate session ID
        session_id = str(uuid.uuid4())
        
        # Store session data
        sessions[session_id] = {
            "headers": headers,
            "filename": file.filename
        }
        
        return {
            "session_id": session_id,
            "headers": headers,
            "column_count": len(headers),
            "filename": file.filename
        }
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing file: {str(e)}")

