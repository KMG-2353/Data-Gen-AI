import google.generativeai as genai
import json
import os

def configure_gemini(api_key: str = None):
    """Configure Gemini with API key"""
    key = "AIzaSyCT5Pnxw5xMPt9RAW_-mCZ8pJlv6ejTIgk" or os.getenv("GEMINI_API_KEY")
    if not key:
        raise ValueError("GEMINI_API_KEY not found")
    genai.configure(api_key=key)

def generate_test_data(headers: list[str], row_count: int) -> list[dict]:
    """Generate test data using Gemini for US insurance domain"""
    
    configure_gemini()
    
    model = genai.GenerativeModel("gemini-3-flash-preview")
    
    prompt = f"""You are a test data generator for US insurance applications.

Generate exactly {row_count} rows of realistic test data for the following columns:
{json.dumps(headers)}

IMPORTANT RULES:
1. Generate realistic US insurance test data
2. Use valid US formats:
   - Phone: (XXX) XXX-XXXX
   - SSN: XXX-XX-XXXX (use fake but valid format)
   - ZIP: 5 digits or ZIP+4
   - State: 2-letter abbreviations (CA, TX, NY, etc.)
   - Dates: MM/DD/YYYY format
3. For policy numbers, use realistic formats like: POL-XXXXXXXX
4. For currency amounts, use realistic insurance values
5. Names should be diverse and realistic
6. Addresses should be realistic US addresses

Return ONLY a valid JSON array of objects with the exact column names as keys.
No markdown, no explanation, just the JSON array.

Example format:
[{{"column1": "value1", "column2": "value2"}}, ...]
"""

    response = model.generate_content(prompt)
    
    # Parse the response
    text = response.text.strip()
    
    # Clean up any markdown formatting
    if text.startswith("```json"):
        text = text[7:]
    if text.startswith("```"):
        text = text[3:]
    if text.endswith("```"):
        text = text[:-3]
    
    data = json.loads(text.strip())
    
    return data

