# Colpensiones PDF Parser FastAPI Service

This FastAPI service wraps the Python PDF parsing logic and provides a REST API endpoint.

## Local Development

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Copy your PDF extractor:**
   ```bash
   # Copy pdf_table_extractor.py to this directory
   ```

3. **Run the server:**
   ```bash
   python main.py
   # or
   uvicorn main:app --reload --port 8000
   ```

4. **Test the API:**
   ```bash
   curl -X POST "http://localhost:8000/parse-pension-pdf" \
     -H "Content-Type: multipart/form-data" \
     -F "file=@your-pension.pdf"
   ```

## API Endpoints

### POST /parse-pension-pdf
- **Description:** Parse a Colpensiones PDF
- **Content-Type:** multipart/form-data
- **Body:** file (PDF)
- **Response:**
  ```json
  {
    "weeks_data": [...],
    "summary_values": {...},
    "payments_data": [...]
  }
  ``
