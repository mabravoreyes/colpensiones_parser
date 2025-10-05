# Colpensiones PDF Parser FastAPI Service

This FastAPI service wraps your Python PDF parsing logic and provides a REST API endpoint.

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

## Deployment Options

### Option 1: Deploy to Railway (Easiest)

1. Create account at [railway.app](https://railway.app)
2. Click "New Project" → "Deploy from GitHub repo"
3. Connect your repo or upload this folder
4. Railway will auto-detect the Dockerfile and deploy
5. Copy the public URL (e.g., `https://your-app.railway.app`)

### Option 2: Deploy to Render

1. Create account at [render.com](https://render.com)
2. Click "New" → "Web Service"
3. Connect your repo or upload this folder
4. Select "Docker" as environment
5. Deploy and copy the public URL

### Option 3: Deploy to Google Cloud Run

1. Install gcloud CLI and authenticate
2. Build and push container:
   ```bash
   gcloud builds submit --tag gcr.io/YOUR_PROJECT_ID/pension-parser
   ```
3. Deploy to Cloud Run:
   ```bash
   gcloud run deploy pension-parser \
     --image gcr.io/YOUR_PROJECT_ID/pension-parser \
     --platform managed \
     --region us-central1 \
     --allow-unauthenticated
   ```
4. Copy the service URL

### Option 4: Deploy to AWS Lambda (with Function URLs)

1. Use AWS Lambda with container image support
2. Deploy using AWS SAM or CDK
3. Enable Function URL for HTTP access

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
  ```

### GET /health
- **Description:** Health check endpoint
- **Response:** `{"status": "healthy"}`

## Environment Variables

None required - the service is stateless.

## After Deployment

1. Copy your public API URL
2. Update the Supabase edge function with: `FASTAPI_PARSER_URL=https://your-service.com`
3. Test the integration
