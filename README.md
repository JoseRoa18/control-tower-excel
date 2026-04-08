# Control Tower Excel Generator - Deployment Guide

## For Render.com (Free, Recommended)

### Step 1: Create a GitHub repository
1. Go to github.com → New Repository
2. Name it: `control-tower-excel`
3. Upload these 4 files:
   - `app.py`
   - `requirements.txt`
   - `Procfile`
   - `render.yaml`

### Step 2: Deploy on Render
1. Go to render.com → Sign up (free)
2. Click "New" → "Web Service"
3. Connect your GitHub repo
4. Settings:
   - Name: `control-tower-excel`
   - Runtime: Python
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app --bind 0.0.0.0:$PORT --timeout 300`
   - Plan: Free
5. Click "Deploy"
6. Wait 2-3 minutes
7. Your URL will be: `https://control-tower-excel.onrender.com`

### Step 3: Test
```bash
curl https://control-tower-excel.onrender.com/health
```
Should return: `{"status": "ok"}`

## API Usage

### POST /generate
Send CSV data as JSON, receive xlsx file.

```json
{
  "p2s_csv": "SKU,Marketplace,Price\nA-02,Amazon,57\n...",
  "sales_csv": "Date,SKU,Qty,Amount,Class\n...",
  "inventory_csv": "SKU,Available\n...",
  "master_csv": "Product/Service,Type,Description\n...",
  "promo_csv": "SKU,Retail,WC\n...(optional)",
  "pricelist_csv": "SKU,Retail,WC\n...(optional)",
  "return_base64": true
}
```

Response:
```json
{
  "file": "base64-encoded-xlsx...",
  "filename": "Marketplace_Control_Tower_USA_2026_04_08.xlsx"
}
```

## n8n Integration

In n8n, use HTTP Request node:
- Method: POST
- URL: https://control-tower-excel.onrender.com/generate
- Body: JSON with CSV data from Extract nodes
- Response contains base64 xlsx to send via email
