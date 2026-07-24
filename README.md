# 📊 RCM AI Tool
**Risk & Control Matrix Redrafting & Design Effectiveness Review**
Powered by Groq AI | Big 4 Audit Methodology | ICFR / SOX / IFC

---

## 🚀 Deploy to Streamlit Community Cloud (Free)

### Step 1 — Upload to GitHub
1. Go to [github.com](https://github.com) and sign in (or create a free account)
2. Click **"New repository"**
3. Name it `rcm-ai-tool` (or any name you like)
4. Set it to **Private** (recommended — keeps your code safe)
5. Click **"Create repository"**
6. Upload these three files:
   - `app.py`
   - `requirements.txt`
   - `README.md`

### Step 2 — Deploy on Streamlit Cloud
1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Sign in with your GitHub account
3. Click **"New app"**
4. Select your repository: `rcm-ai-tool`
5. Branch: `main`
6. Main file path: `app.py`
7. Click **"Deploy"**

Streamlit will install all dependencies automatically and your app will be live at:
`https://<your-app-name>.streamlit.app`

### Step 3 — Get your Groq API Key
1. Go to [console.groq.com/keys](https://console.groq.com/keys)
2. Sign up for a free account
3. Click **"Create API Key"**
4. Copy the key
5. Paste it into the sidebar of your deployed app

---

## 🔐 Optional — Store API Key Securely (Recommended)

Instead of entering the key manually every time, store it as a Streamlit secret:

1. In Streamlit Cloud, open your app → click **"⋮" → "Settings" → "Secrets"**
2. Add the following:
```toml
GROQ_API_KEY = "your_groq_api_key_here"
```
3. The app will auto-read this key on startup

---

## 📦 Dependencies

| Package | Purpose |
|---|---|
| `streamlit` | Web app framework |
| `pandas` | Excel file handling |
| `openpyxl` | Read/write .xlsx files |
| `groq` | Groq AI API client |
| `pydantic` | Data validation for API responses |
| `reportlab` | PDF report generation |
| `python-docx` | Word (.docx) report generation |

Install locally:
```bash
pip install streamlit pandas openpyxl groq pydantic reportlab python-docx
```

Run locally:
```bash
streamlit run app.py
```

---

## 🛠️ Features

### ✏️ Mode 1 — Redraft & Classify Controls
- Rewrites every control in **Big 4 audit style** (WHO / WHAT / HOW / WHEN / EVIDENCE)
- Classifies **Key vs Non-Key** using ICFR/SOX criteria
- Classifies **Anti-Fraud (Y/N)** using fraud risk framework
- Assigns **Risk Rating** (High / Medium / Low)
- Exports updated RCM as **Excel (.xlsx)**

### 🔍 Mode 2 — Full RCM Design Effectiveness Review
- 14-part review benchmarked against **COSO 2013, IFC, ICAI, SOX**
- Identifies missing controls, weak controls, SoD issues, anti-fraud gaps
- Maps controls to **Financial Statement Assertions** and **COSO components**
- Provides **control testing procedures** for every key control
- Suggests **automation opportunities** via ERP
- Generates **Executive Report** with Priority Matrix and Maturity Score
- Downloads report as **PDF**, **Word (.docx)**, or **Markdown**

---

## 📋 Supported Processes & Industries

**Processes:** P2P, O2C, R2R, H2R, Fixed Assets, Treasury, Payroll, Revenue Recognition, and more

**Industries:** Steel, FMCG, Pharma, Auto, IT/Software, NBFC, Real Estate, Power, Retail, and more

**ERP Systems:** SME Assist, SAP S/4HANA, SAP ECC, Oracle Fusion, Tally Prime, and more

---

## ⚠️ Troubleshooting

| Error | Fix |
|---|---|
| `Invalid API key` | Get a new key from console.groq.com/keys |
| `Rate limit hit` | Reduce batch size in sidebar or increase retry delay |
| `No rows processed` | Reduce batch size to 5 and retry |
| `Word export failed` | Run `pip install python-docx` |
| `PDF export failed` | Run `pip install reportlab` |

---

*Built with Groq AI · Streamlit · Big 4 ICFR Methodology*
