# SmartStore Price Monitor

Naver SmartStore competitor price monitoring tool built with Python and Tkinter.

This public repository is a sanitized demo version. Proprietary license generation,
customer data, real license files, build artifacts, and API credentials are excluded.

## Features

- Search Naver Shopping products by keyword
- Compare competitor prices against your own price
- Filter results by price range
- Save price analysis reports as Excel files
- Monitor registered products periodically
- Optional Naver Shopping API configuration

## Public Demo Mode

The public build runs in demo mode by default.

To disable demo mode:

```powershell
$env:SSTORE_PUBLIC_DEMO = "0"
python main.py
```

## Setup

```powershell
pip install -r requirements.txt
copy config.example.json config.json
python main.py
```

## Naver API

Create `config.json` from `config.example.json`, then fill in:

```json
{
  "naver_client_id": "",
  "naver_client_secret": ""
}
```

## Excluded From GitHub

- `keygen.py`
- `customers.json`
- `license.json`
- `config.json`
- `history.json`
- `reports/`
- `build/`, `dist/`, `deploy/`
- `.venv/`, `__pycache__/`
- PDF/manual files
