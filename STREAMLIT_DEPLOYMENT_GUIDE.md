# Deploy FPA Monthly Reconciliation Tool on Streamlit Cloud

## Recommendation

Use Streamlit Cloud for a private demo or controlled internal test. The app handles student names, fee amounts, Tally vouchers, and income-sheet data, so do not deploy it publicly with open access.

## Files To Push

Push only the app code:

- `streamlit_app.py`
- `fpa_monthly_reconcile.py`
- `income_sheet_posting.py`
- `requirements.txt`
- `README.md`
- `.gitignore`
- `.streamlit/config.toml`

Do not push PDFs, Excel files, `outputs/`, `work/`, `.venv/`, or any real accounting data.

## Local Test

```cmd
cd C:\Users\Syed\Documents\Codex\2026-07-02\fpa
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m streamlit run streamlit_app.py --server.port 8504
```

Open:

```text
http://localhost:8504
```

## GitHub Setup

If GitHub CLI is not logged in:

```cmd
gh auth login -h github.com
```

Choose:

```text
GitHub.com
HTTPS
Login with a web browser
```

Then create and push a private repo:

```cmd
cd C:\Users\Syed\Documents\Codex\2026-07-02\fpa
git init
git branch -M main
git add streamlit_app.py fpa_monthly_reconcile.py income_sheet_posting.py requirements.txt README.md .gitignore .streamlit/config.toml STREAMLIT_DEPLOYMENT_GUIDE.md
git commit -m "Deploy FPA reconciliation Streamlit app"
gh repo create fpa-accounting-reconciliation-tool --private --source . --remote origin --push
```

If the repo already exists:

```cmd
git remote add origin https://github.com/YOUR-USERNAME/fpa-accounting-reconciliation-tool.git
git push -u origin main
```

## Streamlit Cloud Deployment

1. Go to:

```text
https://share.streamlit.io
```

2. Click **New app**.
3. Select the GitHub repo.
4. Set:

```text
Branch: main
Main file path: streamlit_app.py
```

5. Click **Deploy**.

## Security Settings

Keep the GitHub repo private.

If Streamlit Cloud workspace supports viewer restrictions in your plan, restrict access to the accounting/process team.

Do not upload sensitive sample data into GitHub. Use the app upload buttons only during runtime.

## Post-Deploy Test

Test with sanitized/dummy files first:

1. Open the Streamlit app URL.
2. Test `PDF Extract`.
3. Test `Reconciliation`.
4. Test `Post to Income Sheet`.
5. Confirm outputs download correctly.

## Production Recommendation

For production accounting usage, prefer a private internal server or controlled Streamlit deployment after testing. Streamlit Cloud is best as a fast demo and pilot environment.
