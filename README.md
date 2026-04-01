# Carvana Streamlit App

This repository contains a Streamlit dashboard for Carvana (`CVNA`) using live data from SEC EDGAR and the U.S. Bureau of Labor Statistics.

## Local Run

Use the local launcher script:

```powershell
.\\start_carvana.ps1
```

Or run Streamlit directly from the project virtual environment:

```powershell
.\\.venv\\Scripts\\python -m streamlit run carvana_screen.py
```

## Deploy To Streamlit Community Cloud

1. Create a GitHub repository and push this folder.
2. Go to `https://share.streamlit.io/`.
3. Click `Create app`.
4. Select your GitHub repository and branch.
5. Set the entrypoint file to `carvana_screen.py`.
6. In `Advanced settings`, choose the same Python version you use for local development.
7. Click `Deploy`.

## Recommended Secrets

For SEC API etiquette, set a real contact string in Streamlit secrets:

```toml
SEC_USER_AGENT = "carvana-dashboard/1.0 your-email@example.com"
```

## Files Used For Deployment

- `carvana_screen.py`: Streamlit entrypoint
- `requirements.txt`: Python dependencies
- `.streamlit/config.toml`: Streamlit configuration
- `.streamlit/secrets.toml`: local secrets file for items such as `SEC_USER_AGENT` (do not commit)

## Notes

- Streamlit Community Cloud runs your app from the repository root.
- The hosted environment is Linux, so use forward slashes in paths if you add local files later.
- This app currently uses official SEC EDGAR and BLS endpoints for live data.
- If you add API keys or credentials later, store them in Community Cloud `Secrets`, not in the repository.
