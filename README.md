This repo contains a script which
1. Downloads the attachments of emails received on a Gmail account with a certain subject
2. Decrypts the attachments with password protection, presumed to be pdf files
3. Writes the decrypted pdf files to a local directory
4. Uploads all of these decrypted pdf files to Google Drive

## Executing

1. Create a GCP project with Gmail and Google Drive APIs enabled
    - See e.g. the [GMail quickstart](https://developers.google.com/gmail/api/quickstart/python)
2. Run
```bash
pixi run python payslip.py --subject "Payslip"
```

Authentication is handled interactively:
- If no OAuth2 client secrets (`credentials.json`) are found next to `payslip.py`, the Google Cloud console
  opens in the browser. Create a 'Desktop app' OAuth client, download its JSON and confirm the path when
  prompted (the latest `~/Downloads/client_secret_*.json` is suggested).
- If no valid token (`token.json`) is found, or it can't be refreshed, a Google login window opens in the browser.

Both files are git-ignored.
