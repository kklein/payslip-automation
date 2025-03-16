This repo contains a script which
1. Downloads the attachments of emails received on a Gmail account with a certain subject
2. Decrypts the attachments with password protection, presumed to be pdf files
3. Writes the decrypted pdf files to a local directory
4. Uploads all of these decrypted pdf files to Google Drive

## Executing

1. Create a GCP project with Gmail and Google Drive APIs enabled
    - See e.g. the [GMail quickstart](https://developers.google.com/gmail/api/quickstart/python)
3. Download the OAuth2 Client ID and write it to `credentials.json`.
4. Run
```bash
pixi run python payslip.py --subject "Payslip"
```
