This repo contains a script which:
1. Downloads the attachments of gmail emails with a certain subject in the title
2. Decrypts the attachments, presumed to be pdf files
3. Writes the decrypted pdf files to a local directory
4. Uploads all of these decrypted pdf files to Google Drive


## Executing

1. Create a GCP project with gmail and google drive APIs enabled
2. Download the OAuth2 Client ID and save it in `credentials.json`.
3. Run
```bash
pixi run python payslip.py --subject "Payslip"
```
