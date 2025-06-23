# Fabula Data Sync

A scheduled Python script to extract purchase data from Fabula Coatings' MS SQL Server database and export it to Google Sheets. VPN is used for secure access.

## 🔧 Setup

1. Copy `.env.example` to `.env` and fill in the credentials.
2. Place your `.ovpn` VPN file and `credentials.json` in the `credentials/` folder.
3. Install dependencies:
   ```bash
   pip install -r requirements.txt