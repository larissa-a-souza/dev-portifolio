# NFS-e Issuer Assistant

Automation script that helps issue service invoices (NFS-e) in the Brazilian national portal.
It reads an Excel draft, finds the client's location, fills the form and pauses for a human to confirm the tax rate.

> **Status:** In progress 🚧

## Features
- Reads an `.xlsx` draft file from the `executar/` folder
- Detects the location (city/state) inside the spreadsheet
- Picks the correct client tax ID (CNPJ) for that location
- Logs in and fills the invoice form automatically (Playwright)
- Pauses for human review and saves the tax rate back to the spreadsheet

## Tech stack
- Node.js
- Playwright (browser automation)
- SheetJS / xlsx (Excel reading and writing)
- dotenv (environment variables)

## Setup
```bash
npm install
npx playwright install chromium
```

1. Copy `.env.example` to `.env` and add your login.
2. Copy `tomadores.example.json` to `tomadores.json` and add the client tax IDs.
3. Put the Excel draft inside the `executar/` folder.

## Run
```bash
node emissor.js
```

## Security
Credentials and client data are **not** stored in the code.
They live in `.env` and `tomadores.json`, which are ignored by Git.

## Author
Larissa Souza — IT Support | Networking | Infrastructure | Data Analysis
