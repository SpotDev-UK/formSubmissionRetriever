# HubSpot Form Submission Exporter

A Node.js script that exports HubSpot form submissions to Excel.

## Features

- Retrieves all HubSpot forms matching a search term
- Exports form submissions from the last N days
- Generates an Excel workbook with detailed submission data
- Respects HubSpot's rate limits
- Environment-based configuration

## Prerequisites

- Node.js v18 or higher
- A HubSpot private-app token with:
  - forms (read)
  - business-intelligence (read)

## Setup

1. Clone this repository
2. Install dependencies:
   ```bash
   npm install
   ```
3. Copy `.env.example` to `.env`:
   ```bash
   cp .env.example .env
   ```
4. Edit `.env` and add your HubSpot private-app token

## Usage

Run the script with:

```bash
node formSubmissionProcess.js
```

### Environment Variables

- `HUBSPOT_PAT` (Required): Your HubSpot private-app token
- `REQUEST_TERM` (Optional): Text to match in form names (default: 'exprom')
- `DAYS_BACK` (Optional): Look-back window in days (default: 30)

Example with custom settings:
```bash
REQUEST_TERM="Contact" DAYS_BACK=90 node formSubmissionProcess.js
```

## Output

The script generates `form-submissions.xlsx` with the following columns:
- Form GUID
- Form Name
- Form Submission ID
- Time Submitted (ISO 8601)
- Contact Email Address
- Page converted on
- All Properties (JSON-stringified array)

## Security Notes

- Never commit your `.env` file or expose your HubSpot PAT
- The `.gitignore` file is configured to prevent accidental commits of sensitive data 