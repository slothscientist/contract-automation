# Contract Automation System

Automated contract generation and signing system using DocuSeal API. This tool generates custom service agreements in .docx format and automatically sends them for digital signature via DocuSeal.

## Features

- 📝 Generates professional service agreements in .docx format
- ✍️ Integrates with DocuSeal for digital signatures
- 📧 Automatically sends contracts via email
- ⚙️ Configurable via command-line parameters
- 🎨 Professional document formatting with proper structure

## Prerequisites

- Node.js (v14 or higher)
- DocuSeal API key
- npm or yarn

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd contract-automation
```

2. Install dependencies:
```bash
npm install
```

3. Create a `.env` file in the root directory:
```bash
cp .env.example .env
```

4. Edit `.env` and add your credentials:
```
DOCUSEAL_API_KEY=your_api_key_here
PROVIDER_EMAIL=your_email@example.com
```

## Usage

Run the script with the following command-line arguments:

```bash
node generate-contract.js \
  --clientName "Acme Corporation" \
  --signerName "John Doe" \
  --signerEmail "john.doe@acmecorp.com" \
  --projectStartDate "2026-02-01" \
  --projectEndDate "2026-03-01" \
  --totalFee "5000"
```

### Required Parameters

| Parameter | Description | Example |
|-----------|-------------|---------|
| `--clientName` | Client company or individual name | "Acme Corporation" |
| `--signerName` | Name of person signing for client | "John Doe" |
| `--signerEmail` | Email address of signer | "john@acmecorp.com" |
| `--projectStartDate` | Project start date | "2026-02-01" |
| `--projectEndDate` | Project end date | "2026-03-01" |
| `--totalFee` | Total project fee (numeric) | "5000" |

## How It Works

1. **Document Generation**: The script uses the `docx` library to create a professional service agreement with:
   - Company branding (WE GO OAKLAND)
   - Client and project details
   - Standard terms and conditions
   - Signature fields with DocuSeal tags

2. **DocuSeal Integration**: The generated document is:
   - Uploaded to DocuSeal as a template
   - Configured with signature fields for both parties
   - Sent automatically via email to all signers

3. **Output**: Generated contracts are saved in the `output/` directory with timestamped filenames.

## Example Output

```
🚀 Starting contract generation...

📋 Contract Details:
   Client: Acme Corporation
   Signer: John Doe (john@acmecorp.com)
   Duration: 2026-02-01 to 2026-03-01
   Fee: $5,000

📝 Generating contract document...
✅ Contract saved: /output/service-agreement-acme-corporation-2026-01-15.docx

📤 Uploading document to DocuSeal...
✅ Template created: tmpl_xxxxx
✅ Document sent for signing!
📧 Submission ID: sub_xxxxx
🔗 Submission URL: https://docuseal.com/s/xxxxx

✨ Process completed successfully!
```

## File Structure

```
contract-automation/
├── generate-contract.js    # Main script
├── package.json            # Node.js dependencies
├── .env                    # Environment variables (not tracked)
├── .env.example            # Example environment file
├── .gitignore              # Git ignore rules
├── output/                 # Generated contracts (created automatically)
└── README.md               # This file
```

## Environment Variables

- `DOCUSEAL_API_KEY`: Your DocuSeal API key (required)
- `PROVIDER_EMAIL`: Your email address as the service provider (optional, defaults to salaams@wegooakland.com)

## Security Notes

- Never commit your `.env` file to version control
- Keep your DocuSeal API key secure
- The `.gitignore` file excludes sensitive files by default

## Troubleshooting

### "DOCUSEAL_API_KEY not found"
Make sure you've created a `.env` file with your API key.

### "Missing required parameters"
Ensure all required command-line arguments are provided.

### DocuSeal API errors
- Verify your API key is valid
- Check your DocuSeal account has sufficient credits
- Ensure email addresses are valid

## Customization

To customize the contract template, edit the `generateContract()` function in `generate-contract.js`. You can modify:
- Document structure and sections
- Terms and conditions
- Formatting and styling
- Additional fields

## License

ISC

## Support

For issues or questions, please contact: salaams@wegooakland.com
