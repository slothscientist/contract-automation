# AI-Powered Contract Generator Setup

## Overview

The AI-powered contract generator allows you to **paste your entire proposal** and have AI automatically extract all contract details including:
- Client information
- Dates and timelines
- Payment schedule
- Services description

## Quick Start

### 1. Get Your OpenAI API Key

1. Go to https://platform.openai.com/signup
2. Create an account (or sign in)
3. Go to https://platform.openai.com/api-keys
4. Click "Create new secret key"
5. Copy the key (starts with `sk-...`)

### 2. Add API Key to .env

Edit your `.env` file and add your OpenAI API key:

```
OPENAI_API_KEY=sk-your-actual-key-here
```

### 3. Start the Server

```bash
npm start
```

### 4. Open the App

Open your browser to: http://localhost:3000

You'll see the AI-powered interface at `index-ai.html`

## How to Use

### Step 1: Paste Your Proposal

Copy your entire proposal, email, or contract details and paste it into the text area.

Example of what you can paste:

```
Client: Two Sides of Sunlight LLC
Managing Members: Sara Rahimian, Stacie Frederick
Email: sara@client.com
Project: January 19, 2026 to March 2, 2026
Total Fee: $9,500

Payment Schedule:
- $3,800 non-refundable deposit, due December 26, 2025
- $2,850 midpoint payment, due February 16, 2026
- $2,850 final payment, due March 2, 2026

Services:
Website Design & Development
- Up to 10 pages of custom content
- Brand refinement
- SEO setup
- Google Analytics integration
...

Email Marketing System
- Mailchimp setup
- Welcome email sequence
- Newsletter template
...
```

### Step 2: Extract with AI

Click the "Extract Contract Details with AI" button.

AI will parse your proposal and extract:
- Client name and entity type
- Managing members
- Email addresses
- Contract dates
- Payment schedule (all payments with amounts and dates)
- Full services description

### Step 3: Review & Edit

The extracted data will be displayed in editable fields. Review and make any corrections needed.

### Step 4: Generate Contract

Click "Generate & Send Contract" to:
- Create the full WeGo Oakland contract with all 19 sections
- Add DocuSeal signature, initial, and text fields
- Upload to DocuSeal
- Send signing emails to all parties

## Cost

OpenAI API costs approximately **$0.01-0.05 per contract** using GPT-4o-mini.

## Features

### What AI Extracts:
- ✅ Client name, entity type, managing members
- ✅ Email addresses (primary and secondary signers)
- ✅ Contract effective date and term end date
- ✅ Complete payment schedule with amounts and due dates
- ✅ Full services description with all details

### DocuSeal Fields:
- `{{s:Provider}}` - Your signature
- `{{s:Customer1}}` - First customer signature
- `{{s:Customer2}}` - Second customer signature (if applicable)
- `{{i:Customer1}}` - Customer initials throughout Exhibit A
- `{{i:Customer2}}` - Second customer initials (if applicable)
- `{{t:Customer1:Address}}` - Customer fills in their address
- `{{t:Customer1:Name}}` - Customer fills in their name (if needed)

## Troubleshooting

### "OpenAI API Error"
- Check that your OPENAI_API_KEY is set correctly in `.env`
- Make sure your OpenAI account has credits
- Verify your API key is active at https://platform.openai.com/api-keys

### "Failed to parse proposal"
- Make sure your proposal has enough detail (at least 50 characters)
- Include key information: client name, dates, fees, services
- The more detail you provide, the better AI can extract

### AI Extracted Wrong Information
- Edit any fields in Step 2 before generating
- All extracted data is editable before you generate the contract

## Alternative: Manual Mode

If you prefer not to use AI, you can still use the manual form:

```bash
npm run start:basic
```

This will start the basic server with the manual form interface.

## Support

For questions or issues:
- Email: salaams@wegooakland.com
- Check TROUBLESHOOTING.md for common issues
