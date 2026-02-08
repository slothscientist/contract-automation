#!/usr/bin/env node

require('dotenv').config();
const express = require('express');
const cors = require('cors');
const OpenAI = require('openai');
const { DocusealApi } = require('@docuseal/api');

const app = express();
const PORT = process.env.PORT || 3000;

// DocuSeal Template ID - "Contract Template"
const TEMPLATE_ID = 2786147;

// Initialize OpenAI
const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY
});

// Middleware
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static('public'));

// AI-powered proposal parser
app.post('/api/parse-proposal', async (req, res) => {
  try {
    const { proposalText } = req.body;

    if (!proposalText || proposalText.trim().length < 50) {
      return res.status(400).json({ error: 'Please provide a valid proposal text' });
    }

    console.log('Parsing proposal with AI...');

    const completion = await openai.chat.completions.create({
      model: "gpt-4o-mini",
      messages: [
        {
          role: "system",
          content: `You are an expert at extracting contract information from proposals. Extract the following information and return it as JSON:
{
  "clientName": "Full client company/business name",
  "clientEntity": "Entity type (e.g., 'a limited liability company' or leave empty)",
  "clientManagingMembers": "Comma-separated list of managing members/owners (if mentioned)",
  "signerEmail": "Primary email address",
  "signerEmail2": "Second email if multiple signers",
  "effectiveDate": "Contract start date in MM/DD/YYYY format",
  "termEndDate": "Contract end date in MM/DD/YYYY format",
  "totalProjectFee": "Total project fee as number",
  "payment1": "First payment: amount and description, e.g. '$4,500 deposit upon signing'",
  "payment2": "Second payment: amount and description, e.g. '$4,500 balance at launch'",
  "launchDate": "Project completion/launch date in MM/DD/YYYY format",
  "servicesTitle": "Main service title, e.g. 'Website Design & Development'",
  "servicesDescription": "Brief description of services in 1-2 sentences",
  "includesList": "Bulleted list of everything included, one item per line starting with •"
}

Format payment1 and payment2 as readable strings like '$4,500 - 50% deposit due upon contract signing'.
Format includesList as a clean bullet list.`
        },
        {
          role: "user",
          content: proposalText
        }
      ],
      response_format: { type: "json_object" }
    });

    const extractedData = JSON.parse(completion.choices[0].message.content);

    console.log('AI extraction complete:', extractedData);

    res.json({
      success: true,
      data: extractedData
    });

  } catch (error) {
    console.error('AI parsing error:', error);
    res.status(500).json({
      success: false,
      error: error.message || 'Failed to parse proposal'
    });
  }
});

// Generate contract using DocuSeal template
app.post('/api/generate-contract', async (req, res) => {
  try {
    console.log('=== GENERATING CONTRACT FROM TEMPLATE ===');
    console.log('Data received:', JSON.stringify(req.body, null, 2));

    const {
      clientName,
      signerEmail,
      signerEmail2,
      effectiveDate,
      termEndDate,
      payment1,
      payment2,
      launchDate,
      servicesTitle,
      servicesDescription,
      includesList
    } = req.body;

    const docuseal = new DocusealApi({ key: process.env.DOCUSEAL_API_KEY });

    // Build submitters - always include Provider
    const submitters = [
      {
        role: 'First Party',
        email: signerEmail,
        name: clientName,
        // Prefill all the template fields with AI-extracted data
        values: {
          'Customer Name': clientName || '',
          'Payment1': payment1 || '',
          'Payment2': payment2 || '',
          'Launch Date': launchDate || termEndDate || '',
          'Services': servicesTitle || servicesDescription || '',
          'Includes': includesList || '',
          'Date1': effectiveDate || '',
          'Date2': termEndDate || ''
        }
      }
    ];

    console.log('Creating DocuSeal submission from template...');
    console.log('Template ID:', TEMPLATE_ID);
    console.log('Submitters:', JSON.stringify(submitters, null, 2));

    const submission = await docuseal.createSubmission({
      template_id: TEMPLATE_ID,
      send_email: true,
      submitters: submitters
    });

    console.log('Submission created successfully!');
    console.log('Submission ID:', submission.id || submission[0]?.submission_id);

    res.json({
      success: true,
      message: 'Contract sent successfully!',
      submissionId: submission.id || submission[0]?.submission_id,
      clientName: clientName
    });

  } catch (error) {
    console.error('Error generating contract:', error);
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', message: 'Server is running' });
});

// Start server
app.listen(PORT, () => {
  console.log(`\n🚀 WE GO OAKLAND Contract Generator (AI-Powered)`);
  console.log(`📱 Open your browser to: http://localhost:${PORT}`);
  console.log(`\n✨ Ready to generate contracts!\n`);
});
