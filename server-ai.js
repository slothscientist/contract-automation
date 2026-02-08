#!/usr/bin/env node

require('dotenv').config();
const express = require('express');
const cors = require('cors');
const path = require('path');
const { v4: uuidv4 } = require('uuid');
const OpenAI = require('openai');
const { Resend } = require('resend');
const db = require('./database');

const app = express();
const PORT = process.env.PORT || 3000;
const APP_URL = process.env.APP_URL || `http://localhost:${PORT}`;

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
const resend = new Resend(process.env.RESEND_API_KEY);

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static('public'));

// Serve contract page for /contract/:id
app.get('/contract/:id', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'contract.html'));
});

// ─── AI Proposal Parser ───────────────────────────────────────────────────────
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
          content: `You are an expert at extracting contract information from proposals. Extract the following and return as JSON:
{
  "clientName": "Full client company/business name",
  "clientEntity": "Entity type (e.g. 'a limited liability company') or empty string",
  "clientManagingMembers": "Comma-separated managing members/owners or empty string",
  "signerEmail": "Primary email address",
  "signerEmail2": "Second email if multiple signers, otherwise empty string",
  "effectiveDate": "Contract start date in MM/DD/YYYY format",
  "termEndDate": "Contract end date in MM/DD/YYYY format",
  "payment1": "First payment as a string like '$4,500 — 50% deposit due upon signing'",
  "payment2": "Second payment as a string like '$4,500 — 50% balance due at launch'",
  "servicesTitle": "Main service title e.g. 'Website Design & Development'",
  "servicesDescription": "1-2 sentence description of the overall project",
  "includesList": "Everything included, one item per line starting with •"
}
Format payments as clean readable strings. Make the includes list thorough and clear.`
        },
        { role: "user", content: proposalText }
      ],
      response_format: { type: "json_object" }
    });

    const data = JSON.parse(completion.choices[0].message.content);
    console.log('AI extraction complete:', data);
    res.json({ success: true, data });

  } catch (error) {
    console.error('AI parsing error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// ─── Generate Contract & Send Email ──────────────────────────────────────────
app.post('/api/generate-contract', async (req, res) => {
  try {
    console.log('=== GENERATING CONTRACT ===');

    const {
      clientName, clientEntity, clientManagingMembers,
      signerEmail, signerEmail2,
      effectiveDate, termEndDate,
      payment1, payment2,
      servicesTitle, servicesDescription, includesList
    } = req.body;

    // Create unique contract ID
    const contractId = uuidv4();

    // Save to database
    db.createContract({
      id: contractId,
      client_name: clientName,
      client_email: signerEmail,
      effective_date: effectiveDate,
      term_end_date: termEndDate,
      payment_1: payment1,
      payment_2: payment2,
      launch_date: termEndDate,
      services_title: servicesTitle,
      services_description: servicesDescription,
      includes_list: includesList
    });

    const contractUrl = `${APP_URL}/contract/${contractId}`;
    console.log('Contract created:', contractId);
    console.log('Contract URL:', contractUrl);

    // Send email to client
    await sendSigningEmail(signerEmail, clientName, contractUrl);

    // Send email to second signer if present
    if (signerEmail2) {
      await sendSigningEmail(signerEmail2, clientName, contractUrl);
    }

    res.json({
      success: true,
      message: 'Contract created and sent!',
      contractId,
      contractUrl
    });

  } catch (error) {
    console.error('Error generating contract:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// ─── Get Contract Data (for contract page) ───────────────────────────────────
app.get('/api/contract/:id', (req, res) => {
  try {
    const contract = db.getContract(req.params.id);
    if (!contract) {
      return res.status(404).json({ success: false, error: 'Contract not found' });
    }
    res.json({ success: true, contract });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// ─── Sign Contract ────────────────────────────────────────────────────────────
app.post('/api/contract/:id/sign', async (req, res) => {
  try {
    const { signerName, signatureData } = req.body;
    const contract = db.getContract(req.params.id);

    if (!contract) {
      return res.status(404).json({ success: false, error: 'Contract not found' });
    }
    if (contract.status === 'signed') {
      return res.status(400).json({ success: false, error: 'Already signed' });
    }

    db.signContract(req.params.id, signatureData, signerName);
    console.log(`✅ Contract signed by ${signerName} for ${contract.client_name}`);

    // Send confirmation emails
    await sendConfirmationEmail(contract.client_email, contract.client_name, signerName);
    await sendConfirmationEmail(
      process.env.PROVIDER_EMAIL,
      process.env.PROVIDER_NAME || 'WeGo! Oakland',
      signerName,
      contract.client_name
    );

    res.json({ success: true, message: 'Contract signed successfully!' });

  } catch (error) {
    console.error('Signing error:', error);
    res.status(500).json({ success: false, error: error.message });
  }
});

// ─── Dashboard (list all contracts) ─────────────────────────────────────────
app.get('/api/contracts', (req, res) => {
  try {
    const contracts = db.getAllContracts();
    res.json({ success: true, contracts });
  } catch (error) {
    res.status(500).json({ success: false, error: error.message });
  }
});

// ─── Email Functions ─────────────────────────────────────────────────────────
async function sendSigningEmail(toEmail, clientName, contractUrl) {
  try {
    await resend.emails.send({
      from: `WeGo! Oakland <${process.env.PROVIDER_EMAIL}>`,
      to: toEmail,
      subject: `Service Agreement Ready to Sign — WeGo! Oakland`,
      html: `
        <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; background: #faf7f2;">
          <div style="background: #000; padding: 30px 40px;">
            <h1 style="color: #fce81d; margin: 0; font-size: 24px;">WeGo! Oakland</h1>
          </div>
          <div style="padding: 40px;">
            <h2 style="color: #000; margin-bottom: 16px;">Your Service Agreement is Ready</h2>
            <p style="color: #555; line-height: 1.7; margin-bottom: 24px;">
              Hi ${clientName},<br><br>
              Your Web Design & Marketing Service Agreement with WeGo! Oakland is ready for your review and signature.
              Please click the button below to view and sign your agreement.
            </p>
            <div style="text-align: center; margin: 32px 0;">
              <a href="${contractUrl}" style="background: #e9007f; color: white; padding: 16px 40px; border-radius: 8px; text-decoration: none; font-weight: 700; font-size: 16px;">
                Review & Sign Agreement →
              </a>
            </div>
            <p style="color: #999; font-size: 13px; margin-top: 32px;">
              If the button doesn't work, copy this link: ${contractUrl}
            </p>
          </div>
          <div style="background: #000; padding: 20px 40px; text-align: center;">
            <p style="color: #666; font-size: 12px; margin: 0;">© WeGo! Oakland · salaams@wegooakland.com</p>
          </div>
        </div>
      `
    });
    console.log(`📧 Signing email sent to ${toEmail}`);
  } catch (error) {
    console.error('Email send error:', error);
  }
}

async function sendConfirmationEmail(toEmail, toName, signerName, clientName) {
  const isProvider = !!clientName;
  const subject = isProvider
    ? `✅ Contract Signed — ${clientName}`
    : `✅ Agreement Signed — WeGo! Oakland`;

  try {
    await resend.emails.send({
      from: `WeGo! Oakland <${process.env.PROVIDER_EMAIL}>`,
      to: toEmail,
      subject,
      html: `
        <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; background: #faf7f2;">
          <div style="background: #000; padding: 30px 40px;">
            <h1 style="color: #fce81d; margin: 0; font-size: 24px;">WeGo! Oakland</h1>
          </div>
          <div style="padding: 40px; text-align: center;">
            <div style="font-size: 48px; margin-bottom: 16px;">✅</div>
            <h2 style="color: #000; margin-bottom: 16px;">Agreement Signed!</h2>
            <p style="color: #555; line-height: 1.7;">
              ${isProvider
                ? `<strong>${signerName}</strong> has signed the service agreement for <strong>${clientName}</strong>.`
                : `Your service agreement with WeGo! Oakland has been signed by <strong>${signerName}</strong>.`
              }
            </p>
            <p style="color: #999; font-size: 13px; margin-top: 32px;">
              WeGo! Oakland · salaams@wegooakland.com
            </p>
          </div>
        </div>
      `
    });
    console.log(`📧 Confirmation email sent to ${toEmail}`);
  } catch (error) {
    console.error('Confirmation email error:', error);
  }
}

// ─── Start Server ─────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\n🚀 WeGo! Oakland Contract System`);
  console.log(`📱 Open: http://localhost:${PORT}`);
  console.log(`\n✨ Ready!\n`);
});
