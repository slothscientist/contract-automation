#!/usr/bin/env node

require('dotenv').config();
const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel } = require('docx');
const { DocusealApi } = require('@docuseal/api');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// Generate contract document
async function generateContract(params) {
  const {
    clientName,
    signerName,
    signerEmail,
    projectStartDate,
    projectEndDate,
    totalFee
  } = params;

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            text: "WE GO OAKLAND",
            heading: HeadingLevel.TITLE,
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
          }),

          new Paragraph({
            text: "SERVICE AGREEMENT",
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
            spacing: { before: 400, after: 400 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: `Date: ${new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })}`,
              }),
            ],
            spacing: { after: 200 },
          }),

          new Paragraph({
            text: "PARTIES",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "This Service Agreement (\"Agreement\") is entered into between:",
              }),
            ],
            spacing: { after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({ text: "Service Provider: ", bold: true }),
              new TextRun({ text: "WE GO OAKLAND" }),
            ],
            spacing: { after: 100 },
          }),

          new Paragraph({
            children: [
              new TextRun({ text: "Email: ", bold: true }),
              new TextRun({ text: "salaams@wegooakland.com" }),
            ],
            spacing: { after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({ text: "Client: ", bold: true }),
              new TextRun({ text: clientName }),
            ],
            spacing: { after: 100 },
          }),

          new Paragraph({
            children: [
              new TextRun({ text: "Signer: ", bold: true }),
              new TextRun({ text: signerName }),
            ],
            spacing: { after: 100 },
          }),

          new Paragraph({
            children: [
              new TextRun({ text: "Email: ", bold: true }),
              new TextRun({ text: signerEmail }),
            ],
            spacing: { after: 400 },
          }),

          new Paragraph({
            text: "PROJECT DETAILS",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({ text: "Project Duration: ", bold: true }),
              new TextRun({ text: `${projectStartDate} to ${projectEndDate}` }),
            ],
            spacing: { after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({ text: "Total Fee: ", bold: true }),
              new TextRun({ text: `$${parseFloat(totalFee).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` }),
            ],
            spacing: { after: 400 },
          }),

          new Paragraph({
            text: "TERMS AND CONDITIONS",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({ text: "1. Services: ", bold: true }),
              new TextRun({ text: "The Service Provider agrees to provide professional services as outlined in the project scope for the duration specified above." }),
            ],
            spacing: { after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({ text: "2. Payment Terms: ", bold: true }),
              new TextRun({ text: "The Client agrees to pay the total fee specified above. Payment schedule and terms will be agreed upon separately." }),
            ],
            spacing: { after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({ text: "3. Confidentiality: ", bold: true }),
              new TextRun({ text: "Both parties agree to maintain confidentiality of any proprietary information shared during the course of this agreement." }),
            ],
            spacing: { after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({ text: "4. Termination: ", bold: true }),
              new TextRun({ text: "Either party may terminate this agreement with 14 days written notice." }),
            ],
            spacing: { after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({ text: "5. Intellectual Property: ", bold: true }),
              new TextRun({ text: "Upon full payment, all work product created specifically for this project will be owned by the Client." }),
            ],
            spacing: { after: 400 },
          }),

          new Paragraph({
            text: "SIGNATURES",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 600, after: 300 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "By signing below, both parties agree to the terms and conditions outlined in this Service Agreement.",
              }),
            ],
            spacing: { after: 400 },
          }),

          new Paragraph({
            children: [new TextRun({ text: "Service Provider Signature:", bold: true })],
            spacing: { after: 100 },
          }),

          new Paragraph({
            children: [new TextRun({ text: "{{s:Provider}}" })],
            spacing: { after: 100 },
          }),

          new Paragraph({
            children: [new TextRun({ text: "Name: WE GO OAKLAND" })],
            spacing: { after: 100 },
          }),

          new Paragraph({
            children: [new TextRun({ text: "Date: {{d:Provider}}" })],
            spacing: { after: 400 },
          }),

          new Paragraph({
            children: [new TextRun({ text: "Client Signature:", bold: true })],
            spacing: { after: 100 },
          }),

          new Paragraph({
            children: [new TextRun({ text: "{{s:Client}}" })],
            spacing: { after: 100 },
          }),

          new Paragraph({
            children: [new TextRun({ text: `Name: ${signerName}` })],
            spacing: { after: 100 },
          }),

          new Paragraph({
            children: [new TextRun({ text: "{{d:Client}}" })],
            spacing: { after: 200 },
          }),
        ],
      },
    ],
  });

  return doc;
}

// Send to DocuSeal
async function sendViaDocuSeal(filePath, params) {
  const apiKey = process.env.DOCUSEAL_API_KEY;

  if (!apiKey) {
    throw new Error('DOCUSEAL_API_KEY not found in environment variables');
  }

  const docuseal = new DocusealApi({ key: apiKey });

  const fileBuffer = fs.readFileSync(filePath);
  const base64File = fileBuffer.toString('base64');

  const submission = await docuseal.createSubmissionFromDocx({
    name: `Service Agreement - ${params.clientName}`,
    send_email: true,
    order: 'preserved',
    documents: [
      {
        name: 'service_agreement.docx',
        file: base64File,
      },
    ],
    submitters: [
      {
        role: 'Provider',
        email: process.env.PROVIDER_EMAIL || 'salaams@wegooakland.com',
        name: 'WE GO OAKLAND',
      },
      {
        role: 'Client',
        email: params.signerEmail,
        name: params.signerName,
      },
    ],
  });

  return submission;
}

// API endpoint to generate contract
app.post('/api/generate-contract', async (req, res) => {
  try {
    const { clientName, signerName, signerEmail, projectStartDate, projectEndDate, totalFee } = req.body;

    // Validate inputs
    if (!clientName || !signerName || !signerEmail || !projectStartDate || !projectEndDate || !totalFee) {
      return res.status(400).json({ error: 'All fields are required' });
    }

    console.log('Generating contract for:', clientName);

    // Generate contract
    const doc = await generateContract(req.body);

    // Save to file
    const outputDir = path.join(__dirname, 'output');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').split('T')[0];
    const fileName = `service-agreement-${clientName.replace(/\s+/g, '-').toLowerCase()}-${timestamp}.docx`;
    const filePath = path.join(outputDir, fileName);

    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(filePath, buffer);

    // Send via DocuSeal
    const submission = await sendViaDocuSeal(filePath, req.body);

    console.log('Contract sent successfully!');

    res.json({
      success: true,
      message: 'Contract generated and sent successfully!',
      submissionId: submission.id,
      fileName: fileName,
      signingUrls: submission.submitters?.map(s => ({
        email: s.email,
        url: s.slug ? `https://docuseal.com/s/${s.slug}` : null
      }))
    });

  } catch (error) {
    console.error('Error:', error);
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
  console.log(`\n🚀 WE GO OAKLAND Contract Generator`);
  console.log(`📱 Open your browser to: http://localhost:${PORT}`);
  console.log(`\n✨ Ready to generate contracts!\n`);
});
