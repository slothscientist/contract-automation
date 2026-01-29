#!/usr/bin/env node

require('dotenv').config();
const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel } = require('docx');
const { DocusealApi } = require('@docuseal/api');

// Parse command-line arguments
function parseArgs() {
  const args = process.argv.slice(2);
  const params = {};

  for (let i = 0; i < args.length; i += 2) {
    const key = args[i].replace('--', '');
    const value = args[i + 1];
    params[key] = value;
  }

  return params;
}

// Validate required parameters
function validateParams(params) {
  const required = [
    'clientName',
    'signerName',
    'signerEmail',
    'projectStartDate',
    'projectEndDate',
    'totalFee'
  ];

  const missing = required.filter(field => !params[field]);

  if (missing.length > 0) {
    console.error('❌ Missing required parameters:', missing.join(', '));
    console.log('\nUsage:');
    console.log('  node generate-contract.js \\');
    console.log('    --clientName "Acme Corp" \\');
    console.log('    --signerName "John Doe" \\');
    console.log('    --signerEmail "john@acmecorp.com" \\');
    console.log('    --projectStartDate "2026-02-01" \\');
    console.log('    --projectEndDate "2026-03-01" \\');
    console.log('    --totalFee "5000"');
    process.exit(1);
  }

  return true;
}

// Generate the service agreement document
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
          // Header - Company Name
          new Paragraph({
            text: "WE GO OAKLAND",
            heading: HeadingLevel.TITLE,
            alignment: AlignmentType.CENTER,
            spacing: {
              after: 200,
            },
          }),

          // Title
          new Paragraph({
            text: "SERVICE AGREEMENT",
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
            spacing: {
              before: 400,
              after: 400,
            },
          }),

          // Date
          new Paragraph({
            children: [
              new TextRun({
                text: `Date: ${new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })}`,
                bold: false,
              }),
            ],
            spacing: { after: 200 },
          }),

          // Parties Section
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
              new TextRun({
                text: "Service Provider: ",
                bold: true,
              }),
              new TextRun({
                text: "WE GO OAKLAND",
              }),
            ],
            spacing: { after: 100 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "Email: ",
                bold: true,
              }),
              new TextRun({
                text: "salaams@wegooakland.com",
              }),
            ],
            spacing: { after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "Client: ",
                bold: true,
              }),
              new TextRun({
                text: clientName,
              }),
            ],
            spacing: { after: 100 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "Signer: ",
                bold: true,
              }),
              new TextRun({
                text: signerName,
              }),
            ],
            spacing: { after: 100 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "Email: ",
                bold: true,
              }),
              new TextRun({
                text: signerEmail,
              }),
            ],
            spacing: { after: 400 },
          }),

          // Project Details
          new Paragraph({
            text: "PROJECT DETAILS",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "Project Duration: ",
                bold: true,
              }),
              new TextRun({
                text: `${projectStartDate} to ${projectEndDate}`,
              }),
            ],
            spacing: { after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "Total Fee: ",
                bold: true,
              }),
              new TextRun({
                text: `$${parseFloat(totalFee).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
              }),
            ],
            spacing: { after: 400 },
          }),

          // Terms and Conditions
          new Paragraph({
            text: "TERMS AND CONDITIONS",
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "1. Services: ",
                bold: true,
              }),
              new TextRun({
                text: "The Service Provider agrees to provide professional services as outlined in the project scope for the duration specified above.",
              }),
            ],
            spacing: { after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "2. Payment Terms: ",
                bold: true,
              }),
              new TextRun({
                text: "The Client agrees to pay the total fee specified above. Payment schedule and terms will be agreed upon separately.",
              }),
            ],
            spacing: { after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "3. Confidentiality: ",
                bold: true,
              }),
              new TextRun({
                text: "Both parties agree to maintain confidentiality of any proprietary information shared during the course of this agreement.",
              }),
            ],
            spacing: { after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "4. Termination: ",
                bold: true,
              }),
              new TextRun({
                text: "Either party may terminate this agreement with 14 days written notice.",
              }),
            ],
            spacing: { after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "5. Intellectual Property: ",
                bold: true,
              }),
              new TextRun({
                text: "Upon full payment, all work product created specifically for this project will be owned by the Client.",
              }),
            ],
            spacing: { after: 400 },
          }),

          // Signatures Section
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

          // Service Provider Signature
          new Paragraph({
            children: [
              new TextRun({
                text: "Service Provider Signature:",
                bold: true,
              }),
            ],
            spacing: { after: 100 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "{{provider_signature}}",
                italics: true,
              }),
            ],
            spacing: { after: 100 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "Name: WE GO OAKLAND",
              }),
            ],
            spacing: { after: 100 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "Date: {{provider_date}}",
                italics: true,
              }),
            ],
            spacing: { after: 400 },
          }),

          // Client Signature
          new Paragraph({
            children: [
              new TextRun({
                text: "Client Signature:",
                bold: true,
              }),
            ],
            spacing: { after: 100 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "{{client_signature}}",
                italics: true,
              }),
            ],
            spacing: { after: 100 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: `Name: ${signerName}`,
              }),
            ],
            spacing: { after: 100 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "Date: {{client_date}}",
                italics: true,
              }),
            ],
            spacing: { after: 200 },
          }),
        ],
      },
    ],
  });

  return doc;
}

// Send document via DocuSeal
async function sendViaDocuSeal(filePath, params) {
  const apiKey = process.env.DOCUSEAL_API_KEY;

  if (!apiKey) {
    throw new Error('DOCUSEAL_API_KEY not found in environment variables');
  }

  console.log('🔑 API Key found:', apiKey.substring(0, 10) + '...');
  const docuseal = new DocusealApi({ key: apiKey });

  console.log('📤 Uploading document to DocuSeal...');
  console.log('   File:', filePath);
  console.log('   Client:', params.clientName);
  console.log('   Submitters:', [
    `Provider: ${process.env.PROVIDER_EMAIL || 'salaams@wegooakland.com'}`,
    `Client: ${params.signerEmail}`
  ].join(', '));

  try {
    // Read the document and encode as base64
    const fileBuffer = fs.readFileSync(filePath);
    const base64File = fileBuffer.toString('base64');
    console.log(`   File size: ${fileBuffer.length} bytes (${(base64File.length / 1024).toFixed(2)} KB base64)`);

    console.log('\n⏳ Sending to DocuSeal API...');

    // Create a submission directly from the DOCX file
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

    console.log('\n✅ Document sent for signing!');
    console.log('📧 Submission ID:', submission.id);
    console.log('📅 Created:', submission.created_at || 'N/A');
    console.log('⏰ Expires:', submission.expired_at || 'Never');

    // Debug: Log full submission response
    console.log('\n🔍 Full API Response:');
    console.log(JSON.stringify(submission, null, 2));

    // Get the submission URL from the submitters array
    if (submission.submitters && submission.submitters.length > 0) {
      console.log('\n🔗 Signing URLs:');
      submission.submitters.forEach((submitter) => {
        const url = submitter.slug ? `https://docuseal.com/s/${submitter.slug}` : 'Pending';
        console.log(`   ${submitter.email}: ${url}`);
      });
    }

    return submission;
  } catch (error) {
    console.error('\n❌ DocuSeal API Error:', error.message);
    console.error('Error type:', error.constructor.name);

    if (error.response) {
      console.error('Response status:', error.response.status);
      console.error('Response statusText:', error.response.statusText);
      console.error('Response data:', JSON.stringify(error.response.data || error.response, null, 2));
    }

    if (error.request) {
      console.error('Request was made but no response received');
      console.error('Request details:', error.request);
    }

    if (error.message.includes('401') || error.message.includes('403')) {
      console.error('\n💡 Authentication Error - Your API key may be invalid or expired.');
      console.error('Please verify your API key at: https://docuseal.com/settings/api');
    }

    throw error;
  }
}

// Main execution
async function main() {
  try {
    console.log('🚀 Starting contract generation...\n');

    // Parse and validate parameters
    const params = parseArgs();
    validateParams(params);

    console.log('📋 Contract Details:');
    console.log(`   Client: ${params.clientName}`);
    console.log(`   Signer: ${params.signerName} (${params.signerEmail})`);
    console.log(`   Duration: ${params.projectStartDate} to ${params.projectEndDate}`);
    console.log(`   Fee: $${parseFloat(params.totalFee).toLocaleString('en-US')}\n`);

    // Generate the contract
    console.log('📝 Generating contract document...');
    const doc = await generateContract(params);

    // Save to file
    const outputDir = path.join(__dirname, 'output');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').split('T')[0];
    const fileName = `service-agreement-${params.clientName.replace(/\s+/g, '-').toLowerCase()}-${timestamp}.docx`;
    const filePath = path.join(outputDir, fileName);

    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(filePath, buffer);

    console.log(`✅ Contract saved: ${filePath}\n`);

    // Send via DocuSeal
    const result = await sendViaDocuSeal(filePath, params);

    console.log('\n✨ Process completed successfully!');
    console.log(`📄 Local file: ${filePath}`);
    console.log(`📧 Email sent to: ${params.signerEmail}`);

  } catch (error) {
    console.error('\n❌ Error:', error.message);
    process.exit(1);
  }
}

// Run the script
if (require.main === module) {
  main();
}

module.exports = { generateContract, sendViaDocuSeal };
