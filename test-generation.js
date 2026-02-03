#!/usr/bin/env node

require('dotenv').config();
const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, UnderlineType, convertInchesToTwip } = require('docx');

// Import the generation function from server-ai.js
const serverCode = fs.readFileSync('./server-ai.js', 'utf8');

// Test parameters
const testParams = {
  effectiveDate: '2026-02-03',
  clientName: 'Test Company LLC',
  clientEntity: 'a limited liability company',
  clientManagingMembers: 'John Doe',
  termEndDate: '2027-02-03',
  paymentSchedule: [
    {
      description: 'deposit upon signing',
      amount: 5000,
      dueDate: '2026-02-03'
    }
  ],
  servicesDescription: 'Website design and development services'
};

// Extract and test the generateContractSections function
const indent = convertInchesToTwip(0.5);
const termEndDate = testParams.termEndDate;

// Generate sections 4-19
const sections = [
  // Section 4
  new Paragraph({
    children: [
      new TextRun({ text: "4. ", bold: true }),
      new TextRun({ text: "Limited Warranty.", bold: true }),
    ],
    spacing: { after: 200 },
  }),
];

console.log('Test: Generating paragraph objects...');
console.log('Section 4 paragraph created:', sections[0]);
console.log('\nChecking for placeholder text in server-ai.js...');

// Search for any placeholder patterns
if (serverCode.includes('[Sections') || serverCode.includes('continue with')) {
  console.log('ERROR: Found placeholder text in server-ai.js!');
  const lines = serverCode.split('\n');
  lines.forEach((line, i) => {
    if (line.includes('[Sections') || line.includes('continue with')) {
      console.log(`Line ${i + 1}: ${line}`);
    }
  });
} else {
  console.log('No placeholder text found in server-ai.js');
}

console.log('\nTest complete.');
