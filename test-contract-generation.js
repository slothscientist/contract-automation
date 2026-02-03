#!/usr/bin/env node

const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, UnderlineType, convertInchesToTwip } = require('docx');

// Copy the EXACT functions from server-ai.js
function generateWhereasClauses() {
  return [
    new Paragraph({
      children: [
        new TextRun({ text: "WHEREAS ", bold: true }),
        new TextRun({ text: "Service Provider has the capability and capacity to provide certain web design & marketing services; and" }),
      ],
      spacing: { after: 200 },
    }),
    new Paragraph({
      children: [
        new TextRun({ text: "WHEREAS ", bold: true }),
        new TextRun({ text: "Customer desires to retain Service Provider to provide the said services, and Service Provider is willing to perform such services under the terms and conditions hereinafter set forth;" }),
      ],
      spacing: { after: 200 },
    }),
    new Paragraph({
      children: [
        new TextRun({ text: "NOW, THEREFORE, ", bold: true }),
        new TextRun({ text: "in consideration of the mutual covenants and agreements hereinafter set forth and for other good and valuable consideration, the receipt and sufficiency of which are hereby acknowledged, Service Provider and Customer agree as follows:" }),
      ],
      spacing: { after: 400 },
    }),
  ];
}

function generateContractSections(termEndDate) {
  const indent = convertInchesToTwip(0.5);

  return [
    // Section 1: Services
    new Paragraph({
      children: [
        new TextRun({ text: "1. ", bold: true }),
        new TextRun({ text: "Services. ", bold: true}),
        new TextRun({ text: 'Service Provider shall provide to Customer the services (the "Services") set out in the statement of work (See Exhibit A, "Statement of Work"). Service Provider shall provide the Services in accordance with the terms and subject to the conditions set forth in this Agreement and the Statement of Work.' }),
      ],
      spacing: { after: 300 },
    }),

    // Section 2: Customer Obligations
    new Paragraph({
      children: [
        new TextRun({ text: "2. ", bold: true }),
        new TextRun({ text: "Customer Obligations. ", bold: true }),
        new TextRun({ text: "Customer shall:" }),
      ],
      spacing: { after: 200 },
    }),
    new Paragraph({
      text: "Respond promptly to any reasonable requests from Service Provider for instructions, information, or approvals required by Service Provider to provide the Services.",
      spacing: { after: 100 },
      indent: { left: indent },
    }),

    // Section 3: Fees
    new Paragraph({
      children: [
        new TextRun({ text: "3. ", bold: true }),
        new TextRun({ text: "Fees and Expenses.", bold: true }),
      ],
      spacing: { after: 200 },
    }),
    new Paragraph({
      text: "In consideration of the provision of the Services by the Service Provider...",
      spacing: { after: 200 },
    }),

    // SECTION 4 - CRITICAL TEST
    new Paragraph({
      children: [
        new TextRun({ text: "4. ", bold: true }),
        new TextRun({ text: "Limited Warranty.", bold: true }),
      ],
      spacing: { after: 200 },
    }),
    new Paragraph({
      text: "Service Provider warrants that it shall perform the Services in a timely, workmanlike, and professional manner in accordance with generally recognized industry standards for similar services.",
      spacing: { after: 200 },
    }),
  ];
}

// Generate test contract
async function generateTestContract() {
  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: {
            top: convertInchesToTwip(1),
            right: convertInchesToTwip(1),
            bottom: convertInchesToTwip(1),
            left: convertInchesToTwip(1),
          },
        },
      },
      children: [
        new Paragraph({
          text: "Web Design & Marketing Service Agreement",
          heading: HeadingLevel.HEADING_1,
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
        }),
        ...generateWhereasClauses(),
        ...generateContractSections('2027-02-03'),
        new Paragraph({
          text: "=== TEST: If you see Section 4 above, the code works! ===",
          bold: true,
          spacing: { before: 400 },
        }),
      ],
    }],
  });

  const buffer = await Packer.toBuffer(doc);
  const filePath = path.join(__dirname, 'TEST-CONTRACT.docx');
  fs.writeFileSync(filePath, buffer);

  console.log('\n✅ TEST CONTRACT GENERATED!');
  console.log('File:', filePath);
  console.log('Size:', buffer.length, 'bytes');
  console.log('\n🔍 Please open TEST-CONTRACT.docx and check if Section 4 appears!\n');
}

generateTestContract().catch(console.error);
