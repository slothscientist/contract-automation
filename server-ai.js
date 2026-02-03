#!/usr/bin/env node

require('dotenv').config();
const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const OpenAI = require('openai');
const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, UnderlineType, convertInchesToTwip } = require('docx');
const { DocusealApi } = require('@docuseal/api');

const app = express();
const PORT = process.env.PORT || 3000;

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
  "clientAddress": "Full address if mentioned",
  "signerEmail": "Primary email address",
  "signerEmail2": "Second email if multiple signers",
  "effectiveDate": "Contract start date in YYYY-MM-DD format",
  "termEndDate": "Contract end date in YYYY-MM-DD format",
  "totalProjectFee": "Total project fee as number",
  "paymentSchedule": [
    {
      "description": "Payment description",
      "amount": "Amount as number",
      "dueDate": "Due date in YYYY-MM-DD format or readable format"
    }
  ],
  "services": [
    {
      "title": "Service category title (e.g., Website Design & Development)",
      "includes": ["List of included items"],
      "notIncluded": ["List of excluded items (if mentioned)"]
    }
  ],
  "projectPhases": [
    {
      "phase": "Phase name",
      "description": "What happens in this phase",
      "duration": "Duration if mentioned"
    }
  ],
  "revisionPolicy": "Revision/feedback policy if mentioned",
  "importantNotes": ["Any important notices or warnings"]
}

Be thorough in extracting service details, phases, deliverables, exclusions, and timelines.`
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

// Generate complete WeGo Oakland contract
function generateFullContract(params) {
  const {
    effectiveDate,
    clientName,
    clientEntity,
    clientManagingMembers,
    termEndDate,
    servicesDescription,
    paymentSchedule
  } = params;

  const members = clientManagingMembers ? clientManagingMembers.split(',').map(m => m.trim()) : [];
  const hasMultipleSigners = members.length > 1;

  return new Document({
    sections: [
      {
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
          // Title
          new Paragraph({
            text: "Web Design & Marketing Service Agreement",
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
          }),

          // Opening - with address field for customer to fill
          new Paragraph({
            children: [
              new TextRun({
                text: `This Web Design & Marketing Services Agreement (this "Agreement"), dated as of ${effectiveDate} (the "Effective Date"), is by and between WeGo! Oakland, located at 1140 Everett Avenue, Oakland, CA 94602 ("Service Provider") and ${clientName}${clientEntity ? `, ${clientEntity}` : ''}${members.length > 0 ? `, with managing members ${clientManagingMembers}` : ''}, located at `,
              }),
              new TextRun({
                text: "{{t:Customer1:Address}}",
                bold: true,
                color: "0000FF"
              }),
              new TextRun({
                text: ` ("Customer" and together with Service Provider, the "Parties", and each a "Party").`,
              }),
            ],
            spacing: { after: 300 },
          }),

          // WHEREAS clauses
          ...generateWhereasClauses(),

          // Main contract sections
          ...generateContractSections(termEndDate),

          // Signature section
          ...generateSignatureSection(clientName, members, hasMultipleSigners),

          // Exhibit A
          ...generateExhibitA(params),
        ],
      },
    ],
  });
}

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
    new Paragraph({
      text: "Cooperate with Service Provider in its performance of the Services and respond in a timely manner to requests from Service Provider, as outlined Exhibit A attached hereto.",
      spacing: { after: 100 },
      indent: { left: indent },
    }),
    new Paragraph({
      text: "Take all steps necessary to prevent Customer-caused delays in Service Provider's provision of the Services.",
      spacing: { after: 300 },
      indent: { left: indent },
    }),

    // Section 3: Fees and Expenses
    new Paragraph({
      children: [
        new TextRun({ text: "3. ", bold: true }),
        new TextRun({ text: "Fees and Expenses.", bold: true }),
      ],
      spacing: { after: 200 },
    }),
    new Paragraph({
      text: "In consideration of the provision of the Services by the Service Provider and the rights granted to Customer under this Agreement, Customer shall pay the fees set out in Service Provider's then current fee schedule, as published/attached as Exhibit A.",
      spacing: { after: 200 },
    }),
    new Paragraph({
      text: "As described in Exhibit A attached hereto, the fees to be paid pursuant to ‎3.1 shall include a non-refundable deposit that shall be paid by Customer.",
      spacing: { after: 200 },
    }),
    new Paragraph({
      text: "Customer shall reimburse Service Provider for out-of-pocket expenses incurred by Service Provider for reasonable and necessary expenses incurred in connection with the performance of the Services provided that such expenses have been pre-approved by Customer. Customer shall pay all invoices for approved out-of-pocket expenses incurred by Service Provider within Seven (7) days of receipt by the Customer of an invoice from Service Provider accompanied by receipts and reasonable supporting documentation.",
      spacing: { after: 200 },
    }),
    new Paragraph({
      text: "Customer shall be responsible for all sales, use and excise taxes, and any other similar taxes, duties and charges of any kind imposed by any federal, state or local governmental entity on any amounts payable by Customer hereunder; provided that in no event shall Customer pay or be responsible for any taxes imposed on, or regarding, Service Provider's income, revenues, gross receipts, personnel, or real or personal property or other assets.",
      spacing: { after: 300 },
    }),

    // Section 4: Limited Warranty
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
    new Paragraph({
      text: "Service Provider's sole and exclusive liability and Customer's sole and exclusive remedy for breach of this limited warranty shall be reperformance of the affected services. If Service Provider cannot reperform the services in compliance with the warranty set forth above within a reasonable time (but no more than 30 days) after Customer's written notice of such breach, Customer may, at its option, terminate the Agreement by serving written notice of termination in accordance with Section 6.2. Service Provider shall within 30 days after the effective date of such termination, refund to Customer a portion of the fees previously paid by Customer as of the date of termination corresponding to the defective Services.",
      spacing: { after: 200 },
    }),
    new Paragraph({
      children: [
        new TextRun({ text: "SERVICE PROVIDER MAKES NO WARRANTIES EXCEPT FOR THAT PROVIDED IN SECTION ‎4.1, ABOVE, ALL OTHER WARRANTIES, EXPRESS AND IMPLIED, ARE EXPRESSLY DISCLAIMED. PERSONAL DISSATISFACTION, SUBJECTIVE PREFERENCES, OR MATTERS OF TASTE SHALL NOT CONSTITUTE A BREACH OF THESE WARRANTIES.", bold: true }),
      ],
      spacing: { after: 300 },
    }),

    // Section 5: Intellectual Property
    new Paragraph({
      children: [
        new TextRun({ text: "5. ", bold: true }),
        new TextRun({ text: "Intellectual Property. ", bold: true }),
        new TextRun({ text: 'Upon final payment, all intellectual property rights, including copyrights, patents, patent disclosures and inventions (whether patentable or not), trademarks, service marks, trade dress, trade names, logos, corporate names and domain names, and all other rights (collectively, "Intellectual Property Rights") in and to all customized documents, work product, and other materials specifically created for the Customer under this Agreement (the "Deliverables") shall be transferred to and owned by the Customer. However, Service Provider retains all Intellectual Property Rights in and to any underlying templates, code, standard layouts, general design elements, and pre-existing materials used in the course of performing the Services, and grants the Customer a non-exclusive, worldwide, fully paid-up, royalty-free license to use such materials as part of the Deliverables. This license is for the sole purpose of enabling the Customer to make reasonable use of the Deliverables as delivered in the context of the Services provided.' }),
      ],
      spacing: { after: 200 },
    }),
    new Paragraph({
      text: "To Customer's knowledge, intellectual property it provides to the Service Provider in connection with the Agreement does not infringe, misappropriate, or otherwise violate in any manner the intellectual property of any third party.",
      spacing: { after: 200 },
    }),
    new Paragraph({
      text: "To the Service Provider's knowledge, the designs provided to Customer under this Agreement do not infringe, misappropriate, or otherwise violate in any manner the intellectual property of any third party.",
      spacing: { after: 200 },
    }),
    new Paragraph({
      text: "Customer acknowledges that it is solely responsible for conducting any necessary trademark searches, obtaining legal clearance, and securing trademark registrations for its logo design. Service Provider shall not be liable for trademark infringement claims related to the use or registration of the Customer's logo, and Customer agrees to indemnify, hold harmless and promptly reimburse Service Provider for or in respect of any damages, losses or liabilities incurred by Service Provider arising from or in connection with any third party claim for infringement of such third party's intellectual property related to the Customer's logo or trademarks.",
      spacing: { after: 300 },
    }),

    // Section 6: Term, Termination, and Survival
    new Paragraph({
      children: [
        new TextRun({ text: "6. ", bold: true }),
        new TextRun({ text: "Term, Termination, and Survival.", bold: true }),
      ],
      spacing: { after: 200 },
    }),
    new Paragraph({
      children: [
        new TextRun({ text: `This Agreement shall commence as of the Effective Date and shall continue thereafter until ${termEndDate} (the "Term") unless sooner terminated pursuant to Sections ‎6.2 or ‎6.3 or unless both parties mutually agree to extend the Term.` }),
      ],
      spacing: { after: 200 },
    }),
    new Paragraph({
      text: "Either Party may terminate this Agreement, effective upon written notice to the other Party (the \"Defaulting Party\") if the Defaulting Party:",
      spacing: { after: 200 },
    }),
    new Paragraph({
      text: "Materially breaches this Agreement, and the Defaulting Party does not cure such material breach within [15] days after receipt of written notice of such breach, or such material breach is incapable of cure. The Parties acknowledge that design, creative, or professional services involve subjective judgment. Customer dissatisfaction with the aesthetics, style, or subjective qualities of the work product shall not constitute a breach of this Agreement, provided that the work substantially complies with the agreed-upon specifications, scope, and industry standard. Customer's sole remedy for dissatisfaction shall be limited to the revision and feedback schedule outlined in Exhibit A.",
      spacing: { after: 100 },
      indent: { left: indent },
    }),
    new Paragraph({
      text: "Becomes insolvent or admits its inability to pay its debts generally as they become due.",
      spacing: { after: 100 },
      indent: { left: indent },
    }),
    new Paragraph({
      text: "Becomes subject, voluntarily or involuntarily, to any proceeding under any domestic or foreign bankruptcy or insolvency law, which is not fully stayed within seven business days or is not dismissed or vacated within 45 days after filing.",
      spacing: { after: 100 },
      indent: { left: indent },
    }),
    new Paragraph({
      text: "Is dissolved or liquidated or takes any corporate action for such purpose.",
      spacing: { after: 100 },
      indent: { left: indent },
    }),
    new Paragraph({
      text: "Makes a general assignment for the benefit of creditors.",
      spacing: { after: 100 },
      indent: { left: indent },
    }),
    new Paragraph({
      text: "Has a receiver, trustee, custodian, or similar agent appointed by order of any court of competent jurisdiction to take charge of or sell any material portion of its property or business.",
      spacing: { after: 200 },
      indent: { left: indent },
    }),
    new Paragraph({
      text: "Notwithstanding anything to the contrary in this section, Service Provider may terminate this Agreement before the expiration date of the Term on written notice if Customer fails to pay any amount when due hereunder and such failure continues for 15 days after Customer's receipt of written notice of nonpayment.",
      spacing: { after: 200 },
    }),
    new Paragraph({
      text: "The rights and obligations of the Parties set forth in this Section ‎6.4 and in Sections 3.2, 4, [5.2, 5.3,] 7, 13, 14, 15, 16, 17, and 19 and any right or obligation of the Parties in this Agreement which, by its nature, should survive termination or expiration of this Agreement, will survive any such termination or expiration of this Agreement.",
      spacing: { after: 300 },
    }),

    // Section 7: Limitation of Liability
    new Paragraph({
      children: [
        new TextRun({ text: "7. ", bold: true }),
        new TextRun({ text: "Limitation of Liability.", bold: true }),
      ],
      spacing: { after: 200 },
    }),
    new Paragraph({
      children: [
        new TextRun({ text: "IN NO EVENT SHALL SERVICE PROVIDER BE LIABLE TO CUSTOMER OR TO ANY THIRD PARTY FOR ANY LOSS OF USE, REVENUE, OR PROFIT OR LOSS OF DATA OR DIMINUTION IN VALUE, OR FOR ANY CONSEQUENTIAL, INCIDENTAL, INDIRECT, EXEMPLARY, SPECIAL, OR PUNITIVE DAMAGES WHETHER ARISING OUT OF BREACH OF CONTRACT, TORT (INCLUDING NEGLIGENCE), OR OTHERWISE, REGARDLESS OF WHETHER SUCH DAMAGE WAS FORESEEABLE AND WHETHER OR NOT SERVICE PROVIDER HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES, AND NOTWITHSTANDING THE FAILURE OF ANY AGREED OR OTHER REMEDY OF ITS ESSENTIAL PURPOSE.", bold: true }),
      ],
      spacing: { after: 200 },
    }),
    new Paragraph({
      children: [
        new TextRun({ text: "IN NO EVENT SHALL SERVICE PROVIDER'S AGGREGATE LIABILITY ARISING OUT OF OR RELATED TO THIS AGREEMENT, WHETHER ARISING OUT OF OR RELATED TO BREACH OF CONTRACT, TORT (INCLUDING NEGLIGENCE), OR OTHERWISE, EXCEED THE AGGREGATE AMOUNTS PAID OR PAYABLE TO SERVICE PROVIDER PURSUANT TO THIS AGREEMENT/IN THE [12-MONTH] PERIOD PRECEDING THE EVENT GIVING RISE TO THE CLAIM.", bold: true }),
      ],
      spacing: { after: 300 },
    }),

    // Section 8: Entire Agreement
    new Paragraph({
      children: [
        new TextRun({ text: "8. ", bold: true }),
        new TextRun({ text: "Entire Agreement. ", bold: true }),
        new TextRun({ text: "This Agreement constitutes the sole and entire agreement of the Parties with respect to the subject matter contained herein, and supersedes all prior and contemporaneous understandings, agreements, representations and warranties, both written and oral, regarding such subject matter." }),
      ],
      spacing: { after: 300 },
    }),

    // Section 9: Severability
    new Paragraph({
      children: [
        new TextRun({ text: "9. ", bold: true }),
        new TextRun({ text: "Severability. ", bold: true }),
        new TextRun({ text: "If any term or provision of this Agreement is found by a court of competent jurisdiction to be invalid, illegal or unenforceable in any jurisdiction, such invalidity, illegality or unenforceability shall not affect any other term or provision of this Agreement or invalidate or render unenforceable such term or provision in any other jurisdiction. Upon a determination that any term or provision is invalid, illegal or unenforceable, the Parties shall negotiate in good faith to modify this Agreement to effect the original intent of the Parties as closely as possible in order that the transactions contemplated hereby be consummated as originally contemplated to the greatest extent possible." }),
      ],
      spacing: { after: 300 },
    }),

    // Section 10: Amendments
    new Paragraph({
      children: [
        new TextRun({ text: "10. ", bold: true }),
        new TextRun({ text: "Amendments. ", bold: true }),
        new TextRun({ text: "No amendment to or modification of this Agreement is effective unless it is in writing and signed by each Party." }),
      ],
      spacing: { after: 300 },
    }),

    // Section 11: Waiver
    new Paragraph({
      children: [
        new TextRun({ text: "11. ", bold: true }),
        new TextRun({ text: "Waiver. ", bold: true }),
        new TextRun({ text: "No waiver by any Party of any of the provisions of this Agreement shall be effective unless explicitly set forth in writing and signed by the Party so waiving. Except as otherwise set forth in this Agreement, no failure to exercise, or delay in exercising, any Right operates as a waiver thereof. No single or partial exercise of any Right precludes any other or further exercise thereof or the exercise of any other Right. The Rights under this Agreement are cumulative and are in addition to any other rights and remedies available at law or in equity or otherwise; provided that, the parties intend that the remedy set out in ‎‎Section 4 (Limited Warranty) is Customer's exclusive remedy for the Service Provider's breach of the limited warranty set out in ‎‎Section 4." }),
      ],
      spacing: { after: 300 },
    }),

    // Section 12: Assignment
    new Paragraph({
      children: [
        new TextRun({ text: "12. ", bold: true }),
        new TextRun({ text: "Assignment. ", bold: true }),
        new TextRun({ text: "Customer shall not assign, transfer, delegate or subcontract any of its rights or delegate any of its obligations under this Agreement without the prior written consent of Service Provider. Any purported assignment or delegation in violation of this ‎Section 12 shall be null and void. No assignment or delegation shall relieve the Customer of any of its obligations under this Agreement. Service Provider may assign any of its rights or delegate any of its obligations to any affiliate or to any person acquiring all or substantially all of Service Provider's assets without Customer's consent." }),
      ],
      spacing: { after: 300 },
    }),

    // Section 13: Successors and Assigns
    new Paragraph({
      children: [
        new TextRun({ text: "13. ", bold: true }),
        new TextRun({ text: "Successors and Assigns. ", bold: true }),
        new TextRun({ text: "This Agreement is binding on and inures to the benefit of the Parties to this Agreement and their respective permitted successors and permitted assigns." }),
      ],
      spacing: { after: 300 },
    }),

    // Section 14: Relationship of the Parties
    new Paragraph({
      children: [
        new TextRun({ text: "14. ", bold: true }),
        new TextRun({ text: "Relationship of the Parties. ", bold: true }),
        new TextRun({ text: "The relationship between the Parties is that of independent contractors. The details of the method and manner for performance of the Services by Service Provider shall be under its own control, Customer being interested only in the results thereof. The Service Provider shall be solely responsible for supervising, controlling and directing the details and manner of the completion of the Services. Nothing in this Agreement shall give the Customer the right to instruct, supervise, control, or direct the details and manner of the completion of the Services. The Services must meet the Customer's final approval and shall be subject to the Customer's general right of inspection throughout the performance of the Services and to secure satisfactory final completion. Nothing contained in this Agreement shall be construed as creating any agency, partnership, joint venture or other form of joint enterprise, employment or fiduciary relationship between the parties, and neither party shall have authority to contract for or bind the other party in any manner whatsoever." }),
      ],
      spacing: { after: 300 },
    }),

    // Section 15: No Third-Party Beneficiaries
    new Paragraph({
      children: [
        new TextRun({ text: "15. ", bold: true }),
        new TextRun({ text: "No Third-Party Beneficiaries. ", bold: true }),
        new TextRun({ text: "This Agreement benefits solely the Parties to this Agreement and their respective permitted successors and assigns and nothing in this Agreement, express or implied, confers on any other Person any legal or equitable right, benefit or remedy of any nature whatsoever under or by reason of this Agreement." }),
      ],
      spacing: { after: 300 },
    }),

    // Section 16: Governing Law
    new Paragraph({
      children: [
        new TextRun({ text: "16. ", bold: true }),
        new TextRun({ text: "Governing Law. ", bold: true }),
        new TextRun({ text: "This Agreement and all related documents, including all exhibits attached hereto, and all matters arising out of or relating to this Agreement, whether sounding in contract, tort, or statute are governed by, and construed in accordance with, the laws of the State of California, United States of America, without giving effect to the conflict of laws provisions thereof to the extent such principles or rules would require or permit the application of the laws of any jurisdiction other than those of the State of California." }),
      ],
      spacing: { after: 300 },
    }),

    // Section 17: ARBITRATION
    new Paragraph({
      children: [
        new TextRun({ text: "17. ", bold: true }),
        new TextRun({ text: "ARBITRATION. ", bold: true }),
        new TextRun({ text: "IN THE EVENT A DISPUTE ARISES UNDER THIS AGREEMENT, SUCH DISPUTE SHALL BE SUBMITTED TO ARBITRATION AND RESOLVED BY A SINGLE ARBITRATOR IN ACCORDANCE WITH THE COMMERCIAL ARBITRATION RULES OF THE AMERICAN ARBITRATION ASSOCIATION (\"AAA\"). THE ARBITRATOR SHALL BE SELECTED BY MUTUAL AGREEMENT OF THE PARTIES. ALL SUCH ARBITRATION SHALL BE CONFIDENTIAL AND BINDING AND SHALL BE CONDUCTED VIRTUALLY THROUGH THE AAA-ICDR® VIRTUAL ARBITRATION PROCESS. THE FEES OF ANY ARBITRATION SHALL BE BORNE EQUALLY BY CUSTOMER AND SERVICE PROVIDER. ANY DECISION OR AWARD AS A RESULT OF ANY SUCH ARBITRATION PROCEEDING SHALL BE IN WRITING AND SHALL PROVIDE AN EXPLANATION FOR ALL CONCLUSIONS OF LAW AND FACT AND SHALL INCLUDE THE ASSESSMENT OF COSTS, EXPENSES, AND REASONABLE ATTORNEYS' FEES. JUDGMENT UPON THE ARBITRAL AWARD MAY BE ENTERED IN ANY COURT OF COMPETENT JURISDICTION, OR APPLICATION MAY BE MADE TO SUCH COURT FOR A JUDICIAL RECOGNITION OF THE AWARD OR ANY ORDER OF ENFORCEMENT THEREOF. TO THE EXTENT JUDICIAL ACTION IN SUPPORT OF ARBITRATION IS NECESSARY, THE PARTIES HERETO HEREBY IRREVOCABLY SUBMIT TO THE EXCLUSIVE JURISDICTION OF THE STATE COURTS OF CALIFORNIA AND THE PARTIES HERETO IRREVOCABLY AGREE THAT ALL CLAIMS WITH RESPECT TO SUCH PROCEEDING SHALL BE HEARD AND DETERMINED IN SUCH STATE COURTS OF CALIFORNIA.", bold: true }),
      ],
      spacing: { after: 300 },
    }),

    // Section 18: Counterparts
    new Paragraph({
      children: [
        new TextRun({ text: "18. ", bold: true }),
        new TextRun({ text: "Counterparts. ", bold: true }),
        new TextRun({ text: "This Agreement may be executed in counterparts, each of which is deemed an original, but all of which together are deemed to be one and the same agreement. Notwithstanding anything to the contrary, a signed copy of this Agreement delivered by email or other means of electronic transmission is deemed to have the same legal effect as delivery of an original signed copy of this Agreement." }),
      ],
      spacing: { after: 300 },
    }),

    // Section 19: Force Majeure
    new Paragraph({
      children: [
        new TextRun({ text: "19. ", bold: true }),
        new TextRun({ text: "Force Majeure. ", bold: true }),
        new TextRun({ text: "No Party shall be liable or responsible to the other Party, or be deemed to have defaulted under or breached this Agreement, for any failure or delay in fulfilling or performing any term of this Agreement (except for any obligations of the Customer to make payments to Service Provider hereunder), when and to the extent such failure or delay is caused by or results from acts beyond the impacted party's (\"Impacted Party\") reasonable control, including, without limitation, the following force majeure events (\"Force Majeure Event(s)\"): (a) acts of God; (b) flood, fire, earthquake, epidemics or explosion; (c) war, invasion, hostilities (whether war is declared or not), terrorist threats or acts, riot or other civil unrest; (d) government order, law, or actions; (e) embargoes or blockades in effect on or after the date of this Agreement; (f) national or regional emergency; (g) strikes, labor stoppages or slowdowns, or other industrial disturbances; (h) telecommunication breakdowns, power outages or shortages, lack of warehouse or storage space, inadequate transportation services, or inability or delay in obtaining supplies of adequate or suitable materials; and (i) other events beyond the control of the Impacted Party." }),
      ],
      spacing: { after: 200 },
    }),
    new Paragraph({
      text: "The Impacted Party shall give notice within 10 days of the Force Majeure Event to the other party, stating the period of time the occurrence is expected to continue. The Impacted Party shall use diligent efforts to end the failure or delay and ensure the effects of such Force Majeure Event are minimized. The Impacted Party shall resume the performance of its obligations as soon as reasonably practicable after the removal of the cause. In the event that the Impacted Party's failure or delay remains uncured for a period of 20 consecutive days following written notice given by it under this Section ‎19, the other Party may thereafter terminate this Agreement upon 10 days' written notice.",
      spacing: { after: 300 },
    }),
  ];
}

function generateSignatureSection(clientName, members, hasMultipleSigners) {
  const paragraphs = [
    new Paragraph({
      text: "[signature page follows]",
      italics: true,
      alignment: AlignmentType.CENTER,
      spacing: { before: 400, after: 600 },
    }),
    new Paragraph({
      text: "IN WITNESS WHEREOF, the parties hereto have caused this Agreement, including all terms set forth in Exhibit A, to be duly executed as of the Effective Date.",
      spacing: { before: 600, after: 400 },
      pageBreakBefore: true,
    }),
    new Paragraph({
      text: "CUSTOMER:",
      bold: true,
      spacing: { before: 600, after: 400 },
    }),
    new Paragraph({
      text: clientName,
      spacing: { after: 300 },
    }),
  ];

  // First signer
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: "By: " }),
        new TextRun({ text: "{{s:Customer1}}", bold: true, color: "0000FF" }),
      ],
      spacing: { after: 100 },
    }),
    new Paragraph({
      text: members.length > 0 ? `Name: ${members[0]}` : "Name: {{t:Customer1:Name}}",
      spacing: { after: 100 },
    }),
    new Paragraph({
      text: "Managing Member",
      spacing: { after: 400 },
    })
  );

  // Second signer if applicable
  if (hasMultipleSigners) {
    paragraphs.push(
      new Paragraph({
        children: [
          new TextRun({ text: "By: " }),
          new TextRun({ text: "{{s:Customer2}}", bold: true, color: "0000FF" }),
        ],
        spacing: { after: 100 },
      }),
      new Paragraph({
        text: `Name: ${members[1]}`,
        spacing: { after: 100 },
      }),
      new Paragraph({
        text: "Managing Member",
        spacing: { after: 600 },
      })
    );
  }

  // Service Provider signature
  paragraphs.push(
    new Paragraph({
      text: "SERVICE PROVIDER:",
      bold: true,
      spacing: { before: 600, after: 400 },
    }),
    new Paragraph({
      children: [
        new TextRun({ text: "By: " }),
        new TextRun({ text: "{{s:Provider}}", bold: true, color: "0000FF" }),
      ],
      spacing: { after: 100 },
    }),
    new Paragraph({
      text: "Name: Salaams J. deRosa",
      spacing: { after: 100 },
    }),
    new Paragraph({
      text: "WeGo! Oakland",
      spacing: { after: 400 },
    })
  );

  return paragraphs;
}

function generateExhibitA(params) {
  const { clientName, paymentSchedule, services, projectPhases, revisionPolicy, importantNotes, servicesDescription } = params;
  const members = params.clientManagingMembers ? params.clientManagingMembers.split(',').map(m => m.trim()) : [];
  const initialsField = members.length > 1 ? "{{i:Customer1}} {{i:Customer2}}" : "{{i:Customer1}}";

  const paragraphs = [
    new Paragraph({
      text: "[Exhibit A follows]",
      italics: true,
      alignment: AlignmentType.CENTER,
      spacing: { before: 400, after: 600 },
    }),
    new Paragraph({
      text: "EXHIBIT A",
      heading: HeadingLevel.HEADING_1,
      alignment: AlignmentType.CENTER,
      spacing: { before: 600, after: 200 },
      pageBreakBefore: true,
    }),
    new Paragraph({
      text: "STATEMENT OF WORK",
      heading: HeadingLevel.HEADING_2,
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
    }),
    new Paragraph({
      children: [
        new TextRun({ text: "Customer:", bold: true, underline: { type: UnderlineType.SINGLE } }),
        new TextRun({ text: ` ${clientName}` }),
      ],
      spacing: { after: 100 },
    }),
    new Paragraph({
      children: [
        new TextRun({ text: "Service Provider:", bold: true, underline: { type: UnderlineType.SINGLE } }),
        new TextRun({ text: " WeGo! Oakland" }),
      ],
      spacing: { after: 100 },
    }),
    new Paragraph({
      children: [
        new TextRun({ text: "Service Provider Business Hours:", bold: true, underline: { type: UnderlineType.SINGLE } }),
        new TextRun({ text: ' Business hours are Monday through Friday, 10 a.m.-5 p.m. PST or PDT, as applicable ("Business Hours")' }),
      ],
      spacing: { after: 400 },
    }),
  ];

  // Payment Schedule
  if (paymentSchedule && paymentSchedule.length > 0) {
    paragraphs.push(
      new Paragraph({
        children: [
          new TextRun({ text: "Fee Schedule and Payment Details", bold: true, underline: { type: UnderlineType.SINGLE } }),
        ],
        spacing: { before: 400, after: 200 },
      })
    );

    // Add general payment text
    paragraphs.push(
      new Paragraph({
        text: "A non-refundable deposit of 40% of the total project fee is required to reserve the design period. The project timeline will not be scheduled until this payment is received. An invoice will be sent upon contract signing, and the deposit is due immediately.",
        spacing: { after: 200 },
      })
    );

    // Payment schedule from AI
    paragraphs.push(
      new Paragraph({
        text: "Customer agrees to pay Service Provider on the following payment schedule:",
        spacing: { after: 200 },
      })
    );

    paymentSchedule.forEach((payment) => {
      const formattedAmount = `$${parseFloat(payment.amount).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
      paragraphs.push(
        new Paragraph({
          children: [
            new TextRun({ text: formattedAmount, bold: true }),
            new TextRun({ text: `, ${payment.description}, due ` }),
            new TextRun({ text: payment.dueDate, bold: true }),
          ],
          spacing: { after: 100 },
          bullet: { level: 0 },
        })
      );
    });

    paragraphs.push(
      new Paragraph({
        text: "Payment can be made via Zelle or Venmo or other electronic payment processor as agreed to by both parties. Service Provider will invoice the Customer on the payment schedule listed above.",
        spacing: { before: 200, after: 200 },
      }),
      new Paragraph({
        text: "Venmo: @salaamssalaams",
        spacing: { after: 100 },
      }),
      new Paragraph({
        text: "Zelle: salaams@wegooakland.com",
        spacing: { after: 200 },
      }),
      new Paragraph({
        children: [
          new TextRun({ text: "Customer Initials: ", italics: true }),
          new TextRun({ text: initialsField, bold: true, color: "0000FF" }),
        ],
        spacing: { before: 300, after: 400 },
      })
    );
  }

  // Services to be Provided
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: "Services to be Provided", bold: true, underline: { type: UnderlineType.SINGLE } }),
      ],
      spacing: { before: 400, after: 200 },
    })
  );

  // If structured services data available, use it
  if (services && services.length > 0) {
    services.forEach((service, index) => {
      // Service title (numbered)
      paragraphs.push(
        new Paragraph({
          children: [
            new TextRun({ text: `${index + 1}. `, bold: true }),
            new TextRun({ text: service.title, bold: true }),
          ],
          spacing: { before: 300, after: 200 },
        })
      );

      // Includes section
      if (service.includes && service.includes.length > 0) {
        paragraphs.push(
          new Paragraph({
            children: [
              new TextRun({ text: "Includes:", italics: true }),
            ],
            spacing: { before: 100, after: 100 },
          })
        );

        service.includes.forEach((item) => {
          paragraphs.push(
            new Paragraph({
              text: item,
              spacing: { after: 100 },
              bullet: { level: 0 },
            })
          );
        });
      }

      // What's Not Included section
      if (service.notIncluded && service.notIncluded.length > 0) {
        paragraphs.push(
          new Paragraph({
            children: [
              new TextRun({ text: "What's Not Included:", bold: true, underline: { type: UnderlineType.SINGLE } }),
            ],
            spacing: { before: 200, after: 100 },
          })
        );

        service.notIncluded.forEach((item) => {
          paragraphs.push(
            new Paragraph({
              text: item,
              spacing: { after: 100 },
              bullet: { level: 0 },
            })
          );
        });
      }
    });
  } else if (servicesDescription) {
    // Fallback to plain text if no structured data
    paragraphs.push(
      new Paragraph({
        text: servicesDescription,
        spacing: { after: 300 },
      })
    );
  }

  // Project Timeline section
  if (projectPhases && projectPhases.length > 0) {
    paragraphs.push(
      new Paragraph({
        children: [
          new TextRun({ text: "Project Timeline", bold: true, underline: { type: UnderlineType.SINGLE } }),
        ],
        spacing: { before: 400, after: 200 },
      })
    );

    projectPhases.forEach((phase, index) => {
      paragraphs.push(
        new Paragraph({
          children: [
            new TextRun({ text: `${index + 1}. ${phase.phase}`, bold: true }),
            phase.duration ? new TextRun({ text: ` (${phase.duration})` }) : new TextRun({ text: "" }),
          ],
          spacing: { before: 200, after: 100 },
        }),
        new Paragraph({
          text: phase.description,
          spacing: { after: 200 },
          indent: { left: convertInchesToTwip(0.5) },
        })
      );
    });
  }

  // Revision Policy
  if (revisionPolicy) {
    paragraphs.push(
      new Paragraph({
        children: [
          new TextRun({ text: "Revision Policy", bold: true, underline: { type: UnderlineType.SINGLE } }),
        ],
        spacing: { before: 400, after: 200 },
      }),
      new Paragraph({
        text: revisionPolicy,
        spacing: { after: 300 },
      })
    );
  }

  // Important Notes
  if (importantNotes && importantNotes.length > 0) {
    importantNotes.forEach((note) => {
      paragraphs.push(
        new Paragraph({
          children: [
            new TextRun({ text: "IMPORTANT: ", bold: true, underline: { type: UnderlineType.SINGLE } }),
            new TextRun({ text: note }),
          ],
          spacing: { before: 300, after: 200 },
        })
      );
    });
  }

  // Final Customer Initials
  paragraphs.push(
    new Paragraph({
      children: [
        new TextRun({ text: "Customer Initials: ", italics: true }),
        new TextRun({ text: initialsField, bold: true, color: "0000FF" }),
      ],
      spacing: { before: 400, after: 400 },
    })
  );

  return paragraphs;
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

  const members = params.clientManagingMembers ? params.clientManagingMembers.split(',').map(m => m.trim()) : [];

  const submitters = [
    {
      role: 'Provider',
      email: process.env.PROVIDER_EMAIL || 'salaams@wegooakland.com',
      name: 'Salaams J. deRosa',
    },
    {
      role: 'Customer1',
      email: params.signerEmail,
      name: members[0] || 'Customer',
    },
  ];

  // Add second customer if multiple managing members
  if (members.length > 1 && params.signerEmail2) {
    submitters.push({
      role: 'Customer2',
      email: params.signerEmail2,
      name: members[1],
    });
  }

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
    submitters: submitters,
  });

  return submission;
}

// Generate contract endpoint
app.post('/api/generate-contract', async (req, res) => {
  try {
    console.log('Generating contract with data:', req.body);

    // Generate contract
    const doc = generateFullContract(req.body);

    // Save to file
    const outputDir = path.join(__dirname, 'output');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').split('T')[0];
    const fileName = `service-agreement-${req.body.clientName.replace(/\s+/g, '-').toLowerCase()}-${timestamp}.docx`;
    const filePath = path.join(outputDir, fileName);

    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(filePath, buffer);

    console.log('Contract saved:', fileName);

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
  console.log(`\n🚀 WE GO OAKLAND Contract Generator (AI-Powered)`);
  console.log(`📱 Open your browser to: http://localhost:${PORT}`);
  console.log(`\n✨ Ready to generate contracts!\n`);
});
