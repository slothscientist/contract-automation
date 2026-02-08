#!/usr/bin/env node
require('dotenv').config();
const { DocusealApi } = require('@docuseal/api');

async function listTemplates() {
  const docuseal = new DocusealApi({ key: process.env.DOCUSEAL_API_KEY });

  console.log('Fetching your DocuSeal templates...\n');

  try {
    const templates = await docuseal.listTemplates();

    if (!templates || templates.length === 0) {
      console.log('No templates found.');
      return;
    }

    templates.forEach(t => {
      console.log('─────────────────────────────────');
      console.log(`Name: ${t.name}`);
      console.log(`ID:   ${t.id}`);
      console.log(`Fields:`);
      if (t.fields && t.fields.length > 0) {
        t.fields.forEach(f => {
          console.log(`  - ${f.name} (${f.type})`);
        });
      }
    });

    console.log('─────────────────────────────────');
    console.log('\nCopy the ID of "Contract Template" - you need it!');
  } catch (error) {
    console.error('Error:', error.message);
  }
}

listTemplates();
