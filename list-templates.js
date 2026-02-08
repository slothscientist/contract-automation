#!/usr/bin/env node
require('dotenv').config();
const { DocusealApi } = require('@docuseal/api');

async function listTemplates() {
  const docuseal = new DocusealApi({ key: process.env.DOCUSEAL_API_KEY });

  console.log('Fetching your DocuSeal templates...\n');

  try {
    const response = await docuseal.listTemplates();

    // Handle different response formats
    const templates = response.data || response;

    console.log('Raw response type:', typeof response);
    console.log('Raw response:', JSON.stringify(response, null, 2));

  } catch (error) {
    console.error('Error:', error.message);
    console.error('Full error:', error);
  }
}

listTemplates();
