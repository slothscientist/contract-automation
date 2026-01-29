#!/usr/bin/env node

require('dotenv').config();
const { DocusealApi } = require('@docuseal/api');

async function testApiConnection() {
  console.log('🔍 Testing DocuSeal API Connection...\n');

  const apiKey = process.env.DOCUSEAL_API_KEY;

  if (!apiKey) {
    console.error('❌ DOCUSEAL_API_KEY not found in .env file');
    process.exit(1);
  }

  console.log('✅ API Key found:', apiKey.substring(0, 10) + '...');

  try {
    const docuseal = new DocusealApi({ key: apiKey });

    console.log('\n📋 Fetching existing templates...');
    const templates = await docuseal.listTemplates();

    console.log('✅ API Connection Successful!');
    console.log(`📄 Found ${templates.data?.length || 0} templates in your account\n`);

    if (templates.data && templates.data.length > 0) {
      console.log('Templates:');
      templates.data.slice(0, 5).forEach((template) => {
        console.log(`  - ${template.name} (ID: ${template.id})`);
      });
    }

    console.log('\n📋 Fetching existing submissions...');
    const submissions = await docuseal.listSubmissions();

    console.log(`✅ Found ${submissions.data?.length || 0} submissions in your account\n`);

    if (submissions.data && submissions.data.length > 0) {
      console.log('Recent submissions:');
      submissions.data.slice(0, 5).forEach((submission) => {
        console.log(`  - ${submission.template?.name || 'Unnamed'} (ID: ${submission.id}) - ${submission.status}`);
      });
    }

    console.log('\n✅ API connection test passed!');
    console.log('Your API key is valid and working.\n');

  } catch (error) {
    console.error('\n❌ API Connection Failed!');
    console.error('Error:', error.message);

    if (error.response) {
      console.error('Response status:', error.response.status);
      console.error('Response data:', JSON.stringify(error.response.data, null, 2));
    }

    if (error.message.includes('401') || error.message.includes('Unauthorized')) {
      console.error('\n💡 Your API key appears to be invalid or expired.');
      console.error('Please check your DocuSeal account and generate a new key.');
    }

    process.exit(1);
  }
}

testApiConnection();
