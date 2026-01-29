#!/bin/bash

echo "🧪 Testing Contract Automation System"
echo "======================================"
echo ""

# Check if .env exists
if [ ! -f .env ]; then
    echo "❌ .env file not found!"
    echo "Creating .env file..."
    cat > .env << EOF
# DocuSeal API Configuration
DOCUSEAL_API_KEY=a5i7mTYMwbk2CkLgyungT2p1H7ndguuTLXGMUvVYNyR

# Provider Information
PROVIDER_EMAIL=salaams@wegooakland.com
EOF
    echo "✅ .env file created"
fi

echo "📋 Running test contract generation..."
echo ""

node generate-contract.js \
  --clientName "Test Academy" \
  --signerName "Test Test-Test" \
  --signerEmail "salaams+test@wegooakland.com" \
  --projectStartDate "2026-02-01" \
  --projectEndDate "2026-03-15" \
  --totalFee "6000"

if [ $? -eq 0 ]; then
    echo ""
    echo "✅ Test completed successfully!"
    echo ""
    echo "📧 Check your email inbox at:"
    echo "   - salaams@wegooakland.com (provider)"
    echo "   - salaams+test@wegooakland.com (client)"
    echo ""
    echo "📄 Generated contract saved to: output/"
else
    echo ""
    echo "⚠️  Test completed with errors (see above)"
    echo ""
    echo "If you see a network error, make sure you have internet access."
    echo "The contract should still be generated in the output/ folder."
fi
