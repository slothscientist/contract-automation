# Troubleshooting Guide

## Nothing Showing Up in DocuSeal Account

If you ran the script but don't see any submissions in your DocuSeal account, follow these steps:

### Step 1: Test API Connection

Run the API connection test:

```bash
node test-api.js
```

This will verify:
- Your API key is valid
- You can connect to DocuSeal
- List existing templates and submissions

**Expected output:**
```
✅ API Connection Successful!
📄 Found X templates in your account
✅ Found X submissions in your account
```

### Step 2: Check for Errors

If you got errors, check:

#### Authentication Errors (401/403)
```
💡 Authentication Error - Your API key may be invalid or expired.
```

**Solution:**
1. Go to https://docuseal.com/settings/api
2. Generate a new API key
3. Update your `.env` file:
   ```
   DOCUSEAL_API_KEY=your_new_key_here
   ```

#### Network Errors
```
❌ DocuSeal API Error: getaddrinfo EAI_AGAIN api.docuseal.com
```

**Solution:**
- Check your internet connection
- Try again in a few minutes
- Check if DocuSeal.com is accessible in your region

#### File Size Errors
```
❌ File too large
```

**Solution:**
- The generated contract might be too large
- Check the file size in the output

### Step 3: Run with Verbose Debugging

Run the contract generation script and carefully review the output:

```bash
node generate-contract.js \
  --clientName "Debug Test" \
  --signerName "Your Name" \
  --signerEmail "your-email@example.com" \
  --projectStartDate "2026-02-01" \
  --projectEndDate "2026-03-01" \
  --totalFee "1000"
```

Look for these sections in the output:

1. **API Key Check:**
   ```
   🔑 API Key found: a5i7mTYMwb...
   ```

2. **File Upload:**
   ```
   File size: 8576 bytes (11.50 KB base64)
   ⏳ Sending to DocuSeal API...
   ```

3. **Success Response:**
   ```
   ✅ Document sent for signing!
   📧 Submission ID: 12345
   ```

4. **Full API Response:**
   Look for the complete JSON response that shows all details

### Step 4: Check Common Issues

#### Issue: Script Exits Without Error

If the script completes but nothing appears:

1. **Check your DocuSeal account dashboard:**
   - Log in to https://docuseal.com
   - Go to "Submissions" tab
   - Look for submissions dated today

2. **Check email spam folder:**
   - Signing requests might be in spam
   - Add @docuseal.com to your safe senders

3. **Verify API key permissions:**
   - Some API keys have limited permissions
   - Make sure your key can create submissions

#### Issue: Wrong DocuSeal Account

Make sure you're:
- Using the correct API key for your account
- Logged into the right DocuSeal account
- Not using a test/sandbox environment key on production (or vice versa)

### Step 5: Manual API Test

If issues persist, test the API manually:

```bash
# Save this as test-manual.js
node -e "
const { DocusealApi } = require('@docuseal/api');
const api = new DocusealApi({ key: 'YOUR_API_KEY_HERE' });
api.listSubmissions().then(r => console.log('Success:', r)).catch(e => console.error('Error:', e));
"
```

### Step 6: Contact Support

If none of the above works:

1. **Check DocuSeal Status:**
   - Visit https://status.docuseal.com
   - See if there are any ongoing issues

2. **Review API Documentation:**
   - https://www.docuseal.com/docs/api
   - Check for any recent API changes

3. **Contact DocuSeal Support:**
   - Email: support@docuseal.com
   - Include:
     - Your account email
     - Error messages from the script
     - Output from `node test-api.js`

## Other Common Issues

### "Cannot find module 'docx'"

**Solution:**
```bash
npm install
```

### "Permission denied" when running script

**Solution:**
```bash
chmod +x generate-contract.js
chmod +x test-api.js
```

### Generated DOCX file is corrupt

**Solution:**
- Check the file size: `ls -lh output/`
- Try opening with different software (MS Word, LibreOffice, Google Docs)
- Regenerate the file

### Email not being sent

If DocuSeal creates the submission but emails aren't sent:

1. Check the `send_email` parameter is `true` (it is by default)
2. Verify email addresses are correct
3. Check spam folders
4. Check DocuSeal email settings in your account

## Getting Help

If you're still stuck:

1. Run diagnostics:
   ```bash
   node test-api.js > debug.log 2>&1
   ```

2. Run the contract script with output:
   ```bash
   node generate-contract.js [your-params] > contract-debug.log 2>&1
   ```

3. Review the log files and share relevant errors

Contact: salaams@wegooakland.com
