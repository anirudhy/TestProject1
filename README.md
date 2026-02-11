# Word Auto-open Add-in

Minimal Word Office Add-in demonstrating auto-open task pane functionality for Word on the web.

## Folder Structure

```
word-autoopen-addin/
├── manifest.xml          # Office Add-in manifest (classic XML)
├── package.json          # Node.js dependencies
├── server.js            # HTTPS server
├── taskpane.html        # Task pane UI
├── taskpane.js          # Task pane logic
├── commands.html        # Required commands file (empty)
└── assets/              # Icons
    ├── icon-16.png
    ├── icon-32.png
    ├── icon-64.png
    └── icon-80.png
```

## Setup Instructions

### 1. Install Dependencies

```bash
cd word-autoopen-addin
npm install
```

### 2. Generate Self-Signed Certificate

```bash
npm run cert
```

This creates `localhost-key.pem` and `localhost-cert.pem` in your project folder.

### 3. Start the HTTPS Server

```bash
npm start
```

The server will run at `https://localhost:3000`

### 4. Sideload into Word on the Web

1. Go to https://office.com and sign in
2. Create or open a Word document
3. Click **Insert** > **Add-ins** (or **Get Add-ins**)
4. Click **Upload My Add-in** (top-right corner)
5. Click **Browse...** and select your `manifest.xml` file
6. Click **Upload**

The add-in will appear in the ribbon under the **Home** tab.

## Testing Auto-Open Functionality

### Test Steps

1. **Open a document in Word on the web**
   - Go to https://office.com
   - Create a new blank document or open an existing one

2. **Insert the add-in manually (first time)**
   - Click **Insert** > **Add-ins** > **Upload My Add-in**
   - Upload `manifest.xml`
   - Click the "Show Taskpane" button in the Home ribbon

3. **Enable auto-open**
   - In the task pane, click **"Enable auto-open for this document"**
   - Status should change to: "Status: Auto-open is ENABLED"

4. **Close the tab**
   - Close the entire browser tab (or sign out of Word on the web)

5. **Reopen the same document**
   - Go back to https://office.com
   - Open the same document again

6. **Verify auto-open works**
   - ✅ The task pane should automatically open without clicking any buttons
   - ✅ Status should still show: "Status: Auto-open is ENABLED"

### Testing Disable

1. Click **"Disable auto-open for this document"**
2. Status should change to: "Status: Auto-open is DISABLED"
3. Close and reopen the document
4. Task pane should NOT open automatically

## How It Works

### Manifest Key Points

- **TaskpaneId**: Set to `Office.AutoShowTaskpaneWithDocument` (special ID that enables auto-open)
- **Host**: Set to "Document" for Word
- **Permissions**: ReadWriteDocument (required to save settings)

### Office.js Implementation

**Enable Auto-open:**
```javascript
Office.context.document.settings.set('Office.AutoShowTaskpaneWithDocument', true);
Office.context.document.settings.saveAsync();
```

**Disable Auto-open:**
```javascript
Office.context.document.settings.remove('Office.AutoShowTaskpaneWithDocument');
Office.context.document.settings.saveAsync();
```

**Check Status:**
```javascript
const isEnabled = Office.context.document.settings.get('Office.AutoShowTaskpaneWithDocument');
```

## Troubleshooting

### Add-in doesn't appear in ribbon
- Verify the server is running at `https://localhost:3000`
- Check that you uploaded the correct `manifest.xml` file
- Try refreshing the Word document

### Task pane doesn't auto-open
- Ensure you clicked "Enable auto-open" and saw the confirmation
- Verify you're reopening the **same document** (not a new one)
- Check that the setting was saved (status should say "ENABLED")
- Some browsers may block auto-open on first load - try refreshing

### Certificate errors
- If you see SSL warnings, click "Advanced" and "Proceed to localhost"
- Make sure you ran `npm run cert` before starting the server
- On Windows, you may need to install the certificate in Trusted Root Certification Authorities

### Status shows "Loading..." forever
- Check browser console for errors (F12)
- Verify Office.js is loading correctly
- Ensure the server is serving files from the correct directory

## Notes

- Auto-open only works for documents where the setting has been explicitly enabled
- The setting is stored per-document in the Office document's custom settings
- This requires Word on the web or Word 2016+ with proper add-in support
- Self-signed certificates are fine for development but use proper certificates for production
