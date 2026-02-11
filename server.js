const https = require('https');
const fs = require('fs');
const path = require('path');
const os = require('os');
const express = require('express');

const app = express();
const PORT = 3000;

// Serve static files
app.use(express.static(__dirname));

// HTTPS options with self-signed certificate from office-addin-dev-certs
const certPath = path.join(os.homedir(), '.office-addin-dev-certs');
const httpsOptions = {
    key: fs.readFileSync(path.join(certPath, 'localhost.key')),
    cert: fs.readFileSync(path.join(certPath, 'localhost.crt'))
};

https.createServer(httpsOptions, app).listen(PORT, () => {
    console.log(`HTTPS Server running at https://localhost:${PORT}`);
    console.log('Add-in is ready to be sideloaded into Word on the web');
});
