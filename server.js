const express = require('express');
const cors = require('cors');
const fetch = require('node-fetch');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// Enable CORS for all origins (you can restrict this to specific domains if needed)
app.use(cors());
app.use(express.json());

// Configuration from environment variables
const CONFIG = {
    tenantId: process.env.SHAREPOINT_TENANT_ID,
    clientId: process.env.SHAREPOINT_CLIENT_ID,
    clientSecret: process.env.SHAREPOINT_CLIENT_SECRET,
    hostname: process.env.SHAREPOINT_HOSTNAME,
    siteName: process.env.SHAREPOINT_SITE_NAME || 'SuniqueKnowledgeBase',
    fileId: '90B92EAC-A9BD-48EC-9881-F6DC23DD5B4F'
};

// Validate configuration
function validateConfig() {
    const missing = [];
    if (!CONFIG.tenantId) missing.push('SHAREPOINT_TENANT_ID');
    if (!CONFIG.clientId) missing.push('SHAREPOINT_CLIENT_ID');
    if (!CONFIG.clientSecret) missing.push('SHAREPOINT_CLIENT_SECRET');
    if (!CONFIG.hostname) missing.push('SHAREPOINT_HOSTNAME');
    
    if (missing.length > 0) {
        throw new Error(`Missing required environment variables: ${missing.join(', ')}`);
    }
}

// Get access token from Microsoft
async function getAccessToken() {
    const tokenEndpoint = `https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/token`;
    
    const params = new URLSearchParams();
    params.append('client_id', CONFIG.clientId);
    params.append('client_secret', CONFIG.clientSecret);
    params.append('scope', 'https://graph.microsoft.com/.default');
    params.append('grant_type', 'client_credentials');

    const response = await fetch(tokenEndpoint, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: params
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Authentication failed: ${response.status} - ${errorText}`);
    }

    const data = await response.json();
    return data.access_token;
}

// Get site ID
async function getSiteId(accessToken) {
    const siteUrl = `https://graph.microsoft.com/v1.0/sites/${CONFIG.hostname}:/sites/${CONFIG.siteName}`;
    
    const response = await fetch(siteUrl, {
        headers: {
            'Authorization': `Bearer ${accessToken}`
        }
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to get site: ${response.status} - ${errorText}`);
    }

    const data = await response.json();
    return data.id;
}

// Download Excel file
async function downloadFile(accessToken, siteId) {
    const fileUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${CONFIG.fileId}/content`;
    
    const response = await fetch(fileUrl, {
        headers: {
            'Authorization': `Bearer ${accessToken}`
        }
    });

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to download file: ${response.status} - ${errorText}`);
    }

    return await response.buffer();
}

// Health check endpoint
app.get('/', (req, res) => {
    res.json({ 
        status: 'ok', 
        message: 'Assembly Schedule API Server',
        endpoints: {
            health: 'GET /',
            downloadFile: 'GET /api/download-schedule'
        }
    });
});

// Main endpoint to download the assembly schedule
app.get('/api/download-schedule', async (req, res) => {
    try {
        console.log('Fetching assembly schedule...');
        
        // Step 1: Authenticate
        console.log('Authenticating with Microsoft...');
        const accessToken = await getAccessToken();
        
        // Step 2: Get site ID
        console.log('Getting site ID...');
        const siteId = await getSiteId(accessToken);
        
        // Step 3: Download file
        console.log('Downloading file...');
        const fileBuffer = await downloadFile(accessToken, siteId);
        
        // Send the file as binary data
        res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(fileBuffer);
        
        console.log('File sent successfully');
        
    } catch (error) {
        console.error('Error:', error.message);
        res.status(500).json({ 
            error: error.message,
            stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
        });
    }
});

// Start server
try {
    validateConfig();
    app.listen(PORT, () => {
        console.log(`Server running on port ${PORT}`);
        console.log(`API available at http://localhost:${PORT}/api/download-schedule`);
    });
} catch (error) {
    console.error('Failed to start server:', error.message);
    process.exit(1);
}
