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
    hostname: process.env.SHAREPOINT_HOSTNAME
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

// Get the site that contains the file (optimized - goes directly to the working site)
async function getFileSite(accessToken) {
    // We know the file is in SuniqueKnowledgeBase site, go there directly
    const siteUrl = `https://graph.microsoft.com/v1.0/sites/${CONFIG.hostname}:/sites/SuniqueKnowledgeBase`;
    
    const response = await fetch(siteUrl, {
        headers: {
            'Authorization': `Bearer ${accessToken}`
        }
    });
    
    if (!response.ok) {
        throw new Error(`Failed to access SuniqueKnowledgeBase site: ${response.status}`);
    }

    const data = await response.json();
    return { id: data.id, name: data.name, displayName: data.displayName };
}

// Find file in the site (optimized - search directly in default drive)
async function findFile(accessToken, site) {
    const fileName = 'Assembly Schedule (New Version).xlsx';
    const searchUrl = `https://graph.microsoft.com/v1.0/sites/${site.id}/drive/root/search(q='${encodeURIComponent(fileName)}')`;
    
    const response = await fetch(searchUrl, {
        headers: {
            'Authorization': `Bearer ${accessToken}`
        }
    });

    if (response.ok) {
        const data = await response.json();
        if (data.value && data.value.length > 0) {
            const file = data.value[0];
            return { driveId: file.parentReference?.driveId, itemId: file.id };
        }
    }

    throw new Error(`File "${fileName}" not found in ${site.displayName || site.name}`);
}

// Download Excel file using drive ID and item ID
async function downloadFile(accessToken, driveId, itemId) {
    const fileUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`;
    
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

// Main endpoint to download the assembly schedule (optimized)
app.get('/api/download-schedule', async (req, res) => {
    try {
        // Step 1: Authenticate
        const accessToken = await getAccessToken();
        
        // Step 2: Get site
        const site = await getFileSite(accessToken);
        
        // Step 3: Find file
        const fileLocation = await findFile(accessToken, site);
        
        // Step 4: Download file
        const fileBuffer = await downloadFile(accessToken, fileLocation.driveId, fileLocation.itemId);
        
        // Send the file as binary data
        res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(fileBuffer);
        
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
