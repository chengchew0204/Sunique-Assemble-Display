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
    siteName: process.env.SHAREPOINT_SITE_NAME || 'SuniqueKnowledgeBase'
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

// Get site ID - try multiple approaches
async function getSiteId(accessToken) {
    // Try with /sites/ prefix first
    let siteUrl = `https://graph.microsoft.com/v1.0/sites/${CONFIG.hostname}:/sites/${CONFIG.siteName}`;
    
    let response = await fetch(siteUrl, {
        headers: {
            'Authorization': `Bearer ${accessToken}`
        }
    });

    // If that fails, try without /sites/ prefix
    if (!response.ok) {
        console.log('Trying root site...');
        siteUrl = `https://graph.microsoft.com/v1.0/sites/${CONFIG.hostname}`;
        response = await fetch(siteUrl, {
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });
    }

    if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to get site: ${response.status} - ${errorText}`);
    }

    const data = await response.json();
    console.log('Site ID:', data.id, 'Site Name:', data.name);
    return data.id;
}

// Find file by searching for its name - returns {driveId, itemId}
async function findFile(accessToken, siteId) {
    const fileName = 'Assembly Schedule (New Version).xlsx';
    
    // Try approach 1: Search in site drive
    console.log('Searching in site drive...');
    let searchUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root/search(q='${encodeURIComponent(fileName)}')`;
    
    let response = await fetch(searchUrl, {
        headers: {
            'Authorization': `Bearer ${accessToken}`
        }
    });

    if (response.ok) {
        const data = await response.json();
        if (data.value && data.value.length > 0) {
            const file = data.value[0];
            console.log('File found in site drive:', file.id, 'Drive:', file.parentReference?.driveId);
            return { driveId: file.parentReference?.driveId, itemId: file.id };
        }
    }

    // Try approach 2: List all drives and search each
    console.log('Listing all drives in site...');
    const drivesUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`;
    response = await fetch(drivesUrl, {
        headers: {
            'Authorization': `Bearer ${accessToken}`
        }
    });

    if (response.ok) {
        const drivesData = await response.json();
        console.log(`Found ${drivesData.value?.length || 0} drives`);
        
        for (const drive of drivesData.value || []) {
            console.log(`Searching in drive: ${drive.name} (${drive.id})`);
            searchUrl = `https://graph.microsoft.com/v1.0/drives/${drive.id}/root/search(q='${encodeURIComponent(fileName)}')`;
            
            response = await fetch(searchUrl, {
                headers: {
                    'Authorization': `Bearer ${accessToken}`
                }
            });

            if (response.ok) {
                const searchData = await response.json();
                if (searchData.value && searchData.value.length > 0) {
                    const file = searchData.value[0];
                    console.log('File found in drive:', drive.name, '- File ID:', file.id);
                    return { driveId: drive.id, itemId: file.id };
                }
            }
        }
    }

    throw new Error(`File "${fileName}" not found in any SharePoint drives`);
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
        
        // Step 3: Find file
        console.log('Searching for file...');
        const fileLocation = await findFile(accessToken, siteId);
        console.log('File found - Drive:', fileLocation.driveId, 'Item:', fileLocation.itemId);
        
        // Step 4: Download file
        console.log('Downloading file...');
        const fileBuffer = await downloadFile(accessToken, fileLocation.driveId, fileLocation.itemId);
        
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
