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
    fileGuid: '90B92EAC-A9BD-48EC-9881-F6DC23DD5B4F'
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

// Get all relevant site IDs to search
async function getAllSiteIds(accessToken) {
    const siteIds = [];
    
    // Try site from config (sccr)
    const sitesToTry = [
        { name: CONFIG.siteName, url: `https://graph.microsoft.com/v1.0/sites/${CONFIG.hostname}:/sites/${CONFIG.siteName}` },
        { name: 'SuniqueKnowledgeBase', url: `https://graph.microsoft.com/v1.0/sites/${CONFIG.hostname}:/sites/SuniqueKnowledgeBase` },
        { name: 'root', url: `https://graph.microsoft.com/v1.0/sites/${CONFIG.hostname}` }
    ];
    
    for (const site of sitesToTry) {
        console.log(`Trying site: ${site.name}...`);
        const response = await fetch(site.url, {
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });
        
        if (response.ok) {
            const data = await response.json();
            console.log(`✓ Found site: ${data.name} (${data.id})`);
            siteIds.push({ id: data.id, name: data.name, displayName: data.displayName });
        } else {
            console.log(`✗ Site ${site.name} not accessible: ${response.status}`);
        }
    }
    
    if (siteIds.length === 0) {
        throw new Error('No accessible SharePoint sites found');
    }
    
    return siteIds;
}

// Try to get file by GUID using SharePoint REST API
async function getFileByGuid() {
    // Try both site names
    const sitesToTry = ['SuniqueKnowledgeBase', CONFIG.siteName, 'sccr'];
    
    // Get SharePoint-specific token
    const tokenEndpoint = `https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/token`;
    const params = new URLSearchParams();
    params.append('client_id', CONFIG.clientId);
    params.append('client_secret', CONFIG.clientSecret);
    params.append('scope', `https://${CONFIG.hostname}/.default`);
    params.append('grant_type', 'client_credentials');

    const tokenResponse = await fetch(tokenEndpoint, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: params
    });

    if (!tokenResponse.ok) {
        console.log('Failed to get SharePoint token');
        return null;
    }

    const tokenData = await tokenResponse.json();
    const spToken = tokenData.access_token;

    for (const siteName of sitesToTry) {
        const webUrl = `https://${CONFIG.hostname}/sites/${siteName}/_api/web/GetFileById(guid'${CONFIG.fileGuid}')/$value`;
        console.log(`Trying SharePoint REST API in site: ${siteName}...`);
        
        const response = await fetch(webUrl, {
            headers: {
                'Authorization': `Bearer ${spToken}`,
                'Accept': 'application/json;odata=verbose'
            }
        });

        if (response.ok) {
            console.log(`✓ File found via SharePoint REST API in site: ${siteName}`);
            return await response.buffer();
        } else {
            console.log(`  ✗ Not found in ${siteName} (${response.status})`);
        }
    }

    return null;
}

// Find file by searching across all accessible sites
async function findFileInAllSites(accessToken, sites) {
    const fileName = 'Assembly Schedule (New Version).xlsx';
    
    for (const site of sites) {
        console.log(`\n=== Searching in site: ${site.displayName || site.name} ===`);
        
        // Try approach 1: Search in site drive
        console.log('Searching in default site drive...');
        let searchUrl = `https://graph.microsoft.com/v1.0/sites/${site.id}/drive/root/search(q='${encodeURIComponent(fileName)}')`;
        
        let response = await fetch(searchUrl, {
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });

        if (response.ok) {
            const data = await response.json();
            console.log('Search results:', data.value?.length || 0, 'files found');
            if (data.value && data.value.length > 0) {
                const file = data.value[0];
                console.log('✓ FILE FOUND in site drive!');
                console.log('File name:', file.name);
                console.log('File ID:', file.id);
                return { driveId: file.parentReference?.driveId, itemId: file.id };
            }
        } else {
            console.log('Site drive search failed:', response.status);
        }

        // Try approach 2: List all drives and search each
        console.log('Listing all drives in site...');
        const drivesUrl = `https://graph.microsoft.com/v1.0/sites/${site.id}/drives`;
        response = await fetch(drivesUrl, {
            headers: {
                'Authorization': `Bearer ${accessToken}`
            }
        });

        if (response.ok) {
            const drivesData = await response.json();
            console.log(`Found ${drivesData.value?.length || 0} drives`);
            
            for (const drive of drivesData.value || []) {
                console.log(`  Searching in drive: ${drive.name}...`);
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
                        console.log(`  ✓ FILE FOUND in drive: ${drive.name}`);
                        console.log('  File name:', file.name);
                        console.log('  File ID:', file.id);
                        return { driveId: drive.id, itemId: file.id };
                    } else {
                        console.log(`  Drive ${drive.name}: 0 results`);
                    }
                }
            }
        } else {
            console.log('Failed to list drives:', response.status);
        }
    }

    throw new Error(`File "${fileName}" not found in any accessible SharePoint sites`);
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
        console.log('\n========================================');
        console.log('Fetching assembly schedule...');
        console.log('========================================');
        
        // Step 1: Authenticate
        console.log('\n[1/4] Authenticating with Microsoft...');
        const accessToken = await getAccessToken();
        console.log('✓ Authentication successful');
        
        // Step 2: Get all accessible site IDs
        console.log('\n[2/4] Finding SharePoint sites...');
        const sites = await getAllSiteIds(accessToken);
        console.log(`✓ Found ${sites.length} accessible site(s)`);
        
        // Step 3: Try to get file by GUID first (SharePoint REST API)
        console.log('\n[3/4] Trying to get file by GUID...');
        let fileBuffer = await getFileByGuid();
        
        if (!fileBuffer) {
            // Step 4: Fallback to searching across all sites
            console.log('GUID approach failed, searching across all sites...');
            const fileLocation = await findFileInAllSites(accessToken, sites);
            console.log('\n✓ File located!');
            console.log('  Drive ID:', fileLocation.driveId);
            console.log('  Item ID:', fileLocation.itemId);
            
            // Step 5: Download file
            console.log('\n[4/4] Downloading file...');
            fileBuffer = await downloadFile(accessToken, fileLocation.driveId, fileLocation.itemId);
        }
        
        // Send the file as binary data
        res.set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(fileBuffer);
        
        console.log('\n✓ File sent successfully!');
        console.log('========================================\n');
        
    } catch (error) {
        console.error('\n✗ Error:', error.message);
        console.error('Stack:', error.stack);
        console.error('========================================\n');
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
