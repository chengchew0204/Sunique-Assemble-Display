# Assembly Schedule Display System

A two-component system for displaying today's assembly schedule from SharePoint:
- **Node.js Proxy Server**: Handles SharePoint authentication and file downloads
- **HTML Client**: Displays the schedule in a clean, minimalist interface

## Architecture

```
HTML (Browser) → Node.js Server (Railway) → SharePoint/Microsoft Graph API
```

The proxy server is necessary to avoid CORS issues and keep credentials secure.

## Deployment Instructions

### 1. Deploy Server to Railway

1. Create a Railway account at https://railway.app
2. Click "New Project" → "Deploy from GitHub repo" (or use CLI)
3. Select this repository
4. Railway will auto-detect Node.js and use `npm start`

### 2. Configure Environment Variables on Railway

In your Railway project dashboard, add these environment variables:

```
SHAREPOINT_TENANT_ID=your-tenant-id
SHAREPOINT_CLIENT_ID=your-client-id
SHAREPOINT_CLIENT_SECRET=your-client-secret
SHAREPOINT_HOSTNAME=your-hostname.sharepoint.com
```

Use the values from your `env` file (do not commit this file to git!).

Note: SHAREPOINT_SITE_NAME is no longer needed as the server automatically searches in the correct site.

### 3. Get Your Railway URL

After deployment, Railway will provide a URL like:
```
https://your-app-name.railway.app
```

### 4. Update HTML File

Open `assembly-schedule.html` and update line 116:

```javascript
const API_SERVER_URL = 'https://your-app-name.railway.app';
```

### 5. Open HTML File

Simply open `assembly-schedule.html` in your browser. It will automatically:
- Fetch today's assembly schedule
- Display orders scheduled for today
- Show clean, minimal interface

## Local Development

To test locally before deploying:

1. Install dependencies:
```bash
npm install
```

2. Rename your `env` file to `.env` (or create a `.env` file with your credentials from step 2 above)

3. Start the server:
```bash
npm start
```

4. Open `assembly-schedule.html` in your browser (API_SERVER_URL should be `http://localhost:3000`)

## API Endpoints

- `GET /` - Health check and API information
- `GET /api/download-schedule` - Download the assembly schedule Excel file

## Security Notes

- Credentials are stored as environment variables on Railway (secure)
- The HTML file contains no credentials (safe to share)
- CORS is enabled to allow browser access
- Never commit the `.env` file to version control

## Files

- `server.js` - Express server with SharePoint proxy
- `package.json` - Node.js dependencies
- `assembly-schedule.html` - Client interface
- `.env` - Local environment variables (not in git)
- `env` - Template with credentials (rename to `.env` for local use)

## Troubleshooting

**Error: "NetworkError when attempting to fetch resource"**
- Make sure the Railway server is running
- Check that API_SERVER_URL in HTML matches your Railway URL
- Verify environment variables are set correctly on Railway

**Error: "Authentication failed"**
- Verify SharePoint credentials in Railway environment variables
- Check that the client secret hasn't expired

**"No orders scheduled for today"**
- This is normal if there are no orders for today's date
- The system filters by exact date match

