const { app } = require('@azure/functions');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const { S3Client, GetObjectCommand, DeleteObjectCommand } = require('@aws-sdk/client-s3');

// ─── CONFIG ───────────────────────────────────────────────────────────────────
const TENANT_ID     = '6c5f12d0-7ab9-463d-b9df-7036176a0685';
const CLIENT_ID     = '5a7089e2-3370-4a02-8546-63e8d92d695c';
const CLIENT_SECRET = 'D_O8Q~BxfzushXOqEcCX2cGeoT1LFsOTNNeZHbCa';

const SHAREPOINT_SITE_URL = 'https://naffcogroup.sharepoint.com/sites/AI-APEX';
const SHAREPOINT_LIBRARY  = 'Shared Documents';

const R2_ACCOUNT_ID       = '13b2f661abaa4c2acce9dc45fe69b8cc';
const R2_ACCESS_KEY_ID    = '94f5152a6e045323d8cc18be33bbbdfc';
const R2_SECRET_ACCESS_KEY= '09e36f22255a22f5fac011be72d8dfb60a50ba2da3b6ece02b9f139a785af43d';
const R2_BUCKET           = 'estimation-docs';
const R2_PUBLIC_URL       = 'https://pub-a821f9113188484f80b16e7995fee042.r2.dev';

const JSONBIN_BIN_ID  = '69dcdFfeaaba882197f3c176';
const JSONBIN_API_KEY = '$2a$10$kpIFmWCwfUxqOw.M.TfqcOyhGnnArBzDluhGquW2s/t.L3vQJtBqW';
// ──────────────────────────────────────────────────────────────────────────────

// R2 Client
const r2 = new S3Client({
  region: 'auto',
  endpoint: `https://${R2_ACCOUNT_ID}.r2.cloudflarestorage.com`,
  credentials: {
    accessKeyId: R2_ACCESS_KEY_ID,
    secretAccessKey: R2_SECRET_ACCESS_KEY,
  },
});

// Graph Client
function getGraphClient() {
  const credential = new ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default'],
  });
  return Client.initWithMiddleware({ authProvider });
}

// Get SharePoint Site ID
async function getSiteId(graphClient) {
  const hostname = 'naffcogroup.sharepoint.com';
  const sitePath = '/sites/AI-APEX';
  const site = await graphClient.api(`/sites/${hostname}:${sitePath}`).get();
  return site.id;
}

// Get Drive ID for EstimationDocs library
async function getDriveId(graphClient, siteId) {
  const drives = await graphClient.api(`/sites/${siteId}/drives`).get();
  const drive = drives.value.find(d => d.name === SHAREPOINT_LIBRARY);
  if (!drive) throw new Error(`Library '${SHAREPOINT_LIBRARY}' not found`);
  return drive.id;
}

// Upload file to SharePoint
async function uploadToSharePoint(graphClient, siteId, driveId, fileName, fileBuffer, mimeType, folder) {
  const folderPath = folder ? `/${folder}` : '';
  const uploadPath = `/sites/${siteId}/drives/${driveId}/root:/Estimation_Docs${folderPath}/${fileName}:/content';
  
  const response = await graphClient
    .api(uploadPath)
    .header('Content-Type', mimeType || 'application/octet-stream')
    .put(fileBuffer);

  return response.webUrl;
}

// Download file from R2
async function downloadFromR2(fileKey) {
  const command = new GetObjectCommand({ Bucket: R2_BUCKET, Key: fileKey });
  const response = await r2.send(command);
  const chunks = [];
  for await (const chunk of response.Body) chunks.push(chunk);
  return { buffer: Buffer.concat(chunks), contentType: response.ContentType };
}

// Delete file from R2
async function deleteFromR2(fileKey) {
  await r2.send(new DeleteObjectCommand({ Bucket: R2_BUCKET, Key: fileKey }));
}

// Update JSONBin URL
async function updateJsonBinUrl(oldUrl, newUrl) {
  const res = await fetch(`https://api.jsonbin.io/v3/b/${JSONBIN_BIN_ID}/latest`, {
    headers: { 'X-Master-Key': JSONBIN_API_KEY }
  });
  const { record } = await res.json();
  const updated = JSON.stringify(record).replaceAll(oldUrl, newUrl);
  await fetch(`https://api.jsonbin.io/v3/b/${JSONBIN_BIN_ID}`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json', 'X-Master-Key': JSONBIN_API_KEY },
    body: updated
  });
}

// ─── FUNCTION 1: Health Check ─────────────────────────────────────────────────
app.http('health', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    return { 
      status: 200,
      jsonBody: { status: 'apex-bridge is running ✅', timestamp: new Date().toISOString() }
    };
  }
});

// ─── FUNCTION 2: Migrate R2 file → SharePoint ─────────────────────────────────
app.http('migrateToSharePoint', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const { fileUrl, fileName, mimeType, folder, requestId } = await request.json();

      if (!fileUrl || !fileName) {
        return { status: 400, jsonBody: { error: 'fileUrl and fileName are required' } };
      }

      context.log(`📦 Migrating: ${fileName} → SharePoint`);

      // Extract R2 key from URL
      const fileKey = fileUrl.replace(`${R2_PUBLIC_URL}/`, '');

      // 1. Download from R2
      const { buffer, contentType } = await downloadFromR2(fileKey);
      context.log(`✅ Downloaded from R2: ${fileName}`);

      // 2. Upload to SharePoint
      const graphClient = getGraphClient();
      const siteId = await getSiteId(graphClient);
      const driveId = await getDriveId(graphClient, siteId);
      const sharePointFolder = folder || requestId || 'General';
      const sharePointUrl = await uploadToSharePoint(
        graphClient, siteId, driveId, fileName, buffer, contentType || mimeType, sharePointFolder
      );
      context.log(`✅ Uploaded to SharePoint: ${sharePointUrl}`);

      // 3. Update JSONBin URL
      await updateJsonBinUrl(fileUrl, sharePointUrl);
      context.log(`✅ JSONBin URL updated`);

      // 4. Delete from R2
      await deleteFromR2(fileKey);
      context.log(`✅ Deleted from R2: ${fileName}`);

      return {
        status: 200,
        jsonBody: {
          success: true,
          fileName,
          sharePointUrl,
          message: 'File migrated to SharePoint successfully'
        }
      };

    } catch (err) {
      context.log(`❌ Error: ${err.message}`);
      return { status: 500, jsonBody: { error: err.message } };
    }
  }
});

// ─── FUNCTION 3: Upload directly to SharePoint ────────────────────────────────
app.http('uploadToSharePoint', {
  methods: ['POST'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    try {
      const formData = await request.formData();
      const file = formData.get('file');
      const folder = formData.get('folder') || 'General';

      if (!file) return { status: 400, jsonBody: { error: 'No file provided' } };

      const fileName = file.name;
      const buffer = Buffer.from(await file.arrayBuffer());

      context.log(`📤 Uploading: ${fileName} → SharePoint/${folder}`);

      const graphClient = getGraphClient();
      const siteId = await getSiteId(graphClient);
      const driveId = await getDriveId(graphClient, siteId);
      const sharePointUrl = await uploadToSharePoint(
        graphClient, siteId, driveId, fileName, buffer, file.type, folder
      );

      context.log(`✅ Uploaded: ${sharePointUrl}`);

      return {
        status: 200,
        jsonBody: { success: true, fileName, sharePointUrl }
      };

    } catch (err) {
      context.log(`❌ Error: ${err.message}`);
      return { status: 500, jsonBody: { error: err.message } };
    }
  }
});
