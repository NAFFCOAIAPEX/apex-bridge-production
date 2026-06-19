const { app } = require('@azure/functions');
const { ClientSecretCredential } = require('@azure/identity');
const { Client } = require('@microsoft/microsoft-graph-client');
const { TokenCredentialAuthenticationProvider } = require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');
const { S3Client, GetObjectCommand, DeleteObjectCommand } = require('@aws-sdk/client-s3');
const { BlobServiceClient } = require('@azure/storage-blob');
const TENANT_ID      = '6c5f12d0-7ab9-463d-b9df-7036176a0685';
const CLIENT_ID      = '5a7089e2-3370-4a02-8546-63e8d92d695c';
const CLIENT_SECRET  = 'D_O8Q~BxfzushXOqEcCX2cGeoT1LFsOTNNeZHbCa';
const SP_HOSTNAME    = 'naffcogroup.sharepoint.com';
const SP_SITE_PATH   = '/sites/AI-APEX';
const SP_FOLDER      = 'Estimation_Docs';
const R2_ACCOUNT_ID        = '13b2f661abaa4c2acce9dc45fe69b8cc';
const R2_ACCESS_KEY_ID     = '94f5152a6e045323d8cc18be33bbbdfc';
const R2_SECRET_ACCESS_KEY = '09e36f22255a22f5fac011be72d8dfb60a50ba2da3b6ece02b9f139a785af43d';
const R2_BUCKET            = 'estimation-docs';
const R2_PUBLIC_URL        = 'https://pub-a821f9113188484f80b16e7995fee042.r2.dev';
const JSONBIN_BIN_ID  = '69dcdFfeaaba882197f3c176';
const JSONBIN_API_KEY = '$2a$10$kpIFmWCwfUxqOw.M.TfqcOyhGnnArBzDluhGquW2s/t.L3vQJtBqW';
const DATAVERSE_URL   = 'https://org50114a58.crm4.dynamics.com';
const DATAVERSE_TABLE = 'cr8a9_estimationdata1s';

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization',
};

const r2 = new S3Client({
  region: 'auto',
  endpoint: `https://${R2_ACCOUNT_ID}.r2.cloudflarestorage.com`,
  credentials: { accessKeyId: R2_ACCESS_KEY_ID, secretAccessKey: R2_SECRET_ACCESS_KEY },
});

function getGraphClient() {
  const credential = new ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
  const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes: ['https://graph.microsoft.com/.default'] });
  return Client.initWithMiddleware({ authProvider });
}

async function getDataverseToken() {
  const credential = new ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
  const token = await credential.getToken(`${DATAVERSE_URL}/.default`);
  return token.token;
}

async function getSiteAndDrive(graphClient) {
  const site = await graphClient.api(`/sites/${SP_HOSTNAME}:${SP_SITE_PATH}`).get();
  const drives = await graphClient.api(`/sites/${site.id}/drives`).get();
  const drive = drives.value.find(d => d.name === 'Documents') || drives.value[0];
  return { siteId: site.id, driveId: drive.id };
}

async function uploadToSharePoint(graphClient, siteId, driveId, fileName, buffer, mimeType, folder) {
  const destFolder = folder ? `${SP_FOLDER}/${folder}` : SP_FOLDER;
  const uploadPath = `/sites/${siteId}/drives/${driveId}/root:/${destFolder}/${fileName}:/content`;
  const response = await graphClient.api(uploadPath).header('Content-Type', mimeType || 'application/octet-stream').put(buffer);
  return response.webUrl;
}

async function downloadFromR2(fileKey) {
  const response = await r2.send(new GetObjectCommand({ Bucket: R2_BUCKET, Key: fileKey }));
  const chunks = [];
  for await (const chunk of response.Body) chunks.push(chunk);
  return { buffer: Buffer.concat(chunks), contentType: response.ContentType };
}

async function deleteFromR2(fileKey) {
  await r2.send(new DeleteObjectCommand({ Bucket: R2_BUCKET, Key: fileKey }));
}

async function updateJsonBinUrl(oldUrl, newUrl) {
  const res = await fetch(`https://api.jsonbin.io/v3/b/${JSONBIN_BIN_ID}/latest`, { headers: { 'X-Master-Key': JSONBIN_API_KEY } });
  const { record } = await res.json();
  const updated = JSON.stringify(record).replaceAll(oldUrl, newUrl);
  await fetch(`https://api.jsonbin.io/v3/b/${JSONBIN_BIN_ID}`, { method: 'PUT', headers: { 'Content-Type': 'application/json', 'X-Master-Key': JSONBIN_API_KEY }, body: updated });
}

app.http('health', {
  methods: ['GET', 'OPTIONS'], authLevel: 'anonymous', route: 'health',
  handler: async (request, context) => {
    return { status: 200, headers: CORS, jsonBody: { status: 'apex-bridge is running', timestamp: new Date().toISOString() } };
  }
});

app.http('getEstimationData', {
  methods: ['GET', 'OPTIONS'], authLevel: 'anonymous', route: 'getEstimationData',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: CORS };
    try {
      const token = await getDataverseToken();
      const url = `${DATAVERSE_URL}/api/data/v9.2/${DATAVERSE_TABLE}?$filter=cr8a9_estimationname eq 'apex-data'&$select=cr8a9_jsondata,cr8a9_updatedat`;
      const response = await fetch(url, { headers: { 'Authorization': `Bearer ${token}`, 'Accept': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' } });
      if (!response.ok) { const err = await response.text(); return { status: response.status, headers: CORS, jsonBody: { error: err } }; }
      const data = await response.json();
      if (!data.value || data.value.length === 0) return { status: 200, headers: CORS, jsonBody: { requests: [], diaryEntries: [] } };
      const jsondata = data.value[0].cr8a9_jsondata || '{"requests":[],"diaryEntries":[]}';
      return { status: 200, headers: CORS, jsonBody: JSON.parse(jsondata) };
    } catch (err) { context.log(`Error: ${err.message}`); return { status: 500, headers: CORS, jsonBody: { error: err.message } }; }
  }
});

app.http('saveEstimationData', {
  methods: ['POST', 'OPTIONS'], authLevel: 'anonymous', route: 'saveEstimationData',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: CORS };
    try {
      const body = await request.json();
      const token = await getDataverseToken();
      const jsondata = JSON.stringify(body);
      const checkUrl = `${DATAVERSE_URL}/api/data/v9.2/${DATAVERSE_TABLE}?$filter=cr8a9_estimationname eq 'apex-data'&$select=cr8a9_estimationdata1id`;
      const checkData = await (await fetch(checkUrl, { headers: { 'Authorization': `Bearer ${token}`, 'Accept': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' } })).json();
      if (checkData.value && checkData.value.length > 0) {
        const recordId = checkData.value[0].cr8a9_estimationdata1id;
        const updateRes = await fetch(`${DATAVERSE_URL}/api/data/v9.2/${DATAVERSE_TABLE}(${recordId})`, { method: 'PATCH', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' }, body: JSON.stringify({ cr8a9_jsondata: jsondata, cr8a9_updatedat: new Date().toISOString() }) });
        if (!updateRes.ok) { const err = await updateRes.text(); return { status: updateRes.status, headers: CORS, jsonBody: { error: err } }; }
      } else {
        const createRes = await fetch(`${DATAVERSE_URL}/api/data/v9.2/${DATAVERSE_TABLE}`, { method: 'POST', headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' }, body: JSON.stringify({ cr8a9_estimationname: 'apex-data', cr8a9_datatype: 'main', cr8a9_jsondata: jsondata, cr8a9_updatedat: new Date().toISOString() }) });
        if (!createRes.ok) { const err = await createRes.text(); return { status: createRes.status, headers: CORS, jsonBody: { error: err } }; }
      }
      return { status: 200, headers: CORS, jsonBody: { success: true, savedAt: new Date().toISOString() } };
    } catch (err) { context.log(`Error: ${err.message}`); return { status: 500, headers: CORS, jsonBody: { error: err.message } }; }
  }
});

app.http('migrateToSharePoint', {
  methods: ['POST', 'OPTIONS'], authLevel: 'anonymous', route: 'migrateToSharePoint',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: CORS };
    try {
      const { fileUrl, fileName, mimeType, folder, requestId } = await request.json();
      if (!fileUrl || !fileName) return { status: 400, headers: CORS, jsonBody: { error: 'fileUrl and fileName required' } };
      const fileKey = fileUrl.replace(`${R2_PUBLIC_URL}/`, '');
      const { buffer, contentType } = await downloadFromR2(fileKey);
      const graphClient = getGraphClient();
      const { siteId, driveId } = await getSiteAndDrive(graphClient);
      const sharePointUrl = await uploadToSharePoint(graphClient, siteId, driveId, fileName, buffer, contentType || mimeType, folder || requestId || 'General');
      await updateJsonBinUrl(fileUrl, sharePointUrl);
      await deleteFromR2(fileKey);
      return { status: 200, headers: CORS, jsonBody: { success: true, fileName, sharePointUrl } };
    } catch (err) { context.log(`Error: ${err.message}`); return { status: 500, headers: CORS, jsonBody: { error: err.message } }; }
  }
});

app.http('uploadToSharePoint', {
  methods: ['POST', 'OPTIONS'], authLevel: 'anonymous', route: 'uploadToSharePoint',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: CORS };
    try {
      const formData = await request.formData();
      const file = formData.get('file');
      const folder = formData.get('folder') || 'General';
      if (!file) return { status: 400, headers: CORS, jsonBody: { error: 'No file provided' } };
      const buffer = Buffer.from(await file.arrayBuffer());
      const graphClient = getGraphClient();
      const { siteId, driveId } = await getSiteAndDrive(graphClient);
      const sharePointUrl = await uploadToSharePoint(graphClient, siteId, driveId, file.name, buffer, file.type, folder);
      return { status: 200, headers: CORS, jsonBody: { success: true, fileName: file.name, sharePointUrl } };
    } catch (err) { context.log(`Error: ${err.message}`); return { status: 500, headers: CORS, jsonBody: { error: err.message } }; }
  }
});
async function downloadFromAzureBlob(blobName) {
  const AZURE_CONNECTION_STRING = "DefaultEndpointsProtocol=https;AccountName=apexfilestorage2;AccountKey=dXpA3aaLOSRD2Xqq/jJ1GxZ2uOMgKaXX3A2V68tb/gyY5Td1qcUudmUiV8/pQpkroBJz9njCPb8x+AStKxUvRA==;EndpointSuffix=core.windows.net";
  const AZURE_CONTAINER = "estimation-docs";
  const blobServiceClient = BlobServiceClient.fromConnectionString(AZURE_CONNECTION_STRING);
  const containerClient = blobServiceClient.getContainerClient(AZURE_CONTAINER);
  const blobClient = containerClient.getBlobClient(blobName);
  const downloadResponse = await blobClient.download();
  const chunks = [];
  for await (const chunk of downloadResponse.readableStreamBody) chunks.push(chunk);
  return { buffer: Buffer.concat(chunks), contentType: downloadResponse.contentType };
}

app.http('migrateAzureToSharePoint', {
  methods: ['POST', 'OPTIONS'], authLevel: 'anonymous', route: 'migrateAzureToSharePoint',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: CORS };
    try {
      const { blobPath, fileName, folder } = await request.json();
      if (!blobPath || !fileName) return { status: 400, headers: CORS, jsonBody: { error: 'blobPath and fileName required' } };
      const { buffer, contentType } = await downloadFromAzureBlob(blobPath);
      const graphClient = getGraphClient();
      const { siteId, driveId } = await getSiteAndDrive(graphClient);
      const sharePointUrl = await uploadToSharePoint(graphClient, siteId, driveId, fileName, buffer, contentType, folder || 'General');
      return { status: 200, headers: CORS, jsonBody: { success: true, fileName, sharePointUrl } };
    } catch (err) {
      context.log(`Error: ${err.message}`);
      return { status: 500, headers: CORS, jsonBody: { error: err.message } };
    }
  }
});app.http('archiveCompletedRequests', {
  methods: ['POST', 'OPTIONS'], authLevel: 'anonymous', route: 'archiveCompletedRequests',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: CORS };
    try {
      const token = await getDataverseToken();
      const url = `${DATAVERSE_URL}/api/data/v9.2/${DATAVERSE_TABLE}?$filter=cr8a9_estimationname eq 'apex-data'&$select=cr8a9_jsondata,cr8a9_estimationdata1id`;
      const response = await fetch(url, { headers: { 'Authorization': `Bearer ${token}`, 'Accept': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' } });
      const data = await response.json();
      if (!data.value || data.value.length === 0) return { status: 200, headers: CORS, jsonBody: { success: true, archived: 0, remaining: 0 } };
      
      const jsondata = JSON.parse(data.value[0].cr8a9_jsondata || '{"requests":[],"diaryEntries":[]}');
      const recordId = data.value[0].cr8a9_estimationdata1id;
      
      // Split: completed vs active
      const completed = jsondata.requests.filter(r => ['Approved', 'Rejected', 'Delivered'].includes(r.status));
      const active = jsondata.requests.filter(r => !['Approved', 'Rejected', 'Delivered'].includes(r.status));
      
      // Save archived to SharePoint
      if (completed.length > 0) {
        const archiveFileName = `APEX_Archive_${new Date().toISOString().slice(0, 10)}.json`;
        const graphClient = getGraphClient();
        const { siteId, driveId } = await getSiteAndDrive(graphClient);
        await uploadToSharePoint(graphClient, siteId, driveId, archiveFileName, Buffer.from(JSON.stringify(completed)), 'application/json', 'Archives');
      }
      
      // Update Dataverse with only active requests
      const updated = JSON.stringify({ requests: active, diaryEntries: jsondata.diaryEntries || [] });
      await fetch(`${DATAVERSE_URL}/api/data/v9.2/${DATAVERSE_TABLE}(${recordId})`, { 
        method: 'PATCH', 
        headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' }, 
        body: JSON.stringify({ cr8a9_jsondata: updated }) 
      });
      
      return { status: 200, headers: CORS, jsonBody: { success: true, archived: completed.length, remaining: active.length } };
    } catch (err) { context.log(`Error: ${err.message}`); return { status: 500, headers: CORS, jsonBody: { error: err.message } }; }
  }
});

app.http('getEstimationDataWithArchive', {
  methods: ['GET', 'OPTIONS'], authLevel: 'anonymous', route: 'getEstimationDataWithArchive',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') return { status: 204, headers: CORS };
    try {
      const token = await getDataverseToken();
      
      // Get active from Dataverse
      const dvUrl = `${DATAVERSE_URL}/api/data/v9.2/${DATAVERSE_TABLE}?$filter=cr8a9_estimationname eq 'apex-data'&$select=cr8a9_jsondata`;
      const dvResponse = await fetch(dvUrl, { headers: { 'Authorization': `Bearer ${token}`, 'Accept': 'application/json', 'OData-MaxVersion': '4.0', 'OData-Version': '4.0' } });
      const dvData = await dvResponse.json();
      const active = dvData.value && dvData.value.length > 0 
        ? JSON.parse(dvData.value[0].cr8a9_jsondata || '{"requests":[]}').requests 
        : [];
      
      // Get archived from SharePoint
      let archived = [];
      try {
        const graphClient = getGraphClient();
        const { siteId, driveId } = await getSiteAndDrive(graphClient);
        const filesUrl = `/sites/${siteId}/drives/${driveId}/root:/Archives:/children`;
        const filesResp = await graphClient.api(filesUrl).get();
        
        if (filesResp.value) {
          for (const file of filesResp.value) {
            if (file.name.startsWith('APEX_Archive_')) {
              const contentUrl = `/sites/${siteId}/drives/${driveId}/items/${file.id}/content`;
              const content = await graphClient.api(contentUrl).get();
              const archivedRequests = JSON.parse(content);
              archived = archived.concat(archivedRequests);
            }
          }
        }
      } catch (err) { context.log(`Archive read warning: ${err.message}`); }
      
      // Merge active + archived
      const allRequests = [...active, ...archived];
      
      return { status: 200, headers: CORS, jsonBody: { requests: allRequests, diaryEntries: [] } };
    } catch (err) { context.log(`Error: ${err.message}`); return { status: 500, headers: CORS, jsonBody: { error: err.message } }; }
  }
});