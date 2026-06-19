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

const R2_ACCOUNT_ID        = '13b2f661abaa4c2acce9dc45fe69b8cc';
const R2_ACCESS_KEY_ID     = '94f5152a6e045323d8cc18be33bbbdfc';
const R2_SECRET_ACCESS_KEY = '09e36f22255a22f5fac011be72d8dfb60a50ba2da3b6ece02b9f139a785af43d';
const R2_BUCKET            = 'estimation-docs';
const R2_PUBLIC_URL        = 'https://pub-a821f9113188484f80b16e7995fee042.r2.dev';

// Dataverse config
const DATAVERSE_URL   = 'https://org50114a58.crm4.dynamics.com';
const TABLE_NAME      = 'cr8a9_estimationdata1s';
const RECORD_NAME     = 'apex-data';
const FIELD_NAME      = 'cr8a9_jsondata';
const NAME_FIELD      = 'cr8a9_estimationname';

// SharePoint archive config
const SP_ARCHIVE_FOLDER = 'Archives';
const SP_ARCHIVE_FILE   = 'APEX_Archive_2026-06-19.json';
// ─────────────────────────────────────────────────────────────────────────────

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

// Get Drive ID for library
async function getDriveId(graphClient, siteId) {
  const drives = await graphClient.api(`/sites/${siteId}/drives`).get();
  const drive = drives.value.find(d => d.name === SHAREPOINT_LIBRARY);
  if (!drive) throw new Error(`Library '${SHAREPOINT_LIBRARY}' not found`);
  return drive.id;
}

// Upload file to SharePoint
async function uploadToSharePoint(graphClient, siteId, driveId, fileName, fileBuffer, mimeType, folder) {
  const folderPath = folder ? `/${folder}` : '';
  const uploadPath = `/sites/${siteId}/drives/${driveId}/root:/Estimation_Docs${folderPath}/${fileName}:/content`;
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

// Get Dataverse access token
async function getDataverseToken() {
  const credential = new ClientSecretCredential(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
  const token = await credential.getToken(`${DATAVERSE_URL}/.default`);
  return token.token;
}

// Batch operations with error handling per-request
async function batchOperations(items, operation, batchSize = 10) {
  const results = [];
  const errors = [];

  for (let i = 0; i < items.length; i += batchSize) {
    const batch = items.slice(i, i + batchSize);
    const promises = batch.map(item =>
      operation(item)
        .then(result => ({ success: true, data: result }))
        .catch(err => ({ success: false, item, error: err.message }))
    );
    const batchResults = await Promise.all(promises);
    batchResults.forEach(result => {
      if (result.success) results.push(result.data);
      else errors.push(result);
    });
  }

  return { results, errors };
}

// Get all requests from individual rows (skip 'apex-data' and 'diary-data')
async function getAllRequestsFromDataverse() {
  const token = await getDataverseToken();
  const filter = `$filter=${NAME_FIELD} ne 'apex-data' and ${NAME_FIELD} ne 'diary-data'&$select=${FIELD_NAME},${NAME_FIELD},cr8a9_estimationdata1id`;
  const url = `${DATAVERSE_URL}/api/data/v9.2/${TABLE_NAME}?${filter}`;
  const res = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${token}`,
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      'Accept': 'application/json',
    }
  });
  if (!res.ok) throw new Error(`Dataverse GET failed: ${res.status} ${await res.text()}`);
  const data = await res.json();

  const requests = (data.value || []).map(row => {
    try {
      const parsed = JSON.parse(row[FIELD_NAME] || '{}');
      return { ...parsed, _recordId: row.cr8a9_estimationdata1id };
    } catch (e) {
      console.error(`Failed to parse row ${row[NAME_FIELD]}:`, e);
      return null;
    }
  }).filter(r => r !== null);

  return requests;
}

// Get diary entries from special row
async function getDiaryEntriesFromDataverse() {
  const token = await getDataverseToken();
  const filter = `$filter=${NAME_FIELD} eq 'diary-data'&$select=${FIELD_NAME},cr8a9_estimationdata1id`;
  const url = `${DATAVERSE_URL}/api/data/v9.2/${TABLE_NAME}?${filter}`;
  const res = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${token}`,
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      'Accept': 'application/json',
    }
  });
  if (!res.ok) throw new Error(`Dataverse GET failed for diary: ${res.status} ${await res.text()}`);
  const data = await res.json();

  if (!data.value || data.value.length === 0) {
    return { entries: [], recordId: null };
  }

  const record = data.value[0];
  const parsed = JSON.parse(record[FIELD_NAME] || '{}');
  return {
    entries: Array.isArray(parsed) ? parsed : (parsed.entries || []),
    recordId: record.cr8a9_estimationdata1id
  };
}

// Save single request (new or update)
async function saveRequestToDataverse(request) {
  const token = await getDataverseToken();
  const requestId = request.id;

  // Check if request already exists
  const filter = `$filter=${NAME_FIELD} eq '${requestId.replace(/'/g, "''")}'&$select=cr8a9_estimationdata1id`;
  const checkUrl = `${DATAVERSE_URL}/api/data/v9.2/${TABLE_NAME}?${filter}`;
  const checkRes = await fetch(checkUrl, {
    headers: {
      'Authorization': `Bearer ${token}`,
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
      'Accept': 'application/json',
    }
  });

  if (!checkRes.ok) throw new Error(`Dataverse check failed: ${checkRes.status}`);
  const checkData = await checkRes.json();
  const existingRecord = checkData.value?.[0];

  // Prepare payload - remove internal fields
  const { _recordId, ...requestData } = request;
  const body = JSON.stringify({ [FIELD_NAME]: JSON.stringify(requestData) });

  if (existingRecord) {
    // Update existing record
    const recordId = existingRecord.cr8a9_estimationdata1id;
    const updateUrl = `${DATAVERSE_URL}/api/data/v9.2/${TABLE_NAME}(${recordId})`;
    const res = await fetch(updateUrl, {
      method: 'PATCH',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        'If-Match': '*',
      },
      body
    });
    if (!res.ok && res.status !== 204) {
      throw new Error(`Dataverse PATCH failed for ${requestId}: ${res.status} ${await res.text()}`);
    }
    return { action: 'updated', id: requestId, recordId };
  } else {
    // Create new record
    const createUrl = `${DATAVERSE_URL}/api/data/v9.2/${TABLE_NAME}`;
    const res = await fetch(createUrl, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
      },
      body: JSON.stringify({ [NAME_FIELD]: requestId, [FIELD_NAME]: JSON.stringify(requestData) })
    });
    if (!res.ok) {
      throw new Error(`Dataverse POST failed for ${requestId}: ${res.status} ${await res.text()}`);
    }
    const responseData = await res.json();
    return { action: 'created', id: requestId, recordId: responseData.cr8a9_estimationdata1id };
  }
}

// Save diary entries to special row
async function saveDiaryEntriesToDataverse(entries) {
  const token = await getDataverseToken();

  // Get existing diary row
  const { recordId } = await getDiaryEntriesFromDataverse();

  const body = JSON.stringify({ [FIELD_NAME]: JSON.stringify(entries) });

  if (recordId) {
    // Update existing
    const url = `${DATAVERSE_URL}/api/data/v9.2/${TABLE_NAME}(${recordId})`;
    const res = await fetch(url, {
      method: 'PATCH',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        'If-Match': '*',
      },
      body
    });
    if (!res.ok && res.status !== 204) {
      throw new Error(`Dataverse diary PATCH failed: ${res.status} ${await res.text()}`);
    }
    return { action: 'updated', recordId };
  } else {
    // Create new
    const url = `${DATAVERSE_URL}/api/data/v9.2/${TABLE_NAME}`;
    const res = await fetch(url, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
      },
      body: JSON.stringify({ [NAME_FIELD]: 'diary-data', [FIELD_NAME]: JSON.stringify(entries) })
    });
    if (!res.ok) {
      throw new Error(`Dataverse diary POST failed: ${res.status} ${await res.text()}`);
    }
    const responseData = await res.json();
    return { action: 'created', recordId: responseData.cr8a9_estimationdata1id };
  }
}

// Get estimation data from Dataverse (legacy - for backward compatibility with getEstimationDataWithArchive)
async function getFromDataverse() {
  const requests = await getAllRequestsFromDataverse();
  const { entries: diaryEntries } = await getDiaryEntriesFromDataverse();
  return { requests, diaryEntries, recordId: null };
}

// Save estimation data to Dataverse (legacy - routes to per-row saves)
async function saveToDataverse(payload, recordId) {
  const requests = payload.requests || [];
  const diaryEntries = payload.diaryEntries || [];

  // Save requests using batch operations
  const { results, errors } = await batchOperations(requests, saveRequestToDataverse, 10);

  if (diaryEntries.length > 0) {
    await saveDiaryEntriesToDataverse(diaryEntries);
  }

  if (errors.length > 0) {
    throw new Error(`Failed to save ${errors.length} requests: ${errors.map(e => e.error).join('; ')}`);
  }
}

// Get archived requests from SharePoint
async function getArchivedRequests(graphClient, siteId, driveId) {
  const filePath = `/sites/${siteId}/drives/${driveId}/root:/Estimation_Docs/${SP_ARCHIVE_FOLDER}/${SP_ARCHIVE_FILE}:/content`;
  const response = await graphClient.api(filePath).get();
  const text = Buffer.isBuffer(response) ? response.toString() : JSON.stringify(response);
  return JSON.parse(text);
}

// ─── FUNCTION 1: Health Check ─────────────────────────────────────────────────
app.http('health', {
  methods: ['GET'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    return {
      status: 200,
      headers: { 'Access-Control-Allow-Origin': '*' },
      jsonBody: { status: 'apex-bridge is running', timestamp: new Date().toISOString() }
    };
  }
});

// ─── FUNCTION 2: Get Estimation Data (active only) ───────────────────────────
app.http('getEstimationData', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') {
      return { status: 204, headers: { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET,POST,OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' } };
    }
    try {
      context.log('Fetching estimation data from Dataverse...');
      const { requests, diaryEntries } = await getFromDataverse();
      context.log(`Fetched ${requests.length} requests`);
      return {
        status: 200,
        headers: { 'Access-Control-Allow-Origin': '*', 'Content-Type': 'application/json' },
        jsonBody: { requests, diaryEntries }
      };
    } catch (err) {
      context.log(`Error: ${err.message}`);
      return {
        status: 500,
        headers: { 'Access-Control-Allow-Origin': '*' },
        jsonBody: { error: err.message }
      };
    }
  }
});

// ─── FUNCTION 3: Get Estimation Data With Archive (active + archived) ─────────
app.http('getEstimationDataWithArchive', {
  methods: ['GET', 'OPTIONS'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') {
      return { status: 204, headers: { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET,POST,OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' } };
    }
    try {
      // 1. Fetch active requests from Dataverse
      context.log('Fetching active requests from Dataverse...');
      const { requests: activeRequests, diaryEntries } = await getFromDataverse();
      context.log(`Active requests: ${activeRequests.length}`);

      // 2. Fetch archived requests from SharePoint
      let archivedRequests = [];
      try {
        const graphClient = getGraphClient();
        const siteId = await getSiteId(graphClient);
        const driveId = await getDriveId(graphClient, siteId);
        const archiveData = await getArchivedRequests(graphClient, siteId, driveId);
        archivedRequests = Array.isArray(archiveData) ? archiveData : (archiveData.requests || []);
        context.log(`Archived requests: ${archivedRequests.length}`);
      } catch (archiveErr) {
        context.log(`Warning: Could not fetch archive: ${archiveErr.message}`);
        archivedRequests = [];
      }

      // 3. Merge — active takes priority, no duplicates
      const activeIds = new Set(activeRequests.map(r => r.id));
      const uniqueArchived = archivedRequests.filter(a => !activeIds.has(a.id)).map(a => ({ ...a, archived: true }));
      const allRequests = [...activeRequests, ...uniqueArchived];

      context.log(`Total merged: ${allRequests.length}`);

      return {
        status: 200,
        headers: { 'Access-Control-Allow-Origin': '*', 'Content-Type': 'application/json' },
        jsonBody: {
          requests: allRequests,
          diaryEntries,
          summary: {
            total: allRequests.length,
            active: activeRequests.length,
            archived: uniqueArchived.length
          },
          timestamp: new Date().toISOString()
        }
      };
    } catch (err) {
      context.log(`Error: ${err.message}`);
      return {
        status: 500,
        headers: { 'Access-Control-Allow-Origin': '*' },
        jsonBody: { error: err.message }
      };
    }
  }
});

// ─── FUNCTION 4: Save Estimation Data ────────────────────────────────────────
app.http('saveEstimationData', {
  methods: ['POST', 'OPTIONS'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') {
      return { status: 204, headers: { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET,POST,OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' } };
    }
    try {
      const payload = await request.json();
      context.log(`Saving ${(payload.requests || []).length} requests to Dataverse...`);

      await saveToDataverse(payload, null);

      context.log('Saved successfully');
      return {
        status: 200,
        headers: { 'Access-Control-Allow-Origin': '*', 'Content-Type': 'application/json' },
        jsonBody: { success: true, timestamp: new Date().toISOString() }
      };
    } catch (err) {
      context.log(`Error: ${err.message}`);
      return {
        status: 500,
        headers: { 'Access-Control-Allow-Origin': '*' },
        jsonBody: { error: err.message }
      };
    }
  }
});

// ─── FUNCTION 4b: Migrate to Per-Row Storage ──────────────────────────────────
app.http('migrateToPerRow', {
  methods: ['POST', 'OPTIONS'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') {
      return { status: 204, headers: { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET,POST,OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' } };
    }
    try {
      const payload = await request.json();
      const backupJsonString = payload.backup || null;

      context.log('Starting migration from blob to per-row storage...');

      const token = await getDataverseToken();

      // 1. Read the old 'apex-data' blob row
      const filter = `$filter=${NAME_FIELD} eq 'apex-data'&$select=${FIELD_NAME},cr8a9_estimationdata1id`;
      const url = `${DATAVERSE_URL}/api/data/v9.2/${TABLE_NAME}?${filter}`;
      const res = await fetch(url, {
        headers: {
          'Authorization': `Bearer ${token}`,
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
          'Accept': 'application/json',
        }
      });

      if (!res.ok) throw new Error(`Failed to fetch old blob: ${res.status} ${await res.text()}`);
      const data = await res.json();

      let requests = [];
      let oldBlobRecordId = null;

      if (data.value && data.value.length > 0) {
        const oldRecord = data.value[0];
        oldBlobRecordId = oldRecord.cr8a9_estimationdata1id;
        const oldBlob = JSON.parse(oldRecord[FIELD_NAME] || '{}');
        requests = oldBlob.requests || [];
        context.log(`Found old blob with ${requests.length} requests`);
      }

      // 2. Load from backup if provided
      if (backupJsonString) {
        try {
          const backupData = JSON.parse(backupJsonString);
          const backupRequests = Array.isArray(backupData) ? backupData : (backupData.requests || []);
          const existingIds = new Set(requests.map(r => r.id));
          const newRequests = backupRequests.filter(r => !existingIds.has(r.id));
          requests.push(...newRequests);
          context.log(`Merged backup: +${newRequests.length} new requests`);
        } catch (e) {
          throw new Error(`Failed to parse backup JSON: ${e.message}`);
        }
      }

      if (requests.length === 0) {
        return {
          status: 200,
          headers: { 'Access-Control-Allow-Origin': '*', 'Content-Type': 'application/json' },
          jsonBody: { success: true, migratedCount: 0, message: 'No requests to migrate', timestamp: new Date().toISOString() }
        };
      }

      // 3. Create individual rows using batch operations
      const { results, errors } = await batchOperations(requests, saveRequestToDataverse, 10);

      const created = results.filter(r => r.action === 'created').length;
      const updated = results.filter(r => r.action === 'updated').length;

      context.log(`Migration complete: ${created} created, ${updated} updated, ${errors.length} errors`);

      return {
        status: 200,
        headers: { 'Access-Control-Allow-Origin': '*', 'Content-Type': 'application/json' },
        jsonBody: {
          success: errors.length === 0,
          migratedCount: requests.length,
          created,
          updated,
          errors: errors.length > 0 ? errors.map(e => ({ id: e.item.id, error: e.error })) : undefined,
          oldBlobRecordId: oldBlobRecordId,
          message: errors.length === 0 ? 'Migration completed successfully' : `Migration completed with ${errors.length} errors`,
          timestamp: new Date().toISOString()
        }
      };
    } catch (err) {
      context.log(`Migration error: ${err.message}`);
      return {
        status: 500,
        headers: { 'Access-Control-Allow-Origin': '*' },
        jsonBody: { error: err.message }
      };
    }
  }
});

// ─── FUNCTION 5: Upload file to SharePoint ────────────────────────────────────
app.http('uploadFile', {
  methods: ['POST', 'OPTIONS'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') {
      return { status: 204, headers: { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET,POST,OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' } };
    }
    try {
      const formData = await request.formData();
      const file = formData.get('file');
      const folder = formData.get('folder') || 'General';
      const customName = formData.get('fileName') || file.name;

      if (!file) return { status: 400, headers: { 'Access-Control-Allow-Origin': '*' }, jsonBody: { error: 'No file provided' } };

      const buffer = Buffer.from(await file.arrayBuffer());
      const graphClient = getGraphClient();
      const siteId = await getSiteId(graphClient);
      const driveId = await getDriveId(graphClient, siteId);
      const url = await uploadToSharePoint(graphClient, siteId, driveId, customName, buffer, file.type, folder);

      return {
        status: 200,
        headers: { 'Access-Control-Allow-Origin': '*', 'Content-Type': 'application/json' },
        jsonBody: { success: true, url, fileName: customName }
      };
    } catch (err) {
      context.log(`Upload error: ${err.message}`);
      return {
        status: 500,
        headers: { 'Access-Control-Allow-Origin': '*' },
        jsonBody: { error: err.message }
      };
    }
  }
});

// ─── FUNCTION 6: Migrate R2 → SharePoint ─────────────────────────────────────
app.http('migrateToSharePoint', {
  methods: ['POST', 'OPTIONS'],
  authLevel: 'anonymous',
  handler: async (request, context) => {
    if (request.method === 'OPTIONS') {
      return { status: 204, headers: { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Methods': 'GET,POST,OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type' } };
    }
    try {
      const { fileUrl, fileName, mimeType, folder, requestId } = await request.json();
      if (!fileUrl || !fileName) return { status: 400, headers: { 'Access-Control-Allow-Origin': '*' }, jsonBody: { error: 'fileUrl and fileName are required' } };

      const fileKey = fileUrl.replace(`${R2_PUBLIC_URL}/`, '');
      const { buffer, contentType } = await downloadFromR2(fileKey);
      const graphClient = getGraphClient();
      const siteId = await getSiteId(graphClient);
      const driveId = await getDriveId(graphClient, siteId);
      const sharePointUrl = await uploadToSharePoint(graphClient, siteId, driveId, fileName, buffer, contentType || mimeType, folder || requestId || 'General');
      await deleteFromR2(fileKey);

      return {
        status: 200,
        headers: { 'Access-Control-Allow-Origin': '*', 'Content-Type': 'application/json' },
        jsonBody: { success: true, fileName, sharePointUrl }
      };
    } catch (err) {
      context.log(`Migration error: ${err.message}`);
      return {
        status: 500,
        headers: { 'Access-Control-Allow-Origin': '*' },
        jsonBody: { error: err.message }
      };
    }
  }
});
