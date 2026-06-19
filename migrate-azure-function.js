// ADD THIS to apex-bridge/index.js
// New endpoint: migrate directly from Azure Blob -> SharePoint

const { BlobServiceClient } = require('@azure/storage-blob');

const AZURE_CONNECTION_STRING = "DefaultEndpointsProtocol=https;AccountName=apexfilestorage2;AccountKey=dXpA3aaLOSRD2Xqq/jJ1GxZ2uOMgKaXX3A2V68tb/gyY5Td1qcUudmUiV8/pQpkroBJz9njCPb8x+AStKxUvRA==;EndpointSuffix=core.windows.net";
const AZURE_CONTAINER = "estimation-docs";

async function downloadFromAzureBlob(blobName) {
  const blobServiceClient = BlobServiceClient.fromConnectionString(AZURE_CONNECTION_STRING);
  const containerClient = blobServiceClient.getContainerClient(AZURE_CONTAINER);
  const blobClient = containerClient.getBlobClient(blobName);
  const downloadResponse = await blobClient.download();
  const chunks = [];
  for await (const chunk of downloadResponse.readableStreamBody) chunks.push(chunk);
  return { buffer: Buffer.concat(chunks), contentType: downloadResponse.contentType };
}

// ─── MIGRATE FROM AZURE BLOB TO SHAREPOINT ────────────────────────────────────
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
});
