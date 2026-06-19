// Check which document files still exist in Azure Blob
// Run: node check-blob-files.mjs

import { BlobServiceClient } from "@azure/storage-blob";
import { readFileSync, writeFileSync } from "fs";

const AZURE_CONNECTION_STRING = "DefaultEndpointsProtocol=https;AccountName=apexfilestorage2;AccountKey=dXpA3aaLOSRD2Xqq/jJ1GxZ2uOMgKaXX3A2V68tb/gyY5Td1qcUudmUiV8/pQpkroBJz9njCPb8x+AStKxUvRA==;EndpointSuffix=core.windows.net";
const AZURE_CONTAINER = "estimation-docs";

async function main() {
  const blobServiceClient = BlobServiceClient.fromConnectionString(AZURE_CONNECTION_STRING);
  const containerClient = blobServiceClient.getContainerClient(AZURE_CONTAINER);

  console.log("📦 Listing all blobs in container...\n");

  let count = 0;
  const blobNames = new Set();

  for await (const blob of containerClient.listBlobsFlat()) {
    blobNames.add(blob.name);
    count++;
  }

  console.log(`✅ Total blobs found in container: ${count}\n`);

  // Save list
  writeFileSync('existing_blobs.txt', Array.from(blobNames).sort().join('\n'));
  console.log('Saved blob list to existing_blobs.txt');

  // Sample first 10
  console.log('\nSample blobs:');
  Array.from(blobNames).slice(0, 10).forEach(n => console.log(' -', n));
}

main().catch(console.error);
