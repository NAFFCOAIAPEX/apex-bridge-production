// Full Document Migration: Azure Blob -> SharePoint, then update Dataverse
// Run: node full-doc-migration.mjs

import { readFileSync, writeFileSync, existsSync } from "fs";

const APEX_BRIDGE_MIGRATE_URL = 'https://apex-bridge-gtamg9f8d3a9afc8.westeurope-01.azurewebsites.net/api/migrateAzureToSharePoint';
const APEX_BRIDGE_GET_URL     = 'https://apex-bridge-gtamg9f8d3a9afc8.westeurope-01.azurewebsites.net/api/getEstimationData';
const APEX_BRIDGE_SAVE_URL    = 'https://apex-bridge-gtamg9f8d3a9afc8.westeurope-01.azurewebsites.net/api/saveEstimationData';

const AZURE_BLOB_BASE = 'https://apexfilestorage2.blob.core.windows.net/estimation-docs/';

const PROGRESS_FILE = 'doc-migration-progress.json';

let progress = { mapping: {}, failed: [] };
if (existsSync(PROGRESS_FILE)) {
  progress = JSON.parse(readFileSync(PROGRESS_FILE, 'utf8'));
  console.log(`Resuming - ${Object.keys(progress.mapping).length} already migrated, ${progress.failed.length} previously failed`);
}

function extractUrls(obj, urls = new Set()) {
  if (Array.isArray(obj)) {
    obj.forEach(item => extractUrls(item, urls));
  } else if (obj && typeof obj === 'object') {
    for (const [k, v] of Object.entries(obj)) {
      if (k === 'url' && typeof v === 'string' && v.includes('blob.core.windows.net') && !v.endsWith('/')) {
        urls.add(v);
      } else {
        extractUrls(v, urls);
      }
    }
  }
  return urls;
}

function replaceUrls(obj, mapping) {
  if (Array.isArray(obj)) {
    return obj.map(item => replaceUrls(item, mapping));
  } else if (obj && typeof obj === 'object') {
    const result = {};
    for (const [k, v] of Object.entries(obj)) {
      if (k === 'url' && typeof v === 'string' && mapping[v]) {
        result[k] = mapping[v];
      } else {
        result[k] = replaceUrls(v, mapping);
      }
    }
    return result;
  }
  return obj;
}

async function migrateFile(blobUrl) {
  const afterContainer = blobUrl.replace(AZURE_BLOB_BASE, '');
  const decoded = decodeURIComponent(afterContainer);
  const parts = decoded.split('/');
  const fileName = parts.pop();
  const folder = parts.join('/');

  const response = await fetch(APEX_BRIDGE_MIGRATE_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ blobPath: decoded, fileName, folder }),
  });

  if (!response.ok) {
    const err = await response.text();
    throw new Error(err.substring(0, 150));
  }
  return await response.json();
}

async function main() {
  console.log('Starting full document migration...\n');

  const dataRes = await fetch(APEX_BRIDGE_GET_URL);
  const data = await dataRes.json();
  console.log(`Requests: ${data.requests.length}`);

  const urls = Array.from(extractUrls(data.requests));
  console.log(`Unique blob URLs to migrate: ${urls.length}\n`);

  let success = 0, failed = 0, skipped = 0;

  for (let i = 0; i < urls.length; i++) {
    const url = urls[i];
    if (progress.mapping[url]) { skipped++; continue; }

    const shortName = decodeURIComponent(url.split('/').pop()).substring(0, 50);
    process.stdout.write(`[${i+1}/${urls.length}] ${shortName}... `);

    try {
      const result = await migrateFile(url);
      progress.mapping[url] = result.sharePointUrl;
      success++;
      console.log('OK');
    } catch (err) {
      progress.failed = progress.failed.filter(f => f.url !== url);
      progress.failed.push({ url, error: err.message });
      failed++;
      console.log(`FAIL: ${err.message}`);
    }

    if ((success + failed) % 5 === 0) {
      writeFileSync(PROGRESS_FILE, JSON.stringify(progress, null, 2));
    }

    await new Promise(r => setTimeout(r, 300));
  }

  writeFileSync(PROGRESS_FILE, JSON.stringify(progress, null, 2));

  console.log(`\nDone! Success: ${success}, Skipped(already done): ${skipped}, Failed: ${failed}`);

  if (success > 0) {
    console.log('\nUpdating Dataverse with new URLs...');
    const updatedRequests = replaceUrls(data.requests, progress.mapping);
    const saveRes = await fetch(APEX_BRIDGE_SAVE_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ requests: updatedRequests, diaryEntries: data.diaryEntries }),
    });
    const saveResult = await saveRes.json();
    console.log('Dataverse updated:', JSON.stringify(saveResult));
  }

  if (failed > 0) {
    console.log(`\n${failed} files failed - re-run script to retry. Check ${PROGRESS_FILE} for details.`);
  } else if (skipped + success === urls.length) {
    console.log('\nAll files migrated successfully!');
  }
}

main().catch(console.error);
