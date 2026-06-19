// Bulk Migration Script — Cloudflare R2 → SharePoint via apex-bridge
// Run: node bulk-migrate.mjs

const APEX_BRIDGE_URL = 'https://apex-bridge-gtamg9f8d3a9afc8.westeurope-01.azurewebsites.net/api/migrateToSharePoint';
const R2_PUBLIC_URL   = 'https://pub-a821f9113188484f80b16e7995fee042.r2.dev';

const R2_ACCOUNT_ID        = '13b2f661abaa4c2acce9dc45fe69b8cc';
const R2_ACCESS_KEY_ID     = '94f5152a6e045323d8cc18be33bbbdfc';
const R2_SECRET_ACCESS_KEY = '09e36f22255a22f5fac011be72d8dfb60a50ba2da3b6ece02b9f139a785af43d';
const R2_BUCKET            = 'estimation-docs';

import { S3Client, ListObjectsV2Command } from '@aws-sdk/client-s3';
import { writeFileSync, readFileSync, existsSync } from 'fs';

const r2 = new S3Client({
  region: 'auto',
  endpoint: `https://${R2_ACCOUNT_ID}.r2.cloudflarestorage.com`,
  credentials: {
    accessKeyId: R2_ACCESS_KEY_ID,
    secretAccessKey: R2_SECRET_ACCESS_KEY,
  },
});

// Load progress file (resume if interrupted)
const PROGRESS_FILE = 'migration-progress.json';
let progress = { completed: [], failed: [] };
if (existsSync(PROGRESS_FILE)) {
  progress = JSON.parse(readFileSync(PROGRESS_FILE, 'utf8'));
  console.log(`📋 Resuming migration — ${progress.completed.length} already done`);
}

async function listAllR2Files() {
  const files = [];
  let continuationToken = undefined;

  do {
    const command = new ListObjectsV2Command({
      Bucket: R2_BUCKET,
      ContinuationToken: continuationToken,
    });
    const response = await r2.send(command);
    if (response.Contents) {
      files.push(...response.Contents);
    }
    continuationToken = response.NextContinuationToken;
  } while (continuationToken);

  return files;
}

async function migrateFile(file) {
  const key = file.Key;
  const fileUrl = `${R2_PUBLIC_URL}/${key}`;
  const fileName = key.split('/').pop();
  const folder = key.includes('/') ? key.split('/').slice(0, -1).join('/') : 'General';

  const response = await fetch(APEX_BRIDGE_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ fileUrl, fileName, folder, requestId: folder }),
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`${response.status}: ${error}`);
  }

  return await response.json();
}

async function bulkMigrate() {
  console.log('🚀 Starting Bulk Migration: R2 → SharePoint\n');

  // List all files in R2
  console.log('📦 Listing all files in R2...');
  const files = await listAllR2Files();
  console.log(`   Found: ${files.length} files\n`);

  let success = 0;
  let failed = 0;
  let skipped = 0;

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const key = file.Key;

    // Skip already completed
    if (progress.completed.includes(key)) {
      skipped++;
      continue;
    }

    process.stdout.write(`[${i + 1}/${files.length}] ${key} ... `);

    try {
      const result = await migrateFile(file);
      progress.completed.push(key);
      success++;
      console.log(`✅`);
    } catch (err) {
      progress.failed.push({ key, error: err.message });
      failed++;
      console.log(`❌ ${err.message}`);
    }

    // Save progress every 10 files
    if ((success + failed) % 10 === 0) {
      writeFileSync(PROGRESS_FILE, JSON.stringify(progress, null, 2));
    }

    // Small delay to avoid throttling
    await new Promise(r => setTimeout(r, 500));
  }

  // Final save
  writeFileSync(PROGRESS_FILE, JSON.stringify(progress, null, 2));

  console.log(`\n✅ Migration Complete!`);
  console.log(`   Success: ${success} files`);
  console.log(`   Skipped: ${skipped} files (already done)`);
  console.log(`   Failed:  ${failed} files`);

  if (failed > 0) {
    console.log(`\n❌ Failed files saved in: ${PROGRESS_FILE}`);
    console.log(`   Run script again to retry failed files`);
  } else {
    console.log(`\n🎉 All files migrated to SharePoint!`);
    console.log(`   You can now disable Cloudflare R2 public access`);
  }
}

bulkMigrate().catch(console.error);
