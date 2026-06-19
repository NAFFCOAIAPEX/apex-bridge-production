// Migrate apex-data-backup.json → Dataverse via apex-bridge
// Run: node migrate-to-dataverse.mjs

import { readFileSync } from 'fs';

const APEX_BRIDGE_SAVE_URL = 'https://apex-bridge-gtamg9f8d3a9afc8.westeurope-01.azurewebsites.net/api/saveEstimationData';
const BACKUP_FILE = 'apex-data-backup.json';

async function migrate() {
  console.log('🚀 Starting migration to Dataverse...\n');

  // Read backup file
  const raw = readFileSync(BACKUP_FILE, 'utf8');
  const data = JSON.parse(raw);

  console.log(`📦 Data loaded:`);
  console.log(`   Requests:     ${data.requests.length}`);
  console.log(`   DiaryEntries: ${data.diaryEntries.length}`);
  console.log(`   File size:    ${(raw.length / 1024).toFixed(1)} KB\n`);

  console.log('📤 Pushing to Dataverse via apex-bridge...');

  const response = await fetch(APEX_BRIDGE_SAVE_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(data),
  });

  if (!response.ok) {
    const error = await response.text();
    console.log(`❌ Failed: ${error}`);
    return;
  }

  const result = await response.json();
  console.log(`✅ Migration complete!`);
  console.log(`   Success:  ${result.success}`);
  console.log(`   Saved at: ${result.savedAt}`);
  console.log(`\n🎉 All ${data.requests.length} requests saved to Dataverse!`);
}

migrate().catch(console.error);
