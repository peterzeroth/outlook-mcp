#!/usr/bin/env node
require('dotenv').config();
const fs = require('fs');
const os = require('os');
const path = require('path');

const CLIENT_ID     = (process.env.MS_CLIENT_ID     || '').trim();
const CLIENT_SECRET = (process.env.MS_CLIENT_SECRET || '').trim();
const TENANT_ID     = (process.env.MS_TENANT_ID     || 'common').trim();
const TOKEN_PATH    = path.join(os.homedir(), '.outlook-mcp-tokens.json');

const SCOPES = [
  'offline_access', 'User.Read',
  'Mail.Read', 'Mail.ReadWrite', 'Mail.Send',
  'Calendars.Read', 'Calendars.ReadWrite',
  'Files.Read', 'Files.ReadWrite',
].join(' ');

if (!CLIENT_ID || !CLIENT_SECRET) {
  console.error('MS_CLIENT_ID or MS_CLIENT_SECRET missing from .env');
  process.exit(1);
}

console.log(`client_id:     ${CLIENT_ID}`);
console.log(`client_secret: ${CLIENT_SECRET.length} chars, ends with "${CLIENT_SECRET.slice(-4)}"`);
console.log(`tenant_id:     ${TENANT_ID}\n`);

async function post(url, params, debug = false) {
  const body = Object.entries(params)
    .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
    .join('&');
  if (debug) console.log('\nPOST body:', body, '\n');
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
  });
  return res.json();
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

async function main() {
  const deviceRes = await post(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/devicecode`,
    { client_id: CLIENT_ID, scope: SCOPES }
  );

  if (deviceRes.error) {
    console.error('Device code error:', deviceRes.error_description);
    process.exit(1);
  }

  console.log('='.repeat(60));
  console.log(deviceRes.message);
  console.log('='.repeat(60));
  console.log('\nWaiting for sign-in...\n');

  const interval = (deviceRes.interval || 5) * 1000;
  const expires  = Date.now() + (deviceRes.expires_in || 900) * 1000;

  while (Date.now() < expires) {
    await sleep(interval);
    const tokenRes = await post(
      `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
      {
        client_id:  CLIENT_ID,
        grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
        device_code: deviceRes.device_code,
      },
      true
    );

    if (tokenRes.error === 'authorization_pending') { process.stdout.write('.'); continue; }
    if (tokenRes.error) {
      console.error('\nAuth error:', tokenRes.error_description);
      process.exit(1);
    }

    const tokens = {
      access_token:  tokenRes.access_token,
      refresh_token: tokenRes.refresh_token,
      expires_at:    Date.now() + (tokenRes.expires_in || 3600) * 1000,
    };
    fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens, null, 2), { mode: 0o600 });
    console.log(`\n\nAuthenticated! Tokens saved to ${TOKEN_PATH}`);
    console.log('Restart Claude Code / Claude Desktop to activate the m365-assistant MCP.\n');
    process.exit(0);
  }

  console.error('Timed out.');
  process.exit(1);
}

main().catch(e => { console.error(e); process.exit(1); });
