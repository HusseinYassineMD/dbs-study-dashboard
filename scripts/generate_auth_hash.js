#!/usr/bin/env node
/** Print SHA-256 hex for a dashboard password. Usage: node scripts/generate_auth_hash.js "secret" */
const crypto = require("crypto");
const password = process.argv[2];
if (!password) {
  console.error("Usage: node scripts/generate_auth_hash.js \"YourPassword\"");
  process.exit(1);
}
console.log(crypto.createHash("sha256").update(password, "utf8").digest("hex"));
