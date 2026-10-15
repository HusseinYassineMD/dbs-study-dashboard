/**
 * Copy to auth.config.js (gitignored) to set dashboard login locally.
 *
 * Generate passwordHash:
 *   node scripts/generate_auth_hash.js "YourPasswordHere"
 *
 * Production (GitHub Pages): set repo secrets
 *   DASHBOARD_AUTH_USER
 *   DASHBOARD_AUTH_PASSWORD
 * The deploy workflow writes auth.config.js automatically.
 */
window.DBS_AUTH = {
  username: "dbs-operations",
  passwordHash: "a85c9ab08e5d4218b4b8699aa575844d57b96d3e41a5f29b2a939db7e4bda19d",
};
