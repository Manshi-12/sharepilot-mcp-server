const selfsigned = require('selfsigned');
const fs = require('fs');

(async () => {
  const attrs = [{ name: 'commonName', value: 'SharePilot-MCP-App' }];
  const pems = await selfsigned.generate(attrs, { days: 730, keySize: 2048 });

  fs.writeFileSync('sharepilot-key.pem', pems.private);
  fs.writeFileSync('sharepilot-cert.pem', pems.cert);

  console.log('Done — created sharepilot-key.pem and sharepilot-cert.pem');
})();