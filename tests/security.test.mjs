import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';

const code = fs.readFileSync(new URL('../src/Code.gs', import.meta.url), 'utf8');
const index = fs.readFileSync(new URL('../src/index.html', import.meta.url), 'utf8');
const scripts = fs.readFileSync(new URL('../src/Scripts.html', import.meta.url), 'utf8');
const styles = fs.readFileSync(new URL('../src/Styles.html', import.meta.url), 'utf8');
const manifest = JSON.parse(fs.readFileSync(new URL('../src/appsscript.json', import.meta.url), 'utf8'));
const pkg = JSON.parse(fs.readFileSync(new URL('../package.json', import.meta.url), 'utf8'));
const readme = fs.readFileSync(new URL('../README.md', import.meta.url), 'utf8');

function backendContext() {
  const context = { console };
  vm.createContext(context);
  vm.runInContext(code, context);
  return context;
}

test('spreadsheet text and template JSON neutralize injection payloads', () => {
  const context = backendContext();
  assert.equal(context.asSheetText_('=IMPORTXML("https://example.com")'), "'=IMPORTXML(\"https://example.com\")");
  assert.equal(context.asSheetText_('+1-555-0100'), "'+1-555-0100");
  assert.equal(context.asSheetText_('Northline'), 'Northline');
  assert.doesNotMatch(context.safeJsonForScript_({ message: '</script><script>alert(1)</script>' }), /<\/script/i);
});

test('server API has no arbitrary dispatcher and every data endpoint is role-gated', () => {
  assert.doesNotMatch(code, /function\s+runCentral\s*\(/);
  assert.doesNotMatch(scripts, /\.runCentral\s*\(/);
  const readEndpoints = ['getSettings', 'getClients', 'getPackingSlips', 'getSuggestions', 'getClientInvoiceNumbers', 'duplicateSlip', 'getOrders', 'getLinkedDeliveryStats', 'getLinkedOrderMetadata', 'getSlipPDFEnrichment'];
  const writeEndpoints = ['saveClient', 'savePackingSlip', 'saveOrder', 'deleteOrder'];
  const adminEndpoints = ['getUsers', 'saveUser', 'deleteUser', 'saveSettings', 'getSystemDeploymentInfo'];
  for (const name of readEndpoints) assert.match(code, new RegExp(`function ${name}\\([^)]*\\) \\{\\s*requireReadAccess_\\(\\);`), name);
  for (const name of writeEndpoints) assert.match(code, new RegExp(`function ${name}\\([^)]*\\) \\{\\s*requireWriteAccess_\\(\\);`), name);
  for (const name of adminEndpoints) assert.match(code, new RegExp(`function ${name}\\([^)]*\\) \\{\\s*requireAdminAccess_\\(\\);`), name);
});

test('deployment uses least privilege and pinned browser dependencies', () => {
  assert.doesNotMatch(code, /DriveApp|XFrameOptionsMode\.ALLOWALL/);
  assert.ok(!manifest.oauthScopes.includes('https://www.googleapis.com/auth/drive'));
  assert.doesNotMatch(index, /lucide@latest/);
  assert.doesNotMatch(styles, /@import\s+url\([^)]*fonts\.googleapis\.com/);
  assert.equal(pkg.license, 'MIT');
  assert.equal(pkg.private, true);
});

test('repository instructions are portable and complete', () => {
  assert.doesNotMatch(readme, /file:\/\//);
  assert.match(readme, /OrderController/);
  assert.match(readme, /InvoiceController/);
  assert.match(readme, /appsscript\.json/);
});
