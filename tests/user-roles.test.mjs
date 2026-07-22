import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';

const code = fs.readFileSync(new URL('../src/Code.gs', import.meta.url), 'utf8');
const scripts = fs.readFileSync(new URL('../src/Scripts.html', import.meta.url), 'utf8');
const index = fs.readFileSync(new URL('../src/index.html', import.meta.url), 'utf8');

function backendContext() {
  const context = { console };
  vm.createContext(context);
  vm.runInContext(code, context);
  return context;
}

test('duplicate user IDs are assigned stable unique replacements', () => {
  const { planUniqueUserIds_ } = backendContext();
  const result = planUniqueUserIds_([
    ['User ID', 'Name', 'Email', 'Role', 'Created At'],
    ['U-0001', 'Owner', 'owner@example.com', 'Admin', ''],
    ['U-0001', 'Pat', 'pat@example.com', 'User', ''],
    ['', 'Lee', 'lee@example.com', 'Viewer', ''],
    ['U-0010', 'Sam', 'sam@example.com', 'User', ''],
  ]);

  assert.deepEqual(Array.from(result.ids), ['U-0001', 'U-0011', 'U-0012', 'U-0010']);
  assert.equal(result.max, 12);
  assert.deepEqual(Array.from(result.changedRows), [3, 4]);
});

test('Viewer is available and non-admins never receive system deployment UI', () => {
  assert.match(index, /<option value="Viewer">Viewer/);
  assert.match(index, /id="system-info-card"/);
  assert.match(scripts, /systemCard\.style\.display = isAdmin \? 'block' : 'none'/);
  assert.match(scripts, /if \(isAdmin\) loadSystemDeploymentInfo\(\)/);
  const loadSettingsStart = scripts.indexOf('function loadSettings()');
  const loadSettingsEnd = scripts.indexOf('\n}', loadSettingsStart) + 2;
  assert.doesNotMatch(scripts.slice(loadSettingsStart, loadSettingsEnd), /loadSystemDeploymentInfo/);
});

test('Viewer is labelled separately from Editor in the user list', () => {
  assert.match(scripts, /u\.role === 'Viewer'/);
  assert.match(scripts, />Viewer<\/span>/);
  assert.match(scripts, />Editor<\/span>/);
});
