import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import test from 'node:test';
import vm from 'node:vm';

const source = await fs.readFile(new URL('../public/app.js', import.meta.url), 'utf8');

test('prefers Feishu row checkboxes over the active cell', () => {
  const context = vm.createContext({});
  vm.runInContext(source.slice(source.indexOf('const selectionRecordIds'), source.indexOf('const layoutVariableValue')) + '\nthis.ids = { selectionRecordIds, currentReadRecordIds };', context);
  assert.deepEqual([...context.ids.currentReadRecordIds({ recordId: 'active' }, ['checked-1', 'checked-2'])], ['checked-1', 'checked-2']);
  assert.deepEqual([...context.ids.currentReadRecordIds({ recordId: 'active' }, [])], ['active']);
});

test('selection polling and record loading include checked IDs without scanning the view', () => {
  assert.match(source, /view\.getSelectedRecordIdList\(\)/);
  assert.match(source, /selectedRecordIds\.join\(','\)/);
  assert.match(source, /currentReadRecordIds\(current, checked\)/);
  assert.match(source, /activeRecordFields\(\)\.map/);
  assert.match(source, /getFieldList\(\)\)\.map\(basicFieldName\)/);
  assert.doesNotMatch(source, /getDependencyFields|getRecords\(\{ pageSize: 5000|getRecordListByPage/);
  assert.doesNotMatch(source, /batchRecord|batchDialog|batchColumn|batchPage/);
});

test('keeps linked records, exact placeholder matching and field diagnostics', () => {
  assert.match(source, /linkedSchemaCache\.get\(tableId\)/);
  assert.match(source, /item\.link_record_ids/);
  assert.match(source, /loops\[templateField\.name\] = linked/);
  assert.match(source, /legacyPlaceholderName/);
  assert.match(source, /const templateFieldDiagnostics/);
  assert.match(source, /字段 ID/);
  assert.doesNotMatch(source, /lineTotal \/ quantity/);
});
