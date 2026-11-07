const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function backend() {
  const saved = [];
  const status = [];
  const sheet = { getLastRow: () => 2, getRange: () => ({ setValue: value => status.push(value) }) };
  const context = vm.createContext({
    console,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({ getSheetByName: name => name === 'review_queue' ? null : sheet }),
      getUi: () => ({ alert() {} })
    },
    Utilities: { getUuid: () => 'test' }
  });
  vm.runInContext(fs.readFileSync('apps-script/Smore.gs', 'utf8'), context);
  Object.assign(context, {
    setupSiteSheet() {},
    smoreHeaders_: (sheet, columns) => Array.from(columns),
    appendRecordByHeaders_: (sheet, headers, record) => saved.push(record),
    appendAuditLog_() {},
    smoreJson_: value => value
  });
  return { context, saved, status };
}

test('submission saves the image separately from resource links and rejects invalid URLs', () => {
  const { context, saved } = backend();
  const payload = { submission_type: 'article', title: 'Title', content: 'Body',
    links: 'https://example.com/resource', cover_image_url: ' https://drive.google.com/file/d/image/view ' };
  assert.equal(context.doPost({ parameter: payload }).ok, true);
  assert.equal(saved[0].cover_image_url, payload.cover_image_url.trim());
  assert.equal(saved[0].links, payload.links);
  assert.equal(saved[0].status, 'pending');
  assert.equal(context.doPost({ parameter: { ...payload, cover_image_url: 'javascript:alert(1)' } }).ok, false);
  assert.equal(saved.length, 1);
  assert.equal(context.doPost({ parameter: { ...payload, cover_image_url: '' } }).ok, true);
});

test('approval carries images into articles and preserves older image-free submissions', () => {
  for (const image of ['https://example.com/image.jpg', undefined]) {
    const { context, saved, status } = backend();
    context.smoreReadIssues_ = () => [{ row: 2, issue: {
      id: 'submission', status: 'approved', submission_type: 'article',
      title: 'Title', content: 'Body', cover_image_url: image
    } }];
    context.publishApprovedSubmissions();
    assert.equal(saved[0].cover_image_url, image || '');
    assert.equal(saved[0].article_body_text, 'Body');
    assert.equal(saved[0].is_live, 'TRUE');
    assert.equal(status[0], 'published');
  }
});
