const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const page = fs.readFileSync('index.html', 'utf8');
const start = page.indexOf('    // BACKEND PAGE CONTENT START');
const end = page.indexOf('    // BACKEND PAGE CONTENT END');
const code = page.slice(start, end);

function frontend(settings = {}) {
  const root = { innerHTML: '', children: [], appendChild(el) { this.children.push(el); this.innerHTML += el.innerHTML; } };
  const status = { textContent: '', className: '', hidden: null, classList: { toggle(_, on) { status.hidden = on; } } };
  const context = vm.createContext({
    URL, window: {},
    document: {
      getElementById: id => ({ academySections: root, academyStatus: status })[id] || null,
      createElement: () => ({ set innerHTML(v) { this._html = v; }, get innerHTML() { return this._html; } })
    }
  });
  vm.runInContext(`
    const escapeHtml = v => String(v ?? '').replace(/[&<>"']/g, c => '&#' + c.charCodeAt(0) + ';');
    const sanitizeHttpUrl = v => { try { const u = new URL(String(v).trim()); return /^https?:$/.test(u.protocol) ? u.href : ''; } catch { return ''; } };
    let siteSettings = ${JSON.stringify(settings)};
    let siteLoadFailed = false;
    const APP_DATA = { techTools: [] };
    function renderTechTools() {}
    ${code}
    globalThis.api = { parseTimestamp, videoEmbedUrl, normalizePageContent, renderAcademyPage,
      setFailed: v => { siteLoadFailed = v; }, setAvailable: v => { pageContentAvailable = v; },
      load: data => { pageContent = normalizePageContent(data); pageContentAvailable = true; } };
  `, context);
  return { api: context.api, root, status };
}

test('timestamps parse m:ss, h:mm:ss and seconds, and reject junk', () => {
  const { api } = frontend();
  assert.equal(api.parseTimestamp('1:25'), 85);
  assert.equal(api.parseTimestamp('1:02:03'), 3723);
  assert.equal(api.parseTimestamp('90'), 90);
  assert.equal(api.parseTimestamp('soon'), null);
});

test('video links become embeds with chapter start times; unsupported hosts do not', () => {
  const { api } = frontend();
  assert.equal(api.videoEmbedUrl('https://www.youtube.com/watch?v=abcdefghijk', 85),
    'https://www.youtube.com/embed/abcdefghijk?rel=0&start=85');
  assert.equal(api.videoEmbedUrl('https://youtu.be/abcdefghijk', 0), 'https://www.youtube.com/embed/abcdefghijk?rel=0');
  assert.equal(api.videoEmbedUrl('https://vimeo.com/12345', 5), 'https://player.vimeo.com/video/12345#t=5s');
  assert.equal(api.videoEmbedUrl('https://drive.google.com/file/d/abc_1/view', 5), 'https://drive.google.com/file/d/abc_1/preview');
  assert.equal(api.videoEmbedUrl('https://example.com/video', 0), '');
  assert.equal(api.videoEmbedUrl('javascript:alert(1)', 0), '');
});

test('sections and blocks follow sort_order and render only supported types', () => {
  const { api, root } = frontend();
  api.load({
    page_sections: [
      { section_id: 'b', page: 'tech_ai', kind: 'blocks', title: 'Second', sort_order: '20' },
      { section_id: 'a', page: 'tech_ai', kind: 'blocks', title: 'First', sort_order: '10' }
    ],
    content_blocks: [
      { block_id: '1', section_id: 'a', block_type: 'text', body: '<b>hi</b>' },
      { block_id: '2', section_id: 'a', block_type: 'script', body: 'bad' },
      { block_id: '3', section_id: 'b', block_type: 'prompt', title: 'P', body: 'Copy me' }
    ],
    videos: [], video_chapters: []
  });
  api.renderAcademyPage();
  assert.ok(root.innerHTML.indexOf('First') < root.innerHTML.indexOf('Second'));
  assert.ok(!root.innerHTML.includes('<b>hi</b>'));
  assert.ok(!root.innerHTML.includes('bad'));
  assert.match(root.innerHTML, /data-prompt-id="3"/);
});

test('no video row means no panel; chapters only appear under a configured video', () => {
  const { api, root } = frontend();
  api.load({
    page_sections: [{ section_id: 'v', page: 'tech_ai', kind: 'videos' }],
    content_blocks: [],
    videos: [{ video_id: 'x', section_id: 'v', url: '' }],
    video_chapters: [{ chapter_id: 'c', video_id: 'x', title: 'Orphan' }]
  });
  api.renderAcademyPage();
  assert.equal(root.children.length, 0);

  api.load({
    page_sections: [{ section_id: 'v', page: 'tech_ai', kind: 'videos' }],
    content_blocks: [],
    videos: [{ video_id: 'x', section_id: 'v', url: 'https://youtu.be/abcdefghijk', title: 'Demo' }],
    video_chapters: []
  });
  api.renderAcademyPage();
  assert.match(root.innerHTML, /<iframe/);
  assert.ok(!root.innerHTML.includes('video-chapter-btn'));

  api.load({
    page_sections: [{ section_id: 'v', page: 'tech_ai', kind: 'videos' }],
    content_blocks: [],
    videos: [{ video_id: 'x', section_id: 'v', url: 'https://youtu.be/abcdefghijk' }],
    video_chapters: [{ chapter_id: 'c', video_id: 'x', title: 'Intro', start_time: '0:00' }]
  });
  api.renderAcademyPage();
  assert.match(root.innerHTML, /data-chapter-id="c"/);
});

test('intentionally empty shows the configured message; load failure shows an error', () => {
  const empty = frontend({ page_empty_message: 'Nothing yet.' });
  empty.api.load({ page_sections: [], content_blocks: [], videos: [], video_chapters: [] });
  empty.api.renderAcademyPage();
  assert.equal(empty.status.textContent, 'Nothing yet.');
  assert.doesNotMatch(empty.status.className, /red/);

  const blank = frontend({ page_empty_message: '' });
  blank.api.load({ page_sections: [], content_blocks: [], videos: [], video_chapters: [] });
  blank.api.renderAcademyPage();
  assert.equal(blank.status.textContent, '');

  const failed = frontend();
  failed.api.setFailed(true);
  failed.api.renderAcademyPage();
  assert.match(failed.status.textContent, /could not be loaded/);
  assert.match(failed.status.className, /red/);
});

test('Apps Script serves page tabs, hides unpublished rows, and errors on a missing tab', () => {
  const src = fs.readFileSync('apps-script/Smore.gs', 'utf8');
  const tabs = {
    page_sections: [['section_id', 'page', 'is_live'], ['s1', 'tech_ai', 'TRUE'], ['s2', 'tech_ai', 'FALSE'], ['s3', 'tech_ai', '']]
  };
  const context = vm.createContext({
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: name => tabs[name] && {
      getLastRow: () => tabs[name].length, getDataRange: () => ({ getValues: () => tabs[name] }) } }) }
  });
  vm.runInContext(`${src}; globalThis.items = publicTableItems_;`, context);
  const rows = context.items('page_sections', ['section_id', 'page', 'is_live'], ['section_id'], false, true);
  assert.deepEqual(Array.from(rows, row => row.section_id), ['s1']);
  assert.throws(() => context.items('videos', ['video_id', 'is_live'], ['video_id'], false, true), /Missing tab "videos"/);
});
