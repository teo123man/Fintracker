const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
vm.runInThisContext(fs.readFileSync(`${__dirname}/upload_client.js`, 'utf8'));
const sample = () => ({name: 'ÖRNEK HAREKETLERİ ekim.pdf', size: 4,
                       arrayBuffer: async () => new Uint8Array([37, 80, 68, 70]).buffer});
const reply = (status, result) => ({ok: status < 400, status, json: async () => result});

test('uploads bytes and a safely encoded Turkish filename', async () => {
  const result = await uploadStatement(sample(), 'test-token', async (path, options) => {
    assert.equal(path, '/api/upload');
    assert.equal(options.headers['X-Token'], 'test-token');
    assert.equal(decodeURIComponent(options.headers['X-Filename']), sample().name);
    assert.ok(options.body instanceof ArrayBuffer);
    assert.deepEqual([...new Uint8Array(options.body)], [37, 80, 68, 70]);
    return reply(200, {ok: true});
  });
  assert.equal(result.ok, true);
});
test('a local file read failure does not make a network request', async () => {
  const file = sample();
  file.arrayBuffer = async () => { throw new Error('NotReadableError'); };
  await assert.rejects(uploadStatement(file, 'token', () => assert.fail('must not fetch')), /dosya okunamadı/);
});
test('a disconnected app gets an actionable error', async () => {
  await assert.rejects(uploadStatement(sample(), 'token', async () => { throw new TypeError('Failed to fetch'); }), /Dosya okunabildi ancak/);
});
test('server failures are never reported as copied files', async () => {
  await assert.rejects(uploadStatement(sample(), 'token', async () => reply(500, {ok: false, log: 'Disk dolu'})), /Disk dolu/);
});
test('stale tabs are told to open a fresh session', async () => {
  await assert.rejects(uploadStatement(sample(), 'token', async () => reply(403, {ok: false})), /oturumu sona ermiş/);
});
test('oversized files are rejected before reading', async () => {
  const file = sample();
  file.size = 41 * 1024 * 1024;
  file.arrayBuffer = () => assert.fail('must not read');
  await assert.rejects(uploadStatement(file, 'token'), /40 MB/);
});
