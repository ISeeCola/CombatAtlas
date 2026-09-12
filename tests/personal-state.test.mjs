import assert from 'node:assert/strict';
import test from 'node:test';
import { parseAndMergeBackup, MAX_BACKUP_BYTES } from '../lib/personal-state.ts';

const backup = (items, exportedAt = '2026-09-12T00:00:00.000Z') => JSON.stringify({ app: 'CombatAtlas', version: 1, exportedAt, items });
const item = (sourceId, updatedAt = '2026-09-12T00:00:00.000Z') => ({ sourceId, rating: 4, isRead: true, updatedAt });

test('backup merge keeps the newest state', () => {
  const merged = parseAndMergeBackup(backup([item('src-a')]), { 'src-a': { rating: 2, isRead: false, updatedAt: '2026-09-11T00:00:00.000Z' } });
  assert.deepEqual(merged['src-a'], { rating: 4, isRead: true, updatedAt: '2026-09-12T00:00:00.000Z' });
});

test('backup import rejects prototype keys', () => {
  assert.throws(() => parseAndMergeBackup(backup([item('__proto__')]), {}));
  assert.throws(() => parseAndMergeBackup(backup([item('constructor')]), {}));
});

test('backup import rejects oversized payloads', () => {
  assert.throws(() => parseAndMergeBackup(' '.repeat(MAX_BACKUP_BYTES + 1), {}));
});
