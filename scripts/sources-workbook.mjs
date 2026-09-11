import crypto from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const XLSX = require('xlsx');
const root = path.resolve(import.meta.dirname, '..');
const workbookPath = path.resolve(root, process.env.COMBAT_ATLAS_WORKBOOK || 'source/combat_atlas_main.xlsm');
const outputPath = path.resolve(root, 'app/generated-sources.json');
const manifestPath = path.resolve(root, 'app/source-manifest.json');
const mode = process.argv.includes('--verify-generated') ? 'verify-generated' : process.argv.includes('--check') ? 'check' : process.argv.includes('--validate') ? 'validate' : 'sync';

const headers = {
  status: '发布状态', title: '中文标题', originalTitle: '原标题', author: '作者', publisher: '来源', url: '原文 URL',
  year: '发布年份', language: '语言', medium: '媒介', tier: '来源层级', topics: '主题', summary: '简介',
  takeaways: '核心结论', readingTime: '阅读时间（分钟）', curatorScore: '策展价值（1-5）', featured: '精选状态',
  addedAt: '收录日期', id: 'sourceId',
};
const enums = {
  status: new Set(['发布', '下架']), language: new Set(['中文', '英文', '日文']), medium: new Set(['文章', '视频', '演讲', '书籍']),
  tier: new Set(['一手', '精选二手']), featured: new Set(['是', '否']),
};

function text(value) {
  if (value == null) return '';
  if (typeof value === 'object') {
    if ('text' in value) return String(value.text ?? '').trim();
    if ('hyperlink' in value) return String(value.hyperlink ?? '').trim();
    if (Array.isArray(value.richText)) return value.richText.map((part) => part.text).join('').trim();
  }
  return String(value).trim();
}
function integer(value, label, row, min, max, errors) {
  const number = typeof value === 'number' ? value : Number(text(value));
  if (!Number.isInteger(number) || number < min || number > max) errors.push(`第 ${row} 行“${label}”必须是 ${min}-${max} 的整数`);
  return number;
}
function isoDate(value, row, errors) {
  if (value instanceof Date && !Number.isNaN(value.valueOf())) return `${value.getUTCFullYear()}-${String(value.getUTCMonth() + 1).padStart(2, '0')}-${String(value.getUTCDate()).padStart(2, '0')}`;
  const valueText = text(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(valueText) && !Number.isNaN(Date.parse(`${valueText}T00:00:00Z`))) return valueText;
  errors.push(`第 ${row} 行“收录日期”必须是有效日期`);
  return valueText;
}
function uniqueList(value, separator) { return [...new Set(text(value).split(separator).map((item) => item.trim()).filter(Boolean))]; }
function serialize(document) { return `${JSON.stringify(document, null, 2)}\n`; }
function digest(serialized) { return crypto.createHash('sha256').update(serialized, 'utf8').digest('hex'); }
function makeManifest(document, serialized) {
  return { schemaVersion: 1, generatorVersion: 2, recordCount: document.records.length, publishedCount: document.records.filter((record) => record.publicationStatus === 'published').length, dataSha256: digest(serialized) };
}
async function readJson(file, fallback) { try { return JSON.parse(await fs.readFile(file, 'utf8')); } catch { return fallback; } }

function validateDocument(document) {
  const errors = [];
  if (document?.schemaVersion !== 1 || !Array.isArray(document.records)) return ['数据文档结构无效'];
  const ids = new Set();
  for (const [index, record] of document.records.entries()) {
    const row = index + 1;
    for (const key of ['id', 'title', 'author', 'publisher', 'url', 'language', 'medium', 'tier', 'summary', 'addedAt']) if (!text(record[key])) errors.push(`记录 ${row} 缺少 ${key}`);
    if (ids.has(record.id)) errors.push(`sourceId 重复：${record.id}`);
    ids.add(record.id);
    try { const url = new URL(record.url); if (!['http:', 'https:'].includes(url.protocol)) throw new Error(); } catch { errors.push(`URL 无效：${record.id}`); }
    if (!['published', 'unpublished'].includes(record.publicationStatus)) errors.push(`发布状态无效：${record.id}`);
    if (!['中文', '英文', '日文'].includes(record.language)) errors.push(`语言无效：${record.id}`);
    if (!['文章', '视频', '演讲', '书籍'].includes(record.medium)) errors.push(`媒介无效：${record.id}`);
    if (!['一手', '精选二手'].includes(record.tier)) errors.push(`来源层级无效：${record.id}`);
    if (!Number.isInteger(record.year) || record.year < 1900 || record.year > 2100) errors.push(`年份无效：${record.id}`);
    if (!Number.isInteger(record.readingTime) || record.readingTime < 1) errors.push(`阅读时间无效：${record.id}`);
    if (!Number.isInteger(record.curatorScore) || record.curatorScore < 1 || record.curatorScore > 5) errors.push(`展示价值无效：${record.id}`);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(record.addedAt)) errors.push(`收录日期无效：${record.id}`);
    if (!Array.isArray(record.topics) || !record.topics.length || !Array.isArray(record.takeaways) || !record.takeaways.length) errors.push(`主题或核心结论为空：${record.id}`);
  }
  return errors;
}

if (mode === 'verify-generated') {
  const document = await readJson(outputPath, null);
  const errors = validateDocument(document);
  const serialized = serialize(document);
  const expected = makeManifest(document, serialized);
  const actual = await readJson(manifestPath, null);
  if (JSON.stringify(actual) !== JSON.stringify(expected)) errors.push('source-manifest.json 与生成数据不一致');
  if (errors.length) throw new Error(`公开数据校验失败：\n- ${errors.join('\n- ')}`);
  console.log(JSON.stringify(expected));
  process.exit(0);
}

const previous = await readJson(outputPath, { schemaVersion: 1, records: [] });
const workbook = XLSX.readFile(workbookPath, { cellDates: true });
const sheet = workbook.Sheets['知识库文章'];
if (!sheet) throw new Error('工作簿缺少“知识库文章”页签');
const range = XLSX.utils.decode_range(sheet['!ref']);
const cellValue = (row, column) => {
  const cell = sheet[XLSX.utils.encode_cell({ r: row - 1, c: column - 1 })];
  if (cell?.t === 'd' && /^\d{4}-\d{2}-\d{2}$/.test(cell.w || '')) return cell.w;
  return cell?.v;
};
let headerRow = 0;
let columns = new Map();
for (let candidate = range.s.r + 1; candidate <= Math.min(range.e.r + 1, 30); candidate += 1) {
  const map = new Map();
  for (let column = range.s.c + 1; column <= range.e.c + 1; column += 1) map.set(text(cellValue(candidate, column)), column);
  if (Object.values(headers).every((header) => map.has(header))) { headerRow = candidate; columns = map; break; }
}
if (!headerRow) throw new Error(`找不到完整表头：${Object.values(headers).join('、')}`);

const errors = [];
const records = [];
const ids = new Set();
for (let rowNumber = headerRow + 1; rowNumber <= range.e.r + 1; rowNumber += 1) {
  const raw = Object.fromEntries(Object.entries(headers).map(([key, header]) => [key, cellValue(rowNumber, columns.get(header))]));
  if (!Object.values(raw).some((value) => text(value))) continue;
  for (const key of ['status', 'title', 'author', 'publisher', 'url', 'language', 'medium', 'tier', 'topics', 'summary', 'takeaways', 'featured', 'id']) if (!text(raw[key])) errors.push(`第 ${rowNumber} 行“${headers[key]}”不能为空`);
  for (const key of Object.keys(enums)) if (text(raw[key]) && !enums[key].has(text(raw[key]))) errors.push(`第 ${rowNumber} 行“${headers[key]}”值无效：${text(raw[key])}`);
  const id = text(raw.id);
  if (id && ids.has(id)) errors.push(`第 ${rowNumber} 行 sourceId 重复：${id}`);
  ids.add(id);
  let url = text(raw.url);
  try { const parsed = new URL(url); if (!['https:', 'http:'].includes(parsed.protocol)) throw new Error(); url = parsed.href; } catch { errors.push(`第 ${rowNumber} 行“原文 URL”无效：${url}`); }
  records.push({
    publicationStatus: text(raw.status) === '发布' ? 'published' : 'unpublished', id, title: text(raw.title), ...(text(raw.originalTitle) ? { originalTitle: text(raw.originalTitle) } : {}),
    author: text(raw.author), publisher: text(raw.publisher), url, year: integer(raw.year, headers.year, rowNumber, 1900, 2100, errors),
    language: text(raw.language), medium: text(raw.medium), tier: text(raw.tier), topics: uniqueList(raw.topics, /[；;\n]+/), summary: text(raw.summary),
    takeaways: uniqueList(raw.takeaways, /\r?\n+/), readingTime: integer(raw.readingTime, headers.readingTime, rowNumber, 1, 10000, errors),
    curatorScore: integer(raw.curatorScore, headers.curatorScore, rowNumber, 1, 5, errors), featured: text(raw.featured) === '是', addedAt: isoDate(raw.addedAt, rowNumber, errors),
  });
}
const document = { schemaVersion: 1, records };
errors.push(...validateDocument(document));
const previousById = new Map((previous.records || []).map((record) => [record.id, record]));
const currentById = new Map(records.map((record) => [record.id, record]));
const removed = [...previousById.keys()].filter((id) => !currentById.has(id));
if (removed.length) errors.push(`禁止删除或修改既有 sourceId，请改为“下架”：${removed.join('、')}`);
if (!records.length) errors.push('知识库文章不能为空');
if (errors.length) throw new Error(`工作簿校验失败：\n- ${[...new Set(errors)].join('\n- ')}`);

const serialized = serialize(document);
const manifest = makeManifest(document, serialized);
const previousSerialized = serialize(previous);
const added = records.filter((record) => !previousById.has(record.id));
const changed = records.filter((record) => previousById.has(record.id) && JSON.stringify(previousById.get(record.id)) !== JSON.stringify(record));
const downlisted = changed.filter((record) => previousById.get(record.id)?.publicationStatus === 'published' && record.publicationStatus === 'unpublished');
const report = { total: records.length, published: manifest.publishedCount, added: added.length, modified: changed.length, downlisted: downlisted.length, highRisk: downlisted.length > 0 || changed.length > 5, dataSha256: manifest.dataSha256 };
if (mode === 'check') {
  const existingManifest = await readJson(manifestPath, null);
  if (serialized !== previousSerialized || JSON.stringify(existingManifest) !== JSON.stringify(manifest)) throw new Error('本地主表与公开派生数据不一致，请点击工作簿中的“发布到网页”');
} else if (mode === 'sync') {
  await fs.writeFile(outputPath, serialized, 'utf8');
  await fs.writeFile(manifestPath, `${JSON.stringify(manifest, null, 2)}\n`, 'utf8');
}
console.log(JSON.stringify(report));
