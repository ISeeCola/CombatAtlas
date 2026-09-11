import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const XLSX = require('xlsx');

export const REVIEW_HEADERS = ['候选 ID','sourceId','审核状态','发现日期','中文标题','原标题','作者','来源','原文 URL','规范 URL','平台 ID','发布年份','语言','媒介','建议来源层级','建议主题','证据状态','短读证据','深读证据','去重结果','置信度','预筛-相关度','预筛-信息密度','预筛-新颖性','预筛-来源可信度','预筛-证据可得性','预筛总分','正式-主题价值','正式-方法迁移','正式-证据案例','正式-库内增量','正式-可操作启发','正式-来源可信度','正式-结构价值','正式总分','总评','摘要','核心方法','证据与案例','术语对照','短摘录','适用边界','库内连接','实验落点','可迁移假设','最小验证动作','观察指标','失败信号','归档路径','阅读时间（分钟）','展示价值（1-5）','精选状态','入库日期','审核备注','评论'];
export const REVIEW_STATUSES = ['待复核','待修改','已接受','已入库','已拒绝','重复'];
const allowedStatuses = new Set(REVIEW_STATUSES);
const advisoryFields = ['深读证据','去重结果','适用边界','可迁移假设','最小验证动作','观察指标','失败信号','归档路径'];
const autoAdviceMarker = '【自动审核建议】';

export function clean(value) { return value == null ? '' : String(value).trim(); }
function argumentValue(name) { const index = process.argv.indexOf(name); return index >= 0 ? process.argv[index + 1] : ''; }
function writeReport(reportPath, payload) { if (reportPath) fs.writeFileSync(reportPath, JSON.stringify({ version: 2, ...payload }, null, 2), 'utf8'); }
function findHeaderRow(matrix) {
  const required = ['候选 ID','sourceId','审核状态','中文标题','规范 URL','审核备注','评论'];
  const row = matrix.findIndex((cells) => required.every((header) => cells.some((cell) => clean(cell) === header)));
  if (row < 0) throw new Error('未找到 CombatAtlasInbox 必需表头');
  return row;
}
function validHttpUrl(value) { try { return ['http:', 'https:'].includes(new URL(value).protocol); } catch { return false; } }
function labelFor(row, index) {
  const title = clean(row['中文标题']) || '未命名条目';
  const sourceId = clean(row.sourceId);
  return `记录 ${index + 1}「${title}」${sourceId ? ` (${sourceId})` : ''}`;
}

export function readReview(workbookPath = process.env.COMBAT_ATLAS_REVIEW || path.resolve('source/combat_atlas_review.xlsm')) {
  if (!fs.existsSync(workbookPath)) throw new Error(`找不到审核工作簿：${workbookPath}`);
  const workbook = XLSX.readFile(workbookPath, { cellDates: true, cellFormula: true });
  const sheet = workbook.Sheets['待审批文章'] ?? workbook.Sheets[workbook.SheetNames[0]];
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
  const headerRow = findHeaderRow(matrix);
  const headers = matrix[headerRow].map(clean);
  const rows = matrix.slice(headerRow + 1).filter((cells) => cells.some((cell) => clean(cell))).map((cells) => Object.fromEntries(headers.map((header, index) => [header, cells[index] ?? ''])));
  return { workbookPath, headers, headerRow: headerRow + 1, rows };
}

// Machine integrity faults only. Editorial quality never blocks a human acceptance.
export function validateReview(rows) {
  const errors = [];
  const candidateIds = new Set();
  const sourceIds = new Set();
  for (const [index, row] of rows.entries()) {
    const label = labelFor(row, index);
    const candidateId = clean(row['候选 ID']);
    const sourceId = clean(row.sourceId);
    const status = clean(row['审核状态']);
    const canonicalUrl = clean(row['规范 URL']);
    if (!candidateId) errors.push(`${label}：缺少候选 ID`);
    else if (candidateIds.has(candidateId)) errors.push(`${label}：候选 ID 重复：${candidateId}`);
    else candidateIds.add(candidateId);
    if (sourceId) {
      if (sourceIds.has(sourceId)) errors.push(`${label}：sourceId 重复：${sourceId}`);
      else sourceIds.add(sourceId);
    }
    if (!allowedStatuses.has(status)) errors.push(`${label}：审核状态非法：${status || '空值'}`);
    if (!clean(row['中文标题'])) errors.push(`${label}：缺少中文标题`);
    if (!canonicalUrl) errors.push(`${label}：缺少规范 URL`);
    else if (!validHttpUrl(canonicalUrl)) errors.push(`${label}：规范 URL 非法：${canonicalUrl}`);
    const preText = clean(row['预筛总分']);
    const pre = Number(preText);
    if (preText && (!Number.isFinite(pre) || pre < 0 || pre > 30)) errors.push(`${label}：预筛总分必须在 0–30`);
    const formalText = clean(row['正式总分']);
    const formal = Number(formalText);
    if (formalText && (!Number.isFinite(formal) || formal < 0 || formal > 100)) errors.push(`${label}：正式总分必须在 0–100`);
  }
  return [...new Set(errors)];
}

export function collectReviewAdvisories(rows) {
  const advisories = [];
  for (const [index, row] of rows.entries()) {
    if (!['已接受','已入库'].includes(clean(row['审核状态']))) continue;
    const issues = [];
    const formal = Number(clean(row['正式总分']));
    if (!Number.isFinite(formal) || formal < 70) issues.push('正式总分未达到 70');
    const missing = advisoryFields.filter((field) => !clean(row[field]));
    if (missing.length) issues.push(`建议补充：${missing.join('、')}`);
    if (issues.length) advisories.push({ candidateId: clean(row['候选 ID']), sourceId: clean(row.sourceId), label: labelFor(row, index), issues, text: `${autoAdviceMarker}${issues.join('；')}` });
  }
  return advisories;
}

export function mergeAutomaticAdvice(existing, advice) {
  const human = clean(existing).replace(/(?:\r?\n)?【自动审核建议】[^\r\n]*/g, '').trim();
  return [human, clean(advice)].filter(Boolean).join('\n');
}

if (process.argv.includes('--validate')) {
  const { rows, workbookPath } = readReview();
  const errors = validateReview(rows);
  const advisories = collectReviewAdvisories(rows);
  if (errors.length) { console.error(`审核技术校验未通过：\n${errors.join('\n')}`); process.exit(1); }
  const counts = Object.groupBy(rows, (row) => clean(row['审核状态']));
  console.log(JSON.stringify({ workbookPath, total: rows.length, statuses: Object.fromEntries(Object.entries(counts).map(([key, value]) => [key, value.length])), advisoryCount: advisories.length }));
}

if (process.argv.includes('--accepted-json')) {
  const outputPath = argumentValue('--accepted-json');
  const reportPath = argumentValue('--report-json');
  if (!outputPath) throw new Error('--accepted-json 需要输出路径');
  const { rows, headers } = readReview();
  const errors = validateReview(rows);
  const advisoryByCandidate = new Map(collectReviewAdvisories(rows).map((item) => [item.candidateId, item]));
  if (errors.length) {
    writeReport(reportPath, { ok: false, title: '入库技术校验未通过：', errors, warnings: [], accepted: 0 });
    process.exit(2);
  }
  const accepted = rows.filter((row) => clean(row['审核状态']) === '已接受').map((row) => {
    const advice = advisoryByCandidate.get(clean(row['候选 ID']))?.text ?? '';
    return { candidateId: clean(row['候选 ID']), sourceId: clean(row.sourceId), finalReviewNote: mergeAutomaticAdvice(row['审核备注'], advice), advisory: advice, row };
  });
  const columns = Object.fromEntries(['候选 ID','sourceId','审核状态','审核备注'].map((header) => [header, headers.indexOf(header) + 1]));
  fs.writeFileSync(outputPath, JSON.stringify({ version: 2, columns, accepted }, null, 2), 'utf8');
  writeReport(reportPath, { ok: true, title: '技术校验通过', errors: [], warnings: accepted.filter((item) => item.advisory).map((item) => `${item.row['中文标题']}：${item.advisory}`), accepted: accepted.length });
  console.log(JSON.stringify({ accepted: accepted.length, advisoryCount: accepted.filter((item) => item.advisory).length, outputPath }));
}
