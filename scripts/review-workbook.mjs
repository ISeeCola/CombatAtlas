import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const XLSX = require('xlsx');

export const REVIEW_HEADERS = ['候选 ID','sourceId','审核状态','发现日期','中文标题','原标题','作者','来源','原文 URL','规范 URL','平台 ID','发布年份','语言','媒介','建议来源层级','建议主题','证据状态','短读证据','深读证据','去重结果','置信度','预筛-相关度','预筛-信息密度','预筛-新颖性','预筛-来源可信度','预筛-证据可得性','预筛总分','正式-主题价值','正式-方法迁移','正式-证据案例','正式-库内增量','正式-可操作启发','正式-来源可信度','正式-结构价值','正式总分','总评','摘要','核心方法','证据与案例','术语对照','短摘录','适用边界','库内连接','实验落点','可迁移假设','最小验证动作','观察指标','失败信号','归档路径','阅读时间（分钟）','展示价值（1-5）','精选状态','入库日期','审核备注'];
const allowedStatuses = new Set(['已发现','短名单','候选','待人工复核','已拒绝','已接受','已入库','重复']);

function clean(value) { return value == null ? '' : String(value).trim(); }
function findHeaderRow(matrix) {
  const row = matrix.findIndex((cells) => REVIEW_HEADERS.every((header) => cells.some((cell) => clean(cell) === header)));
  if (row < 0) throw new Error('未找到 CombatAtlasInbox 必需表头');
  return row;
}
export function readReview(workbookPath = process.env.COMBAT_ATLAS_REVIEW || path.resolve('source/combat_atlas_review.xlsm')) {
  if (!fs.existsSync(workbookPath)) throw new Error(`找不到审核工作簿：${workbookPath}`);
  const workbook = XLSX.readFile(workbookPath, { cellDates: true, cellFormula: true });
  const sheet = workbook.Sheets['待审批文章'] ?? workbook.Sheets[workbook.SheetNames[0]];
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
  const headerRow = findHeaderRow(matrix);
  const headers = matrix[headerRow].map(clean);
  const rows = matrix.slice(headerRow + 1).filter((cells) => cells.some((cell) => clean(cell))).map((cells) => Object.fromEntries(headers.map((header, index) => [header, cells[index] ?? ''])));
  return { workbookPath, rows };
}

export function validateReview(rows) {
  const errors = [];
  const candidateIds = new Set();
  const sourceIds = new Set();
  for (const [index, row] of rows.entries()) {
    const label = `第 ${index + 1} 条`;
    const candidateId = clean(row['候选 ID']);
    const sourceId = clean(row.sourceId);
    const status = clean(row['审核状态']);
    if (!candidateId) errors.push(`${label}缺少候选 ID`);
    else if (candidateIds.has(candidateId)) errors.push(`${label}候选 ID 重复：${candidateId}`);
    candidateIds.add(candidateId);
    if (sourceId) {
      if (sourceIds.has(sourceId)) errors.push(`${label} sourceId 重复：${sourceId}`);
      sourceIds.add(sourceId);
    }
    if (!allowedStatuses.has(status)) errors.push(`${label}审核状态非法：${status}`);
    if (!clean(row['中文标题']) || !clean(row['规范 URL'])) errors.push(`${label}缺少标题或规范 URL`);
    const pre = Number(row['预筛总分']);
    if (Number.isFinite(pre) && (pre < 0 || pre > 30)) errors.push(`${label}预筛总分超出 0–30`);
    if (status === '已接受' || status === '已入库') {
      const formal = Number(row['正式总分']);
      const required = ['深读证据','去重结果','适用边界','可迁移假设','最小验证动作','观察指标','失败信号','归档路径'];
      if (!Number.isFinite(formal) || formal < 70 || formal > 100) errors.push(`${label}正式评分未达 70 分门槛`);
      for (const field of required) if (!clean(row[field])) errors.push(`${label}已接受但缺少${field}`);
    }
  }
  return errors;
}

if (process.argv.includes('--validate')) {
  const { rows, workbookPath } = readReview();
  const errors = validateReview(rows);
  if (errors.length) { console.error(errors.join('\n')); process.exit(1); }
  const counts = Object.groupBy(rows, (row) => clean(row['审核状态']));
  console.log(JSON.stringify({ workbookPath, total: rows.length, statuses: Object.fromEntries(Object.entries(counts).map(([key, value]) => [key, value.length])) }));
}
if (process.argv.includes('--accepted-json')) {
  const outputIndex = process.argv.indexOf('--accepted-json') + 1;
  const outputPath = process.argv[outputIndex];
  if (!outputPath) throw new Error('--accepted-json 需要输出路径');
  const { rows } = readReview();
  const errors = validateReview(rows);
  if (errors.length) throw new Error(errors.join('\n'));
  const accepted = rows.filter((row) => String(row['审核状态']).trim() === '已接受');
  fs.writeFileSync(outputPath, JSON.stringify(accepted, null, 2));
  console.log(JSON.stringify({ accepted: accepted.length, outputPath }));
}
