import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import sourceDocument from '../app/generated-sources.json' with { type: 'json' };
import { readReview, validateReview } from './review-workbook.mjs';

const root = path.resolve('..', 'docs', 'research', 'curated-game-design');
const review = readReview();
const errors = validateReview(review.rows);
if (errors.length) throw new Error(errors.join('\n'));
const byStatus = Object.groupBy(review.rows, (row) => String(row['审核状态']).trim());
const accepted = review.rows.filter((row) => ['已接受','已入库'].includes(String(row['审核状态']).trim()));
const esc = (value) => String(value ?? '').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;');
await fs.mkdir(path.join(root, 'items'), { recursive: true });

for (const row of accepted) {
  const id = String(row.sourceId).trim();
  const title = String(row['中文标题']).trim();
  const original = String(row['原标题']).trim();
  const markdown = `# ${title}\n\n- sourceId：\`${id}\`\n- 原标题：${original}\n- 作者 / 来源：${row['作者']} / ${row['来源']}\n- 正式评分：${row['正式总分']}/100\n- 证据：${row['深读证据']}\n- 规范 URL：${row['规范 URL']}\n\n## 摘要\n\n${row['摘要']}\n\n## 核心方法\n\n${row['核心方法']}\n\n## 适用边界\n\n${row['适用边界']}\n\n## 实验卡\n\n- 假设：${row['可迁移假设']}\n- 动作：${row['最小验证动作']}\n- 指标：${row['观察指标']}\n- 失败信号：${row['失败信号']}\n`;
  const bilingual = String(row['语言']).trim() !== '中文';
  const visual = `<svg viewBox="0 0 900 220" role="img" aria-label="CombatAtlas evidence map"><rect width="900" height="220" fill="#171715"/><circle cx="735" cy="110" r="82" fill="none" stroke="#ba503b" stroke-width="18"/><circle cx="735" cy="110" r="48" fill="none" stroke="#f3eee6" stroke-width="2"/><path d="M48 158H620" stroke="#f3eee6"/><path d="M48 168H440" stroke="#ba503b" stroke-width="5"/><text x="48" y="78" fill="#f3eee6" font-size="30" font-family="Arial">EVIDENCE / METHOD / TEST</text><text x="48" y="112" fill="#ba503b" font-size="16" font-family="Arial">COMBAT ATLAS · ${esc(id.toUpperCase())}</text></svg>`;
  const html = `<!doctype html><html lang="zh-CN"><meta charset="utf-8"><title>${esc(title)}</title><style>body{max-width:900px;margin:56px auto;padding:0 24px;background:#f3eee6;color:#181715;font:17px/1.75 system-ui}h1{font-size:42px;line-height:1.15}small{color:#835344}svg{width:100%;height:auto;margin:24px 0}section{border-top:1px solid #bbb;padding-top:18px;margin-top:28px}</style><h1>${esc(title)}</h1><p><small>${esc(original)} · ${esc(row['作者'])} · ${esc(row['来源'])} · ${esc(row['正式总分'])}/100</small></p>${visual}<p><a href="${esc(row['规范 URL'])}">访问原文${bilingual ? ' / Open original source' : ''}</a></p><section><h2>摘要${bilingual ? ' / Curated summary' : ''}</h2><p>${esc(row['摘要'])}</p></section><section><h2>核心方法${bilingual ? ' / Transferable methods' : ''}</h2><p>${esc(row['核心方法']).replaceAll('\n','<br>')}</p></section><section><h2>适用边界${bilingual ? ' / Scope limits' : ''}</h2><p>${esc(row['适用边界'])}</p></section><section><h2>验证实验${bilingual ? ' / Validation experiment' : ''}</h2><p>${esc(row['可迁移假设'])}</p><p>${esc(row['最小验证动作'])}</p></section></html>`;
  await fs.writeFile(path.join(root, 'items', `${id}.md`), markdown, 'utf8');
  await fs.writeFile(path.join(root, 'items', `${id}.html`), html, 'utf8');
}

const catalogRows = review.rows.map((row) => `| ${row.sourceId || '—'} | ${row['中文标题']} | ${row['审核状态']} | ${row['正式总分'] || '—'} | ${row['来源']} |`).join('\n');
await fs.writeFile(path.join(root, 'catalog.md'), `# CombatAtlas 策展目录\n\n网页展示 ${sourceDocument.records.length} 条；审核 accepted ${accepted.length} 条；已入库 ${(byStatus['已入库'] ?? []).length} 条。\n\n| sourceId | 标题 | 审核状态 | 正式分 | 来源 |\n|---|---|---:|---:|---|\n${catalogRows}\n`, 'utf8');
const pending = review.rows.filter((row) => !['已接受','已入库','已拒绝','重复'].includes(String(row['审核状态']).trim()));
await fs.writeFile(path.join(root, 'inbox.md'), `# 待处理摘要\n\n生成时间：2026-09-11（Asia/Hong_Kong）\n\n${pending.map((row) => `- **${row['中文标题']}** — ${row['审核状态']}；${row['审核备注']}`).join('\n')}\n`, 'utf8');
const digest = crypto.createHash('sha256').update(JSON.stringify(review.rows)).digest('hex');
await fs.writeFile(path.join(root, 'update-history.md'), `# 更新历史\n\n- 2026-09-11：全量审核 33 条；已接受 ${accepted.length}，候选 ${(byStatus['候选'] ?? []).length}，待人工复核 ${(byStatus['待人工复核'] ?? []).length}，重复 ${(byStatus['重复'] ?? []).length}。审核数据哈希：\`${digest}\`。\n`, 'utf8');
console.log(JSON.stringify({ root, accepted: accepted.length, pending: pending.length, hash: digest }));
