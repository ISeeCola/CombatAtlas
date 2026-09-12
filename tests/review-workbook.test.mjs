import assert from 'node:assert/strict';
import test from 'node:test';
import { collectReviewAdvisories, mergeAutomaticAdvice, validateReview } from '../scripts/review-workbook.mjs';

const base = { '候选 ID': 'cand-1', sourceId: '', '审核状态': '待复核', '中文标题': '测试资料', '规范 URL': 'https://example.com/source', '预筛总分': '20', '正式总分': '60', '审核备注': '', 评论: '' };

test('technical validation rejects duplicate candidate ids', () => {
  assert.match(validateReview([base, { ...base }]).join('\n'), /候选 ID 重复/);
});

test('editorial advice does not become a technical error', () => {
  const accepted = { ...base, '审核状态': '已接受' };
  assert.equal(validateReview([accepted]).length, 0);
  assert.equal(collectReviewAdvisories([accepted]).length, 1);
});

test('automatic advice replacement preserves human notes', () => {
  const merged = mergeAutomaticAdvice('人工意见\n【自动审核建议】旧建议', '【自动审核建议】新建议');
  assert.equal(merged, '人工意见\n【自动审核建议】新建议');
});
