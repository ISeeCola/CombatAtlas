# CombatAtlas

面向战斗策划的长期知识库：<https://iseecola.github.io/CombatAtlas/>

## 本地双工作簿

- `source/combat_atlas_main.xlsm`：网页文章唯一人工真源，仅本地保存。
- `source/combat_atlas_review.xlsm`：候选、证据与审核真源，仅本地保存。

工作簿被 `.gitignore` 精确排除。公开仓库仅保存代码、`app/generated-sources.json` 与 `app/source-manifest.json`；CI 不读取 Excel。

审核自动化负责给出证据、评分、去重、实验和归档建议，但人工拥有最终决定权。维护者将条目标为“已接受”后，“批准入库”必须直接尝试写入；审核缺项会以 `【自动审核建议】` 保留在“审核备注”，不会阻止人工入库。只有缺标题/URL、非法 URL、重复 ID、非法状态、主表结构损坏或新增网页必需字段缺失等技术错误会阻止。

若已接受条目的 `sourceId` 已存在主表，批准入库会用 Review 的非空公开字段更新主表对应行；Review 空白字段保留主表现值，永久 ID 与发布状态不变。

状态为“待修改”的条目在“评论”列保存人工修改要求；每日巡检优先按当前评论版本重做原记录，且不占每日 3 条新增额度。“已拒绝”代表永久质量拒绝，自动化不得重试。GDC Vault 候选优先寻找 GDC 官方 YouTube 或其他官方公开视频作为收录入口。

永久 `sourceId` 是个人评星/已读的唯一关联键；不得修改或复用。下架请改“发布状态”，不要删除记录。

## 开发校验

```bash
npm ci
npm run sources:verify-generated
npm run review:validate
npm run lint
npx tsc --noEmit --incremental false
npm run build
```

`generated-sources.json` 和 manifest 为派生文件，不应手工编辑。不同设备的个人状态继续通过网页 JSON 导入/导出迁移。
