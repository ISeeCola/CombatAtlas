# CombatAtlas

面向战斗策划的长期知识库：<https://iseecola.github.io/CombatAtlas/>

## 本地双工作簿

- `source/combat_atlas_main.xlsm`：网页文章唯一人工真源，仅本地保存。
- `source/combat_atlas_review.xlsm`：候选、证据与审核真源，仅本地保存。

工作簿被 `.gitignore` 精确排除。公开仓库仅保存代码、`app/generated-sources.json` 与 `app/source-manifest.json`；CI 不读取 Excel。

审核表中只有“已接受”、正式分不低于 70 且证据/去重/实验/归档完整的条目可以点击“批准入库”。该操作只更新本地主表与文档。主表的“发布到网页”才会备份、生成、校验、构建、提交、推送并验证 Pages。

审核表末列“评论”保存拒绝条目的人工修改意见；后续巡检必须先读取评论并修改原记录。GDC Vault 候选则优先寻找 GDC 官方 YouTube 或其他官方公开视频作为收录入口。

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
