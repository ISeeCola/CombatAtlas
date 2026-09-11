import generated from './generated-sources.json' with { type: 'json' };

export type Source = {
  id: string;
  title: string;
  originalTitle?: string;
  author: string;
  publisher: string;
  url: string;
  year: number;
  language: '中文' | '英文' | '日文';
  medium: '文章' | '视频' | '演讲' | '书籍';
  tier: '一手' | '精选二手';
  topics: string[];
  summary: string;
  takeaways: string[];
  readingTime: number;
  featured?: boolean;
  addedAt: string;
  curatorScore: number;
};

type SourceRecord = Source & { publicationStatus: 'published' | 'unpublished' };

function isSourceRecord(value: unknown): value is SourceRecord {
  if (!value || typeof value !== 'object') return false;
  const source = value as Record<string, unknown>;
  return (
    (source.publicationStatus === 'published' || source.publicationStatus === 'unpublished') &&
    typeof source.id === 'string' && source.id.length > 0 &&
    typeof source.title === 'string' && source.title.length > 0 &&
    typeof source.author === 'string' && typeof source.publisher === 'string' &&
    typeof source.url === 'string' && typeof source.year === 'number' &&
    ['中文', '英文', '日文'].includes(String(source.language)) &&
    ['文章', '视频', '演讲', '书籍'].includes(String(source.medium)) &&
    ['一手', '精选二手'].includes(String(source.tier)) &&
    Array.isArray(source.topics) && Array.isArray(source.takeaways) &&
    typeof source.summary === 'string' && typeof source.readingTime === 'number' &&
    typeof source.curatorScore === 'number' && /^\d{4}-\d{2}-\d{2}$/.test(String(source.addedAt))
  );
}

if (generated.schemaVersion !== 1 || !Array.isArray(generated.records)) {
  throw new Error('Invalid generated CombatAtlas source document');
}

const allRecords: SourceRecord[] = generated.records.map((record) => {
  if (!isSourceRecord(record)) throw new Error('Invalid CombatAtlas source record');
  return record;
});

const ids = new Set<string>();
for (const source of allRecords) {
  if (ids.has(source.id)) throw new Error(`Duplicate CombatAtlas source id: ${source.id}`);
  ids.add(source.id);
}

export const sources: Source[] = allRecords
  .filter((source) => source.publicationStatus === 'published')
  .map(({ publicationStatus: _publicationStatus, ...source }) => source);

export const topicOrder = ['全部', '角色3C', '技能', '手感', '打击感', '怪物设计', '怪物AI', 'Boss', '玩法规则', '遭遇战', '基础理论'];
