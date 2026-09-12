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

const LANGUAGES = new Set(['中文', '英文', '日文']);
const MEDIA = new Set(['文章', '视频', '演讲', '书籍']);
const TIERS = new Set(['一手', '精选二手']);

function isValidDate(value: unknown): value is string {
  if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  const [year, month, day] = value.split('-').map(Number);
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
}

function isNonEmptyStringList(value: unknown): value is string[] {
  return Array.isArray(value) && value.length > 0 && value.every((item) => typeof item === 'string' && item.trim().length > 0);
}

function isSourceRecord(value: unknown): value is SourceRecord {
  if (!value || typeof value !== 'object') return false;
  const source = value as Record<string, unknown>;
  return (
    (source.publicationStatus === 'published' || source.publicationStatus === 'unpublished') &&
    typeof source.id === 'string' && source.id.length > 0 &&
    typeof source.title === 'string' && source.title.length > 0 &&
    typeof source.author === 'string' && source.author.trim().length > 0 &&
    typeof source.publisher === 'string' && source.publisher.trim().length > 0 &&
    typeof source.url === 'string' && /^https?:\/\//i.test(source.url) &&
    typeof source.year === 'number' && Number.isInteger(source.year) && source.year >= 1970 && source.year <= 2100 &&
    LANGUAGES.has(String(source.language)) && MEDIA.has(String(source.medium)) && TIERS.has(String(source.tier)) &&
    isNonEmptyStringList(source.topics) && isNonEmptyStringList(source.takeaways) &&
    typeof source.summary === 'string' && source.summary.trim().length > 0 &&
    typeof source.readingTime === 'number' && Number.isInteger(source.readingTime) && source.readingTime > 0 &&
    typeof source.curatorScore === 'number' && Number.isInteger(source.curatorScore) && source.curatorScore >= 1 && source.curatorScore <= 5 &&
    (source.featured === undefined || typeof source.featured === 'boolean') &&
    (source.originalTitle === undefined || typeof source.originalTitle === 'string') &&
    isValidDate(source.addedAt)
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

const preferredTopics = ['角色3C', '技能', '手感', '打击感', '怪物设计', '怪物AI', 'Boss', '玩法规则', '遭遇战', '基础理论'];
const extraTopics = [...new Set(sources.flatMap((source) => source.topics))].filter((topic) => !preferredTopics.includes(topic));
export const topicOrder = ['全部', ...preferredTopics, ...extraTopics];
