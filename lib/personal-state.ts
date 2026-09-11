export type PersonalState = {
  rating: number;
  isRead: boolean;
  updatedAt: string;
};

export type PersonalStateMap = Record<string, PersonalState>;

export const PERSONAL_STATE_KEY = 'combat-atlas:personal-state:v1';
export const PERSONAL_STATE_VERSION = 1;

type PersonalStateBackup = {
  app: 'CombatAtlas';
  version: typeof PERSONAL_STATE_VERSION;
  exportedAt: string;
  items: Array<PersonalState & { sourceId: string }>;
};

function isValidItem(value: unknown): value is PersonalState & { sourceId: string } {
  if (!value || typeof value !== 'object') return false;
  const item = value as Record<string, unknown>;
  return (
    typeof item.sourceId === 'string' &&
    item.sourceId.length > 0 &&
    item.sourceId.length <= 128 &&
    Number.isInteger(item.rating) &&
    Number(item.rating) >= 0 &&
    Number(item.rating) <= 5 &&
    typeof item.isRead === 'boolean' &&
    typeof item.updatedAt === 'string' &&
    !Number.isNaN(Date.parse(item.updatedAt))
  );
}

export function readPersonalState(): PersonalStateMap {
  const raw = window.localStorage.getItem(PERSONAL_STATE_KEY);
  if (!raw) return {};

  const parsed = JSON.parse(raw) as unknown;
  if (!parsed || typeof parsed !== 'object') throw new Error('invalid state');

  const result: PersonalStateMap = {};
  for (const [sourceId, value] of Object.entries(parsed)) {
    if (isValidItem({ sourceId, ...(value as object) })) {
      const item = value as PersonalState;
      result[sourceId] = item;
    }
  }
  return result;
}

export function writePersonalState(state: PersonalStateMap) {
  window.localStorage.setItem(PERSONAL_STATE_KEY, JSON.stringify(state));
}

export function createBackup(state: PersonalStateMap): PersonalStateBackup {
  return {
    app: 'CombatAtlas',
    version: PERSONAL_STATE_VERSION,
    exportedAt: new Date().toISOString(),
    items: Object.entries(state).map(([sourceId, item]) => ({
      sourceId,
      ...item,
    })),
  };
}

export function parseAndMergeBackup(
  raw: string,
  current: PersonalStateMap,
): PersonalStateMap {
  const parsed = JSON.parse(raw) as Partial<PersonalStateBackup>;
  if (
    parsed.app !== 'CombatAtlas' ||
    parsed.version !== PERSONAL_STATE_VERSION ||
    !Array.isArray(parsed.items) ||
    !parsed.items.every(isValidItem)
  ) {
    throw new Error('invalid backup');
  }

  const merged = { ...current };
  for (const item of parsed.items) {
    const existing = merged[item.sourceId];
    if (!existing || Date.parse(item.updatedAt) >= Date.parse(existing.updatedAt)) {
      merged[item.sourceId] = {
        rating: item.rating,
        isRead: item.isRead,
        updatedAt: item.updatedAt,
      };
    }
  }
  return merged;
}
