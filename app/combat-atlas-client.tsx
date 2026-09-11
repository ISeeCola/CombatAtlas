'use client';

import { useEffect, useMemo, useRef, useState } from 'react';
import { useGSAP } from '@gsap/react';
import { gsap, ScrollTrigger } from 'gsap/all';
import {
  ArrowUpRight,
  BookOpenCheck,
  CheckCircle2,
  Download,
  HardDrive,
  Search,
  Star,
  Upload,
  X,
} from 'lucide-react';
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select';
import { sources, topicOrder } from './sources';
import {
  createBackup,
  parseAndMergeBackup,
  PERSONAL_STATE_KEY,
  readPersonalState,
  type PersonalState,
  type PersonalStateMap,
  writePersonalState,
} from '@/lib/personal-state';

type SortMode =
  | 'curator'
  | 'rating'
  | 'year-desc'
  | 'year-asc'
  | 'added-desc'
  | 'added-asc';
type ReadFilter = 'all' | 'unread' | 'read';
type StorageState = 'loading' | 'saved' | 'saving' | 'error';

const initialCollectionDate = '2026-09-09';
const emptyPersonalState: PersonalState = {
  rating: 0,
  isRead: false,
  updatedAt: '',
};

gsap.registerPlugin(useGSAP, ScrollTrigger);

function hongKongDate() {
  return new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Hong_Kong',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).format(new Date());
}

function hongKongDateLabel() {
  return new Intl.DateTimeFormat('zh-CN', {
    timeZone: 'Asia/Hong_Kong',
    month: '2-digit',
    day: '2-digit',
    weekday: 'short',
  }).format(new Date());
}

export default function CombatAtlasClient() {
  const shellRef = useRef<HTMLElement>(null);
  const ghostCountRef = useRef<HTMLDivElement>(null);
  const importInputRef = useRef<HTMLInputElement>(null);
  const personalRef = useRef<PersonalStateMap>({});
  const [query, setQuery] = useState('');
  const [topic, setTopic] = useState('全部');
  const [tier, setTier] = useState('全部来源');
  const [readFilter, setReadFilter] = useState<ReadFilter>('all');
  const [sortMode, setSortMode] = useState<SortMode>('curator');
  const [personal, setPersonal] = useState<PersonalStateMap>({});
  const [storageState, setStorageState] = useState<StorageState>('loading');
  const [backupMessage, setBackupMessage] = useState('');
  const [savingSourceIds, setSavingSourceIds] = useState<Set<string>>(
    () => new Set(),
  );

  useEffect(() => {
    queueMicrotask(() => {
      try {
        const stored = readPersonalState();
        personalRef.current = stored;
        setPersonal(stored);
        setStorageState('saved');
      } catch {
        setStorageState('error');
      }
    });

    const handleStorage = (event: StorageEvent) => {
      if (event.key !== PERSONAL_STATE_KEY) return;
      try {
        const stored = readPersonalState();
        personalRef.current = stored;
        setPersonal(stored);
        setStorageState('saved');
      } catch {
        setStorageState('error');
      }
    };
    window.addEventListener('storage', handleStorage);
    return () => {
      window.removeEventListener('storage', handleStorage);
    };
  }, []);

  const updatePersonal = (
    sourceId: string,
    patch: Partial<PersonalState>,
  ) => {
    const previous = personalRef.current[sourceId] ?? emptyPersonalState;
    const next = {
      ...previous,
      ...patch,
      updatedAt: new Date().toISOString(),
    };
    const nextPersonal = { ...personalRef.current, [sourceId]: next };
    personalRef.current = nextPersonal;
    setPersonal(nextPersonal);
    setSavingSourceIds((current) => new Set(current).add(sourceId));
    setStorageState('saving');
    try {
      writePersonalState(nextPersonal);
      setStorageState('saved');
    } catch {
      personalRef.current = { ...personalRef.current, [sourceId]: previous };
      setPersonal(personalRef.current);
      setStorageState('error');
    }
    window.setTimeout(() => {
      setSavingSourceIds((current) => {
        const nextIds = new Set(current);
        nextIds.delete(sourceId);
        return nextIds;
      });
    }, 220);
  };

  const exportBackup = () => {
    const backup = createBackup(personalRef.current);
    const blob = new Blob([JSON.stringify(backup, null, 2)], {
      type: 'application/json',
    });
    const href = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = href;
    link.download = `combat-atlas-backup-${hongKongDate()}.json`;
    link.click();
    URL.revokeObjectURL(href);
    setBackupMessage('备份已导出');
  };

  const importBackup = async (file: File) => {
    try {
      const merged = parseAndMergeBackup(await file.text(), personalRef.current);
      writePersonalState(merged);
      personalRef.current = merged;
      setPersonal(merged);
      setStorageState('saved');
      setBackupMessage('备份已合并');
    } catch {
      setBackupMessage('备份无效，未修改数据');
    } finally {
      if (importInputRef.current) importInputRef.current.value = '';
    }
  };

  const readCount = useMemo(
    () => sources.filter((source) => personal[source.id]?.isRead).length,
    [personal],
  );
  const unreadCount = sources.length - readCount;
  const todayAdded = sources.filter(
    (item) => item.addedAt === hongKongDate(),
  ).length;
  const readProgress = sources.length ? (readCount / sources.length) * 100 : 0;

  const filtered = useMemo(() => {
    const needle = query.trim().toLowerCase();
    return sources
      .filter((source) => {
        const state = personal[source.id] ?? emptyPersonalState;
        const matchesQuery =
          !needle ||
          [
            source.title,
            source.originalTitle,
            source.author,
            source.publisher,
            source.summary,
            ...source.topics,
            ...source.takeaways,
          ]
            .filter(Boolean)
            .join(' ')
            .toLowerCase()
            .includes(needle);
        const matchesTopic = topic === '全部' || source.topics.includes(topic);
        const matchesTier = tier === '全部来源' || source.tier === tier;
        const matchesRead =
          readFilter === 'all' ||
          (readFilter === 'read' ? state.isRead : !state.isRead);
        return matchesQuery && matchesTopic && matchesTier && matchesRead;
      })
      .sort((a, b) => {
        if (sortMode === 'rating') {
          return (
            (personal[b.id]?.rating ?? 0) - (personal[a.id]?.rating ?? 0) ||
            b.curatorScore - a.curatorScore ||
            b.year - a.year
          );
        }
        if (sortMode === 'year-desc') {
          return b.year - a.year || b.curatorScore - a.curatorScore;
        }
        if (sortMode === 'year-asc') {
          return a.year - b.year || b.curatorScore - a.curatorScore;
        }
        if (sortMode === 'added-desc' || sortMode === 'added-asc') {
          const direction = sortMode === 'added-desc' ? -1 : 1;
          return (
            a.addedAt.localeCompare(b.addedAt) * direction ||
            b.curatorScore - a.curatorScore ||
            b.year - a.year ||
            a.id.localeCompare(b.id)
          );
        }
        return (
          b.curatorScore - a.curatorScore ||
          Number(Boolean(b.featured)) - Number(Boolean(a.featured)) ||
          b.year - a.year
        );
      });
  }, [personal, query, readFilter, sortMode, tier, topic]);

  const resetFilters = () => {
    setQuery('');
    setTopic('全部');
    setTier('全部来源');
    setReadFilter('all');
  };

  const showUnread = () => {
    setReadFilter('unread');
    document.getElementById('archive')?.scrollIntoView({
      behavior: window.matchMedia('(prefers-reduced-motion: reduce)').matches
        ? 'auto'
        : 'smooth',
    });
  };

  useGSAP(
    () => {
      const reduced = window.matchMedia(
        '(prefers-reduced-motion: reduce)',
      ).matches;
      if (reduced) {
        gsap.set('[data-motion]', { clearProps: 'all', opacity: 1 });
        return;
      }

      const counter = { value: 0 };
      if (ghostCountRef.current) ghostCountRef.current.textContent = '000';

      gsap
        .timeline({ defaults: { ease: 'power3.out' } })
        .from('.vinyl-nav > *', {
          opacity: 0,
          y: -16,
          duration: 0.65,
          stagger: 0.07,
        })
        .from(
          '.hero-copy > *',
          { opacity: 0, y: 32, duration: 0.75, stagger: 0.08 },
          '-=.32',
        )
        .from('.vinyl-disc', { opacity: 0, duration: 1.05 }, '-=.82')
        .to(
          counter,
          {
            value: sources.length,
            duration: 1.3,
            snap: { value: 1 },
            onUpdate: () => {
              if (ghostCountRef.current) {
                ghostCountRef.current.textContent = String(
                  Math.round(counter.value),
                ).padStart(3, '0');
              }
            },
          },
          0.08,
        )
        .from('.glitch-divider', { opacity: 0, duration: 0.65 }, '-=.42');

      gsap
        .timeline({
          scrollTrigger: {
            trigger: '.vinyl-home',
            start: 'top top',
            end: 'bottom top',
            scrub: 0.65,
          },
        })
        .to('.vinyl-stage', { scale: 0.88, opacity: 0.22, yPercent: 8 }, 0)
        .to('.hero-copy', { opacity: 0.16, y: -34 }, 0)
        .to('.ghost-count', { yPercent: -16, opacity: 0.02 }, 0);

      gsap.fromTo(
        '.archive-title-part, .archive-title-vinyl',
        { opacity: 0.1, y: 34 },
        {
          opacity: 1,
          y: 0,
          stagger: 0.08,
          ease: 'none',
          scrollTrigger: {
            trigger: '.archive-heading',
            start: 'top 92%',
            end: 'top 58%',
            scrub: 0.7,
          },
        },
      );
    },
    { scope: shellRef },
  );

  useGSAP(
    () => {
      if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
      gsap.utils.toArray<HTMLElement>('.source-card').forEach((card) => {
        const content = card.querySelectorAll<HTMLElement>(
          '.source-card-head, .source-title-block, .source-summary, .source-takeaways, .source-tags, .source-card-footer',
        );
        const line = card.querySelector<HTMLElement>('.card-entry-line');

        ScrollTrigger.create({
          trigger: card,
          start: 'top 94%',
          once: true,
          onEnter: () => {
            if (line) {
              gsap.fromTo(
                line,
                { scaleX: 0 },
                { scaleX: 1, duration: 0.48, ease: 'power2.out' },
              );
            }
            gsap.fromTo(
              content,
              { y: 8 },
              {
                y: 0,
                duration: 0.42,
                stagger: 0.025,
                ease: 'power2.out',
              },
            );
          },
        });
      });
    },
    {
      scope: shellRef,
      dependencies: [filtered.length, query, topic, tier, readFilter, sortMode],
      revertOnUpdate: true,
    },
  );

  useGSAP(
    () => {
      if (window.matchMedia('(prefers-reduced-motion: reduce)').matches) return;
      gsap.fromTo(
        '.live-read-count',
        { y: 4, color: '#d06b54' },
        {
          y: 0,
          color: '',
          duration: 0.28,
          stagger: 0.025,
          ease: 'power2.out',
          clearProps: 'transform,color',
        },
      );
    },
    { scope: shellRef, dependencies: [readCount], revertOnUpdate: true },
  );

  return (
    <main
      ref={shellRef}
      className="atlas-shell overflow-x-hidden w-full max-w-full"
    >
      <a className="skip-link" href="#archive">
        跳到资料库
      </a>

      <section className="vinyl-home" id="top" aria-labelledby="home-title">
        <header className="vinyl-nav">
          <a className="atlas-brand" href="#top" aria-label="返回主页顶部">
            <strong>COMBAT / ATLAS</strong>
            <span>战斗设计知识库</span>
          </a>
          <div
            className="now-playing"
            aria-label={`阅读进度 ${readCount}/${sources.length}`}
            aria-live="polite"
          >
            <strong>NOW PLAYING</strong>
            <span>TRACK {String(sources.length).padStart(3, '0')}</span>
            <i className="play-progress" aria-hidden="true">
              <b style={{ width: `${readProgress}%` }} />
            </i>
            <span className="live-read-count">
              {String(readCount).padStart(2, '0')} /{' '}
              {String(sources.length).padStart(2, '0')}
            </span>
          </div>
          <div className="vinyl-nav-meta">
            <span className="date-chip">{hongKongDateLabel()}</span>
            <button
              className="reading-chip"
              type="button"
              onClick={showUnread}
              disabled={storageState === 'loading'}
            >
              <span>在读</span>
              <strong>
                {storageState === 'loading' ? '--' : unreadCount}/{sources.length}
              </strong>
            </button>
          </div>
        </header>

        <div className="home-grid">
          <div className="hero-pressing-marks" aria-hidden="true">
            <span className="pressing-side">SIDE A · 33⅓ RPM</span>
            <span className="pressing-matrix">BDA–C2 / HK–033 / MASTER 01</span>
            <span className="pressing-halftone" />
          </div>
          <div className="hero-copy" data-motion>
            <p className="hero-kicker">KNOWLEDGE FOR ACTION DESIGN</p>
            <h1 id="home-title">
              战斗知识
              <span>正在播放</span>
            </h1>
            <p className="hero-lede">
              从角色控制、怪物行为到局内规则，保存能回到原始证据的设计资料。
            </p>
            <div className="hero-actions">
              <a className="primary-cta" href="#archive">
                进入资料库 <ArrowUpRight size={17} />
              </a>
              <span>
                首批条目{' '}
                {
                  sources.filter(
                    (item) => item.addedAt === initialCollectionDate,
                  ).length
                }
              </span>
            </div>
          </div>

          <div ref={ghostCountRef} className="ghost-count" aria-hidden="true">
            {String(sources.length).padStart(3, '0')}
          </div>
          <div className="vinyl-stage" aria-hidden="true" data-motion>
            <div className="vinyl-disc">
              <div className="vinyl-color-label">
                <div className="vinyl-label">
                  <strong>CA</strong>
                  <span>DESIGN ARCHIVE</span>
                  <i />
                </div>
              </div>
            </div>
          </div>
        </div>

        <div
          className="glitch-divider"
          aria-label={`已读 ${readCount} 条，共 ${sources.length} 条，今日新增 ${todayAdded} 条`}
          data-motion
        >
          <div className="tag-window" aria-hidden="true">
            <div className="tag-track">
              {[0, 1].map((copy) => (
                <div className="tag-set" key={copy}>
                  <span>SIDE A</span>
                  <i />
                  <span>HIT CONFIRM</span>
                  <i />
                  <span>33⅓ RPM</span>
                  <i />
                  <span>READ THE SIGNAL</span>
                  <i />
                  <span>FRAME DATA</span>
                  <i />
                  <span>MASTER CUT</span>
                  <i />
                  <span>RETURN TO SOURCE</span>
                  <i />
                </div>
              ))}
            </div>
          </div>
          <span className="divider-line" aria-hidden="true" />
          <div className="divider-metrics">
            <span>
              <em className="metric-full">已读</em>
              <em className="metric-short">读</em>
              <b className="live-read-count">
                {String(readCount).padStart(2, '0')}
              </b>
            </span>
            <span>
              <em className="metric-full">总条目</em>
              <em className="metric-short">总</em>
              <b>{String(sources.length).padStart(2, '0')}</b>
            </span>
            <span className="today-metric">
              <em className="metric-full">今日新增</em>
              <em className="metric-short">新</em>
              <b>{String(todayAdded).padStart(2, '0')}</b>
            </span>
          </div>
        </div>
      </section>

      <section
        className="archive-page"
        id="archive"
        aria-labelledby="archive-title"
      >
        <aside className="archive-edge-rail" aria-hidden="true">
          <span>SIDE B</span>
          <i />
          <b>PRESSING ARCHIVE · BDA–C2–033</b>
        </aside>
        <span
          className="archive-crop-mark archive-crop-mark-top"
          aria-hidden="true"
        />
        <span
          className="archive-crop-mark archive-crop-mark-bottom"
          aria-hidden="true"
        />
        <span className="archive-halftone" aria-hidden="true" />
        <div className="archive-inner">
          <header className="archive-heading">
            <div>
              <p>TRACK LIBRARY / {String(sources.length).padStart(2, '0')}</p>
              <h2 id="archive-title">
                <span className="archive-title-part">资料</span>
                <i className="archive-title-vinyl" aria-hidden="true">
                  <b />
                </i>
                <span className="archive-title-part is-accent">库</span>
              </h2>
            </div>
            <div className="archive-status">
              <span className="live-read-count">已读 {readCount}</span>
              <span>未读 {unreadCount}</span>
              <StorageIndicator state={storageState} />
              <div className="backup-controls">
                <button type="button" onClick={exportBackup}>
                  <Download size={14} /> 导出备份
                </button>
                <button type="button" onClick={() => importInputRef.current?.click()}>
                  <Upload size={14} /> 导入备份
                </button>
                <input
                  ref={importInputRef}
                  type="file"
                  accept="application/json,.json"
                  onChange={(event) => {
                    const file = event.target.files?.[0];
                    if (file) void importBackup(file);
                  }}
                  aria-label="选择 CombatAtlas 备份文件"
                />
              </div>
              <span className="backup-message" aria-live="polite">
                {backupMessage}
              </span>
            </div>
          </header>

          <div className="archive-controls">
            <label className="search-control">
              <Search size={18} />
              <input
                value={query}
                onChange={(event) => setQuery(event.target.value)}
                placeholder="搜索取消窗口、怪物 AI、资源循环……"
                aria-label="搜索知识库"
              />
              {query && (
                <button
                  type="button"
                  onClick={() => setQuery('')}
                  aria-label="清空搜索"
                >
                  <X size={17} />
                </button>
              )}
            </label>

            <div
              className="filter-deck"
              aria-label="资料筛选"
              data-pressing={`CATALOGUE / ${String(filtered.length).padStart(2, '0')} CUTS`}
            >
              <FilterGroup
                label="主题"
                values={topicOrder}
                current={topic}
                onChange={setTopic}
              />
              <FilterGroup
                label="来源"
                values={['全部来源', '一手', '精选二手']}
                current={tier}
                onChange={setTier}
              />
              <FilterGroup
                label="阅读"
                values={['all', 'unread', 'read']}
                labels={['全部', '未读', '已读']}
                current={readFilter}
                onChange={(value) => setReadFilter(value as ReadFilter)}
              />
              <div className="sort-control">
                <span>排序</span>
                <Select
                  value={sortMode}
                  onValueChange={(value) => setSortMode(value as SortMode)}
                >
                  <SelectTrigger className="sort-trigger">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent align="end" className="sort-menu">
                    <SelectItem value="curator">策展价值</SelectItem>
                    <SelectItem value="rating">个人评星</SelectItem>
                    <SelectItem value="year-desc">年份：新到旧</SelectItem>
                    <SelectItem value="year-asc">年份：旧到新</SelectItem>
                    <SelectItem value="added-desc">收录：新到旧</SelectItem>
                    <SelectItem value="added-asc">收录：旧到新</SelectItem>
                  </SelectContent>
                </Select>
              </div>
            </div>
          </div>

          <div className="result-heading">
            <p>
              {topic === '全部' ? '全部资料' : topic}
              <span className="result-count" key={filtered.length}>
                {String(filtered.length).padStart(2, '0')} TRACKS
              </span>
            </p>
            <button type="button" onClick={resetFilters}>
              重置筛选
            </button>
          </div>

          {filtered.length ? (
            <div className="source-grid grid-flow-dense">
              {filtered.map((source, index) => (
                <SourceCard
                  key={source.id}
                  source={source}
                  index={index}
                  state={personal[source.id] ?? emptyPersonalState}
                  isSaving={savingSourceIds.has(source.id)}
                  onChange={(patch) => updatePersonal(source.id, patch)}
                />
              ))}
            </div>
          ) : (
            <div className="empty-state">
              <span>NO MATCHING TRACK</span>
              <h3>没有匹配的资料</h3>
              <p>换一个主题或清除搜索词后再试。</p>
              <button type="button" onClick={resetFilters}>
                重置筛选
              </button>
            </div>
          )}

          <footer className="archive-outro">
            <div className="outro-vinyl" aria-hidden="true">
              <i />
            </div>
            <div className="outro-copy">
              <span className="outro-side">SIDE B · RUNOUT</span>
              <p>END OF CURRENT PRESSING</p>
              <strong className="live-read-count">
                {String(readCount).padStart(2, '0')} /{' '}
                {String(sources.length).padStart(2, '0')} TRACKS PLAYED
              </strong>
            </div>
            <span className="outro-runout" aria-hidden="true">
              <i />
            </span>
            <time dateTime={hongKongDate()}>{hongKongDateLabel()} 更新</time>
            <span className="outro-matrix" aria-hidden="true">
              BDA–C2–033 · MASTER 01
            </span>
            <a href="#top">
              回到唱片 <ArrowUpRight size={17} />
            </a>
          </footer>
        </div>
      </section>
    </main>
  );
}

function FilterGroup({
  label,
  values,
  labels,
  current,
  onChange,
}: {
  label: string;
  values: string[];
  labels?: string[];
  current: string;
  onChange: (value: string) => void;
}) {
  return (
    <div className="filter-group">
      <span>{label}</span>
      <div>
        {values.map((value, index) => (
          <button
            type="button"
            key={value}
            onClick={() => onChange(value)}
            className={current === value ? 'is-active' : ''}
            aria-pressed={current === value}
          >
            {labels?.[index] ?? value}
          </button>
        ))}
      </div>
    </div>
  );
}

function StorageIndicator({
  state,
  compact = false,
}: {
  state: StorageState;
  compact?: boolean;
}) {
  if (state === 'error') {
    return (
      <span className={`sync-state sync-error ${compact ? 'is-compact' : ''}`}>
        <HardDrive size={14} />
        本机存储不可用
      </span>
    );
  }
  return (
    <span className={`sync-state ${compact ? 'is-compact' : ''}`}>
      <HardDrive size={14} />
      {state === 'saved' ? '本机已保存' : state === 'saving' ? '保存中' : '读取中'}
    </span>
  );
}

function SourceCard({
  source,
  index,
  state,
  isSaving,
  onChange,
}: {
  source: (typeof sources)[number];
  index: number;
  state: PersonalState;
  isSaving: boolean;
  onChange: (patch: Partial<PersonalState>) => void;
}) {
  const cardRef = useRef<HTMLElement>(null);

  const toggleRead = () => {
    const sweep =
      cardRef.current?.querySelector<HTMLElement>('.pressing-sweep');
    const reduced = window.matchMedia(
      '(prefers-reduced-motion: reduce)',
    ).matches;

    if (sweep && !reduced) {
      const markingRead = !state.isRead;
      gsap.killTweensOf(sweep);
      gsap
        .timeline()
        .set(sweep, {
          xPercent: markingRead ? -110 : 110,
          opacity: 0,
        })
        .to(sweep, {
          xPercent: 0,
          opacity: 0.14,
          duration: 0.17,
          ease: 'power2.out',
        })
        .to(sweep, {
          xPercent: markingRead ? 110 : -110,
          opacity: 0,
          duration: 0.24,
          ease: 'power2.in',
        });
    }

    onChange({ isRead: !state.isRead });
  };

  return (
    <article
      ref={cardRef}
      className={`source-card ${state.isRead ? 'is-read' : ''}`}
      aria-busy={isSaving}
      data-motion
    >
      <span className="card-entry-line" aria-hidden="true" />
      <span className="card-sweep" aria-hidden="true" />
      <span className="card-grooves" aria-hidden="true" />
      <span className="card-runout" aria-hidden="true" />
      <span className="pressing-sweep" aria-hidden="true" />
      <header className="source-card-head">
        <span className="track-number">
          {String(index + 1).padStart(2, '0')}
        </span>
        <div className="source-badges">
          <span
            className={
              source.tier === '一手' ? 'source-tier is-primary' : 'source-tier'
            }
          >
            {source.tier}
          </span>
          <span>{source.medium}</span>
          <span>{source.year}</span>
          <span title={`收录于 ${source.addedAt}`}>
            收录 {source.addedAt.slice(5).replace('-', '/')}
          </span>
          {source.featured && <strong>精选</strong>}
        </div>
      </header>

      <div className="source-title-block">
        <h3 title={source.title}>{source.title}</h3>
        {source.originalTitle && <p>{source.originalTitle}</p>}
        <span>
          {source.author} · {source.publisher}
        </span>
      </div>

      <p className="source-summary">{source.summary}</p>

      <div className="source-takeaways">
        <span>核心结论</span>
        <ul>
          {source.takeaways.map((item) => (
            <li key={item}>{item}</li>
          ))}
        </ul>
      </div>

      <div className="source-tags">
        {source.topics.slice(0, 5).map((item) => (
          <span key={item}>{item}</span>
        ))}
      </div>

      <footer className="source-card-footer">
        <div className="value-block">
          <span>策展价值</span>
          <strong>{source.curatorScore}/5</strong>
        </div>
        <div
          className="rating-control"
          aria-label={`个人评星：${state.rating} 星`}
        >
          {[1, 2, 3, 4, 5].map((star) => (
            <button
              type="button"
              key={star}
              disabled={isSaving}
              onClick={() =>
                onChange({ rating: state.rating === star ? 0 : star })
              }
              className={star <= state.rating ? 'is-active' : ''}
              aria-pressed={state.rating === star}
              aria-label={
                state.rating === star ? '清除评星' : `评为 ${star} 星`
              }
              title={`${star} 星`}
            >
              <Star
                size={17}
                fill={star <= state.rating ? 'currentColor' : 'none'}
              />
            </button>
          ))}
        </div>
        <button
          type="button"
          disabled={isSaving}
          onClick={toggleRead}
          className={`read-control ${state.isRead ? 'is-active' : ''}`}
          aria-pressed={state.isRead}
        >
          <i>
            {state.isRead ? (
              <CheckCircle2 size={16} />
            ) : (
              <BookOpenCheck size={16} />
            )}
          </i>
          <span>
            <small>{state.isRead ? 'PLAYED' : 'QUEUE'}</small>
            <b>{state.isRead ? '已读' : '标记已读'}</b>
          </span>
        </button>
        <a
          className="source-link"
          href={source.url}
          target="_blank"
          rel="noreferrer"
        >
          <span>约 {source.readingTime} 分钟</span>
          查看原文 <ArrowUpRight size={16} />
        </a>
      </footer>
      <span className="card-matrix" aria-hidden="true">
        BDA–{String(index + 1).padStart(2, '0')} · {source.id.toUpperCase()}
      </span>
    </article>
  );
}
