import React, { useCallback, useEffect, useRef, useState } from 'react';
import { Search, File, X, CornerDownLeft, Clock } from 'lucide-react';
import { SearchResult } from '../../types';
import { fileSystemService } from '../../services/fileSystemService';
import { cn } from '../../lib/utils';
import { Spinner } from '../ui';

interface Props {
  isOpen: boolean;
  onClose: () => void;
  workspacePath: string;
  onOpenFile: (filePath: string) => void;
}

const RECENT_KEY = 'ws_recent_files';

export const QuickOpenModal: React.FC<Props> = ({
  isOpen,
  onClose,
  workspacePath,
  onOpenFile,
}) => {
  const [query, setQuery] = useState('');
  const [results, setResults] = useState<SearchResult[]>([]);
  const [selectedIndex, setSelectedIndex] = useState(0);
  const [loading, setLoading] = useState(false);
  const [recent, setRecent] = useState<string[]>([]);
  const inputRef = useRef<HTMLInputElement>(null);
  const debounceRef = useRef<number | undefined>(undefined);

  // Загрузка списка недавних файлов
  useEffect(() => {
    try {
      const raw = window.localStorage.getItem(RECENT_KEY);
      if (raw) setRecent(JSON.parse(raw) as string[]);
    } catch {
      // игнорируем
    }
  }, []);

  useEffect(() => {
    if (!isOpen) return;
    setQuery('');
    setSelectedIndex(0);
    window.setTimeout(() => inputRef.current?.focus(), 40);
    if (workspacePath) void searchFiles('');
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isOpen, workspacePath]);

  const searchFiles = useCallback(
    async (q: string) => {
      if (!workspacePath) return;
      try {
        setLoading(true);
        const res = await fileSystemService.searchWorkspace(workspacePath, q || '', false);
        setResults(res.filter((r) => !r.isDirectory).slice(0, 100));
        setSelectedIndex(0);
      } catch {
        setResults([]);
      } finally {
        setLoading(false);
      }
    },
    [workspacePath]
  );

  // Поиск с задержкой, чтобы не дёргать ФС на каждый символ
  const handleQueryChange = (value: string) => {
    setQuery(value);
    window.clearTimeout(debounceRef.current);
    debounceRef.current = window.setTimeout(() => void searchFiles(value), 130);
  };

  /** Оценка релевантности: точное совпадение имени > начало > вхождение. */
  const ranked = React.useMemo(() => {
    if (!query.trim()) {
      // Без запроса: сначала недавние, затем остальные
      const recentSet = new Set(recent);
      const recentResults = recent
        .map((p) => results.find((r) => r.path === p))
        .filter((r): r is SearchResult => Boolean(r));
      const rest = results.filter((r) => !recentSet.has(r.path));
      return [...recentResults, ...rest];
    }
    const q = query.toLowerCase();
    return [...results].sort((a, b) => {
      const aName = a.name.toLowerCase();
      const bName = b.name.toLowerCase();
      const score = (name: string) =>
        name === q ? 0 : name.startsWith(q) ? 1 : name.includes(q) ? 2 : 3;
      return score(aName) - score(bName) || aName.length - bName.length;
    });
  }, [results, query, recent]);

  const commitOpen = (filePath: string) => {
    // Запоминаем открытый файл в списке недавних
    try {
      const next = [filePath, ...recent.filter((p) => p !== filePath)].slice(0, 12);
      window.localStorage.setItem(RECENT_KEY, JSON.stringify(next));
      setRecent(next);
    } catch {
      // игнорируем
    }
    onOpenFile(filePath);
    onClose();
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      onClose();
    } else if (e.key === 'ArrowDown') {
      e.preventDefault();
      setSelectedIndex((i) => (ranked.length ? (i + 1) % ranked.length : 0));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setSelectedIndex((i) => (ranked.length ? (i - 1 + ranked.length) % ranked.length : 0));
    } else if (e.key === 'Enter') {
      e.preventDefault();
      const hit = ranked[selectedIndex];
      if (hit) commitOpen(hit.path);
    }
  };

  if (!isOpen) return null;

  return (
    <div
      className="fixed inset-0 z-50 flex items-start justify-center bg-black/50 pt-24 backdrop-blur-[2px]"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div className="animate-scale-in flex max-h-[60vh] w-full max-w-2xl flex-col overflow-hidden rounded-xl border border-ws bg-ws-surface shadow-ws">
        {/* Поле ввода */}
        <div className="flex shrink-0 items-center gap-3 border-b border-ws px-4 py-3">
          <Search size={17} className="shrink-0 text-ws-accent" />
          <input
            ref={inputRef}
            type="text"
            value={query}
            onChange={(e) => handleQueryChange(e.target.value)}
            onKeyDown={handleKeyDown}
            placeholder="Начните вводить имя файла…"
            spellCheck={false}
            className="w-full bg-transparent text-sm text-ws placeholder:text-ws-faint focus:outline-none"
          />
          {loading && <Spinner size={14} className="shrink-0 text-ws-accent" />}
          {query && (
            <button
              onClick={() => handleQueryChange('')}
              className="shrink-0 text-ws-faint hover:text-ws"
            >
              <X size={15} />
            </button>
          )}
        </div>

        {/* Результаты */}
        <div className="min-h-0 flex-1 overflow-y-auto p-2">
          {!workspacePath && (
            <div className="p-6 text-center text-xs text-ws-faint">
              Сначала выберите рабочую папку
            </div>
          )}

          {workspacePath && !loading && ranked.length === 0 && (
            <div className="p-6 text-center text-xs text-ws-faint">
              {query ? `Ничего не найдено по запросу «${query}»` : 'В рабочей папке нет файлов'}
            </div>
          )}

          {ranked.map((res, idx) => {
            const isSelected = idx === selectedIndex;
            const relativePath =
              res.path.replace(workspacePath, '').replace(/^[\\/]/, '') || res.name;
            const isRecent = !query.trim() && recent.includes(res.path);

            return (
              <div
                key={res.path}
                onClick={() => commitOpen(res.path)}
                onMouseEnter={() => setSelectedIndex(idx)}
                className={cn(
                  'flex cursor-pointer items-center justify-between rounded-lg px-3 py-2 text-xs transition-colors',
                  isSelected ? 'bg-ws-accent text-white' : 'text-ws-muted hover:bg-ws-hover'
                )}
              >
                <div className="flex min-w-0 items-center gap-2.5">
                  {isRecent ? (
                    <Clock size={14} className={isSelected ? 'text-white' : 'text-ws-faint'} />
                  ) : (
                    <File size={14} className={isSelected ? 'text-white' : 'text-ws-faint'} />
                  )}
                  <span className={cn('truncate font-medium', isSelected ? 'text-white' : 'text-ws')}>
                    {res.name}
                  </span>
                  <span
                    className={cn(
                      'truncate text-[11px]',
                      isSelected ? 'text-blue-100' : 'text-ws-faint'
                    )}
                  >
                    {relativePath}
                  </span>
                </div>

                {isSelected && (
                  <div className="flex shrink-0 items-center gap-1 text-[10px] opacity-85">
                    <span>Открыть</span>
                    <CornerDownLeft size={11} />
                  </div>
                )}
              </div>
            );
          })}
        </div>

        {/* Подвал */}
        <div className="flex shrink-0 items-center justify-between border-t border-ws bg-ws-panel px-4 py-2 text-[10px] text-ws-faint">
          <div className="flex items-center gap-3">
            <span>↑ ↓ навигация</span>
            <span>Enter открыть</span>
            <span>Esc закрыть</span>
          </div>
          <span>Найдено: {ranked.length}</span>
        </div>
      </div>
    </div>
  );
};
