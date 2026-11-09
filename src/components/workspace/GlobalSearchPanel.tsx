import React, { useCallback, useRef, useState } from 'react';
import {
  Search,
  FileText,
  X,
  CaseSensitive,
  Regex,
  Loader2,
  ChevronDown,
  ChevronRight,
} from 'lucide-react';
import { SearchResult } from '../../types';
import { fileSystemService } from '../../services/fileSystemService';
import { cn } from '../../lib/utils';

interface Props {
  workspacePath: string;
  onOpenFile: (filePath: string) => void;
  onClose: () => void;
}

export const GlobalSearchPanel: React.FC<Props> = ({ workspacePath, onOpenFile, onClose }) => {
  const [query, setQuery] = useState('');
  const [results, setResults] = useState<SearchResult[]>([]);
  const [loading, setLoading] = useState(false);
  const [hasSearched, setHasSearched] = useState(false);
  const [error, setError] = useState('');
  const [caseSensitive, setCaseSensitive] = useState(false);
  const [useRegex, setUseRegex] = useState(false);
  const [includeGlob, setIncludeGlob] = useState('');
  const [showOptions, setShowOptions] = useState(false);

  const abortRef = useRef(0);

  const runSearch = useCallback(
    async (q: string, cs: boolean, rx: boolean, glob: string) => {
      if (!q.trim() || !workspacePath) return;
      const token = ++abortRef.current;
      setLoading(true);
      setError('');
      setHasSearched(true);

      try {
        const res = await fileSystemService.searchWorkspace(workspacePath, q.trim(), true, {
          caseSensitive: cs,
          useRegex: rx,
          includeGlob: glob.trim() || '*',
        });
        // Игнорируем устаревшие ответы при быстром наборе
        if (token !== abortRef.current) return;
        setResults(
          res
            .filter((r) => !r.isDirectory)
            .sort((a, b) => (b.totalMatches ?? 0) - (a.totalMatches ?? 0))
        );
      } catch (err) {
        if (token !== abortRef.current) return;
        setError((err as Error).message);
        setResults([]);
      } finally {
        if (token === abortRef.current) setLoading(false);
      }
    },
    [workspacePath]
  );

  const handleSubmit = (e?: React.FormEvent) => {
    e?.preventDefault();
    void runSearch(query, caseSensitive, useRegex, includeGlob);
  };

  const totalMatches = results.reduce((sum, r) => sum + (r.totalMatches ?? 1), 0);

  return (
    <div className="flex h-full flex-col bg-ws-panel text-ws select-none">
      {/* Заголовок */}
      <div className="flex shrink-0 items-center justify-between border-b border-ws px-3 py-2">
        <div className="flex items-center gap-2 text-[11px] font-semibold tracking-wider text-ws-muted uppercase">
          <Search size={13} className="text-ws-accent" />
          Поиск по файлам
        </div>
        <button
          onClick={onClose}
          title="Закрыть поиск"
          className="rounded p-0.5 text-ws-faint hover:text-ws"
        >
          <X size={13} />
        </button>
      </div>

      {/* Форма поиска */}
      <form onSubmit={handleSubmit} className="shrink-0 space-y-2 border-b border-ws p-3">
        <div className="relative">
          <Search size={13} className="absolute top-1/2 left-2.5 -translate-y-1/2 text-ws-faint" />
          <input
            type="text"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Искать текст в файлах…"
            spellCheck={false}
            className="w-full rounded-lg border border-ws bg-ws-elevated py-1.5 pr-2 pl-8 text-xs text-ws placeholder:text-ws-faint focus:border-ws-accent focus:outline-none"
          />
        </div>

        <div className="flex items-center gap-1">
          <button
            type="button"
            onClick={() => {
              const next = !caseSensitive;
              setCaseSensitive(next);
              void runSearch(query, next, useRegex, includeGlob);
            }}
            title="Учитывать регистр"
            className={cn(
              'rounded border p-1 transition-colors',
              caseSensitive
                ? 'border-ws-accent bg-ws-accent-soft text-ws-accent'
                : 'border-ws text-ws-faint hover:text-ws'
            )}
          >
            <CaseSensitive size={13} />
          </button>

          <button
            type="button"
            onClick={() => {
              const next = !useRegex;
              setUseRegex(next);
              void runSearch(query, caseSensitive, next, includeGlob);
            }}
            title="Регулярное выражение"
            className={cn(
              'rounded border p-1 transition-colors',
              useRegex
                ? 'border-ws-accent bg-ws-accent-soft text-ws-accent'
                : 'border-ws text-ws-faint hover:text-ws'
            )}
          >
            <Regex size={13} />
          </button>

          <button
            type="button"
            onClick={() => setShowOptions((v) => !v)}
            className="ml-auto flex items-center gap-1 text-[10px] text-ws-faint hover:text-ws"
          >
            Фильтр
            {showOptions ? <ChevronDown size={11} /> : <ChevronRight size={11} />}
          </button>
        </div>

        {showOptions && (
          <input
            type="text"
            value={includeGlob}
            onChange={(e) => setIncludeGlob(e.target.value)}
            placeholder="Маска файлов, напр. *.ts или src/*"
            spellCheck={false}
            className="w-full rounded border border-ws bg-ws-elevated px-2 py-1 font-mono text-[10px] text-ws placeholder:text-ws-faint focus:border-ws-accent focus:outline-none"
          />
        )}

        <div className="flex items-center justify-between text-[10px] text-ws-faint">
          <span>Enter — искать</span>
          {loading && (
            <span className="flex items-center gap-1 text-ws-accent">
              <Loader2 size={11} className="animate-spin" />
              Поиск…
            </span>
          )}
        </div>
      </form>

      {/* Результаты */}
      <div className="min-h-0 flex-1 overflow-y-auto p-2">
        {error && (
          <div className="mb-2 rounded-lg border border-red-500/40 bg-red-500/10 px-2.5 py-2 text-[11px] text-ws-danger">
            {error}
          </div>
        )}

        {hasSearched && !loading && !error && results.length === 0 && (
          <div className="p-6 text-center text-[11px] text-ws-faint">
            Ничего не найдено по запросу «{query}»
          </div>
        )}

        {!hasSearched && (
          <p className="p-4 text-center text-[11px] leading-relaxed text-ws-faint">
            Введите текст для поиска по содержимому всех файлов рабочей папки.
            <br />
            Служебные папки (node_modules, .git) исключаются автоматически.
          </p>
        )}

        {results.map((res) => {
          const relativePath = res.path.replace(workspacePath, '').replace(/^[\\/]/, '');
          return (
            <div
              key={res.path}
              onClick={() => onOpenFile(res.path)}
              className="mb-1.5 cursor-pointer rounded-lg border border-ws bg-ws-elevated p-2.5 transition-colors hover:border-ws-accent hover:bg-ws-hover"
            >
              <div className="mb-1 flex items-center justify-between gap-2 text-xs font-medium text-ws">
                <div className="flex min-w-0 items-center gap-1.5">
                  <FileText size={12} className="shrink-0 text-ws-accent" />
                  <span className="truncate">{res.name}</span>
                </div>
                {res.line !== undefined && (
                  <span className="shrink-0 rounded border border-ws bg-ws-panel px-1.5 py-0.5 font-mono text-[10px] text-ws-accent">
                    стр. {res.line}
                  </span>
                )}
              </div>

              <div className="mb-1 truncate text-[10px] text-ws-faint">{relativePath}</div>

              {res.matchSnippet && (
                <div className="truncate rounded border border-ws bg-ws-editor-alt px-1.5 py-1 font-mono text-[10px] text-ws-success">
                  {res.matchSnippet}
                </div>
              )}

              {(res.totalMatches ?? 0) > 1 && (
                <div className="mt-1 text-[10px] text-ws-faint">
                  Совпадений в файле: {res.totalMatches}
                </div>
              )}
            </div>
          );
        })}
      </div>

      {/* Подвал */}
      <div className="shrink-0 border-t border-ws bg-ws-elevated px-3 py-1.5 text-[10px] text-ws-faint">
        Файлов: {results.length}
        {totalMatches > 0 && ` · совпадений: ${totalMatches}`}
      </div>
    </div>
  );
};
