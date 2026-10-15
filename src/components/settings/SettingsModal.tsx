import React, { useEffect, useMemo, useState } from 'react';
import {
  Settings as SettingsIcon,
  Sparkles,
  Eye,
  EyeOff,
  Check,
  AlertTriangle,
  ExternalLink,
  KeyRound,
  ShieldCheck,
  Cloud,
  Loader2,
  Trash2,
  RefreshCw,
  Sliders,
  FileKey,
  Download,
  Info,
  Monitor,
} from 'lucide-react';
import { AISettings, AIProvider } from '../../types';
import { PROVIDER_PRESETS, PROVIDER_ORDER, SYSTEM_PRESETS, keyVault } from '../../state/settings';
import { aiService } from '../../services/aiService';
import { cn } from '../../lib/utils';
import { Button, Field, Input, Modal, Select, Textarea } from '../ui';
import { IntegrationsSettings } from './IntegrationsSettings';
import { SystemSettings } from './SystemSettings';

interface Props {
  open: boolean;
  onClose: () => void;
  settings: AISettings;
  onChangeProvider: (provider: AIProvider) => void | Promise<void>;
  onUpdateSettings: (patch: Partial<AISettings>) => void;
  onSaveApiKey: (key: string) => void | Promise<void>;
  availableModels?: string[];
  onFetchModels?: (override?: Partial<AISettings>) => Promise<{ success: boolean; models: string[]; error?: string }>;
  modelsCache?: Record<string, string[]>;
  initialTab?: Tab;
  onNotify?: (message: string) => void;
}

type Tab = 'provider' | 'model' | 'prompt' | 'keys' | 'integrations' | 'system';

interface SecretsInfo {
  found: boolean;
  filePath: string;
  baseUrl?: string;
  textModel?: string;
  imageModel?: string;
  hasKey: boolean;
  hasOtherSecrets: boolean;
  error?: string;
}

export const SettingsModal: React.FC<Props> = ({
  open,
  onClose,
  settings,
  onChangeProvider,
  onUpdateSettings,
  onSaveApiKey,
  availableModels,
  onFetchModels,
  modelsCache,
  initialTab = 'provider',
  onNotify,
}) => {
  const [tab, setTab] = useState<Tab>(initialTab);
  const [keyVisible, setKeyVisible] = useState(false);
  const [keyDraft, setKeyDraft] = useState(settings.apiKey);
  const [testing, setTesting] = useState(false);
  const [testResult, setTestResult] = useState<{ ok: boolean; message: string } | null>(null);
  const [configuredProviders, setConfiguredProviders] = useState<AIProvider[]>([]);
  const [customModel, setCustomModel] = useState('');
  const [savedFlash, setSavedFlash] = useState(false);
  const [secrets, setSecrets] = useState<SecretsInfo | null>(null);
  const [importing, setImporting] = useState(false);
  const [importResult, setImportResult] = useState<{ ok: boolean; message: string } | null>(null);
  const [loadingModels, setLoadingModels] = useState(false);
  const [modelFetchResult, setModelFetchResult] = useState<{ ok: boolean; message: string; count?: number } | null>(null);
  const [modelSearch, setModelSearch] = useState('');

  const preset = PROVIDER_PRESETS[settings.provider];

  const providerModels = useMemo<string[]>(() => {
    const fromApi = modelsCache?.[settings.provider] || [];
    const fromPreset = preset.models || [];
    const fromProps = availableModels || [];
    const set = new Set<string>([...fromApi, ...fromProps, ...fromPreset]);
    if (settings.model) set.add(settings.model);
    return Array.from(set);
  }, [modelsCache, settings.provider, preset.models, availableModels, settings.model]);

  const filteredModels = useMemo<string[]>(() => {
    if (!modelSearch.trim()) return providerModels;
    const q = modelSearch.toLowerCase();
    return providerModels.filter((m: string) => m.toLowerCase().includes(q));
  }, [providerModels, modelSearch]);

  // При открытии синхронизируем черновик ключа и список сохранённых провайдеров
  useEffect(() => {
    if (!open) return;
    setKeyDraft(settings.apiKey);
    setTestResult(null);
    void keyVault.listConfigured().then(setConfiguredProviders);
    // Проверяем наличие локального файла секретов
    if (window.electronAPI?.describeSecrets) {
      void window.electronAPI.describeSecrets().then(setSecrets).catch(() => setSecrets(null));
    }
  }, [open, settings.apiKey, settings.provider]);

  /** Импорт ключа из secrets.local.ps1 — только по явному действию пользователя. */
  const handleImportSecrets = async () => {
    if (!window.electronAPI?.importSecretsKey) return;
    setImporting(true);
    setImportResult(null);
    try {
      const res = await window.electronAPI.importSecretsKey();
      if (!res.success || !res.apiKey) {
        setImportResult({ ok: false, message: res.error || 'Ключ не найден' });
        return;
      }

      // Переключаемся на провайдера, соответствующего импортированному endpoint
      const targetProvider: AIProvider =
        res.baseUrl?.includes('starimg') ? 'starimg' : 'custom';

      if (targetProvider !== settings.provider) {
        await onChangeProvider(targetProvider);
      }
      if (res.baseUrl) onUpdateSettings({ baseUrl: res.baseUrl });
      if (res.model) onUpdateSettings({ model: res.model });
      setKeyDraft(res.apiKey);
      await onSaveApiKey(res.apiKey);

      const list = await keyVault.listConfigured();
      setConfiguredProviders(list);
      setImportResult({
        ok: true,
        message: `Ключ импортирован (${targetProvider === 'starimg' ? 'Starimg AI' : 'свой endpoint'}), модель: ${res.model || '—'}`,
      });

      if (res.apiKey) {
        void handleFetchModels({
          provider: targetProvider,
          baseUrl: res.baseUrl || settings.baseUrl,
          apiKey: res.apiKey,
        });
      }
    } finally {
      setImporting(false);
    }
  };

  // Обновляем черновик при смене провайдера
  useEffect(() => {
    setKeyDraft(settings.apiKey);
    setModelFetchResult(null);
    setModelSearch('');
  }, [settings.provider, settings.apiKey]);

  const handleFetchModels = async (override?: Partial<AISettings>) => {
    setLoadingModels(true);
    setModelFetchResult(null);
    try {
      const target = { ...settings, apiKey: keyDraft.trim(), ...override };
      const res = onFetchModels
        ? await onFetchModels(target)
        : await aiService.fetchModels(target);

      if (res.success && res.models && res.models.length > 0) {
        setModelFetchResult({
          ok: true,
          count: res.models.length,
          message: `Загружено ${res.models.length} моделей от API`,
        });
      } else {
        setModelFetchResult({
          ok: false,
          message: res.error || 'Провайдер не вернул список моделей',
        });
      }
    } catch (err) {
      setModelFetchResult({
        ok: false,
        message: (err as Error).message || 'Ошибка загрузки моделей',
      });
    } finally {
      setLoadingModels(false);
    }
  };

  const handleSaveKey = async () => {
    try {
      await onSaveApiKey(keyDraft.trim());
      void keyVault.listConfigured().then(setConfiguredProviders);
      setSavedFlash(true);
      window.setTimeout(() => setSavedFlash(false), 1800);
      if (keyDraft.trim()) {
        void handleFetchModels({ apiKey: keyDraft.trim() });
      }
    } catch (err) {
      setTestResult({
        ok: false,
        message: 'Не удалось сохранить ключ: ' + ((err as Error).message || 'ошибка записи'),
      });
    }
  };

  const handleTest = async () => {
    setTesting(true);
    setTestResult(null);
    // Тестируем с ключом из черновика, не дожидаясь сохранения
    const effective = { ...settings, apiKey: keyDraft.trim() };
    const result = await aiService.testConnection(effective);
    setTestResult(result);
    setTesting(false);
    if (result.ok) {
      void handleFetchModels(effective);
    }
  };

  // При открытии вкладки «Модель» загружаем список, если он ещё не получен
  useEffect(() => {
    if (tab === 'model') {
      const currentModels = modelsCache?.[settings.provider] || [];
      if (currentModels.length === 0 && (keyDraft.trim() || settings.apiKey || !preset.requiresKey)) {
        void handleFetchModels();
      }
    }
  }, [tab, settings.provider]);

  const handleClearKey = async (provider: AIProvider) => {
    await keyVault.set(provider, '');
    const list = await keyVault.listConfigured();
    setConfiguredProviders(list);
    if (provider === settings.provider) {
      setKeyDraft('');
      await onSaveApiKey('');
    }
  };

  const tabs: { id: Tab; label: string; icon: React.ReactNode }[] = [
    { id: 'provider', label: 'Провайдер', icon: <KeyRound size={13} /> },
    { id: 'model', label: 'Модель', icon: <Sliders size={13} /> },
    { id: 'prompt', label: 'Промт', icon: <Sparkles size={13} /> },
    { id: 'keys', label: 'Хранилище', icon: <ShieldCheck size={13} /> },
    { id: 'integrations', label: 'Сервисы', icon: <Cloud size={13} /> },
    { id: 'system', label: 'Система', icon: <Monitor size={13} /> },
  ];

  return (
    <Modal
      open={open}
      onClose={onClose}
      width="max-w-2xl"
      icon={<SettingsIcon size={15} className="text-ws-accent" />}
      title="Настройки"
      footer={
        <>
          <Button variant="ghost" onClick={onClose}>
            Закрыть
          </Button>
          <Button
            variant="primary"
            onClick={async () => {
              await handleSaveKey();
              onClose();
            }}
          >
            <Check size={13} />
            Сохранить
          </Button>
        </>
      }
    >
      {/* Вкладки */}
      <div className="mb-4 flex items-center gap-1 border-b border-ws pb-2">
        {tabs.map((t) => (
          <button
            key={t.id}
            onClick={() => setTab(t.id)}
            className={cn(
              'flex items-center gap-1.5 rounded px-2.5 py-1.5 text-xs font-medium transition-colors',
              tab === t.id
                ? 'bg-ws-accent-soft text-ws-accent'
                : 'text-ws-muted hover:bg-ws-hover hover:text-ws'
            )}
          >
            {t.icon}
            {t.label}
          </button>
        ))}
      </div>

      {/* ---- Провайдер и ключ ---- */}
      {tab === 'provider' && (
        <div className="space-y-4">
          {/* Импорт из локального файла секретов */}
          {secrets?.found && secrets.hasKey && (
            <div className="rounded-lg border border-ws bg-ws-elevated p-3">
              <div className="flex items-start justify-between gap-3">
                <div className="min-w-0">
                  <div className="flex items-center gap-1.5 text-xs font-medium text-ws">
                    <FileKey size={13} className="text-ws-accent" />
                    Найден файл с ключами
                  </div>
                  <p className="mt-1 truncate font-mono text-[10px] text-ws-faint" title={secrets.filePath}>
                    {secrets.filePath}
                  </p>
                  {(secrets.baseUrl || secrets.textModel) && (
                    <p className="mt-1 text-[10px] text-ws-muted">
                      {secrets.baseUrl && <span className="font-mono">{secrets.baseUrl}</span>}
                      {secrets.textModel && <span> · модель {secrets.textModel}</span>}
                    </p>
                  )}
                </div>
                <Button variant="primary" onClick={handleImportSecrets} disabled={importing}>
                  {importing ? <Loader2 size={12} className="animate-spin" /> : <Download size={12} />}
                  Импортировать
                </Button>
              </div>

              {secrets.hasOtherSecrets && (
                <p className="mt-2 flex items-start gap-1.5 rounded border border-amber-500/30 bg-amber-500/10 px-2 py-1.5 text-[10px] leading-relaxed text-ws-warning">
                  <AlertTriangle size={11} className="mt-0.5 shrink-0" />
                  В файле есть и другие ключи (Notion, Telegram) — импортируется только API-ключ
                  нейросети.
                </p>
              )}
            </div>
          )}

          {secrets && !secrets.found && (
            <div className="rounded-lg border border-ws bg-ws-elevated px-3 py-2 text-[11px] text-ws-muted">
              <span className="flex items-center gap-1.5">
                <Info size={12} className="shrink-0 text-ws-faint" />
                Файл secrets.local.ps1 не найден — введите ключ вручную ниже.
              </span>
            </div>
          )}

          {importResult && (
            <div
              className={cn(
                'flex items-start gap-2 rounded-lg border px-3 py-2 text-[11px]',
                importResult.ok
                  ? 'border-emerald-500/40 bg-emerald-500/10 text-ws-success'
                  : 'border-red-500/40 bg-red-500/10 text-ws-danger'
              )}
            >
              {importResult.ok ? (
                <Check size={13} className="mt-0.5 shrink-0" />
              ) : (
                <AlertTriangle size={13} className="mt-0.5 shrink-0" />
              )}
              <span className="select-text">{importResult.message}</span>
            </div>
          )}

          <Field
            label="Провайдер модели"
            hint="Ключи хранятся только на этом компьютере в зашифрованном виде и отправляются напрямую выбранному провайдеру."
          >
            <div className="grid grid-cols-2 gap-2">
              {PROVIDER_ORDER.map((p) => {
                const info = PROVIDER_PRESETS[p];
                const isActive = settings.provider === p;
                const hasKey = configuredProviders.includes(p);
                return (
                  <button
                    key={p}
                    onClick={() => void onChangeProvider(p)}
                    className={cn(
                      'flex items-center justify-between gap-2 rounded-lg border px-3 py-2 text-left transition-colors',
                      isActive
                        ? 'border-ws-accent bg-ws-accent-soft'
                        : 'border-ws bg-ws-elevated hover:border-ws-strong'
                    )}
                  >
                    <div className="min-w-0">
                      <div
                        className={cn(
                          'truncate text-xs font-medium',
                          isActive ? 'text-ws-accent' : 'text-ws'
                        )}
                      >
                        {info.label}
                      </div>
                      <div className="truncate text-[10px] text-ws-faint">
                        {info.requiresKey ? info.keyHint : 'без ключа'}
                      </div>
                    </div>
                    {hasKey && <Check size={13} className="shrink-0 text-ws-success" />}
                  </button>
                );
              })}
            </div>
          </Field>

          <Field
            label={`API-ключ — ${preset.label}`}
            hint={
              preset.requiresKey
                ? 'Ключ сохраняется рядом с приложением и подхватывается при следующем запуске.'
                : 'Для локальных моделей ключ обычно не требуется — можно оставить поле пустым.'
            }
          >
            <div className="flex gap-2">
              <div className="relative flex-1">
                <Input
                  type={keyVisible ? 'text' : 'password'}
                  value={keyDraft}
                  onChange={(e) => setKeyDraft(e.target.value)}
                  placeholder={preset.keyHint}
                  spellCheck={false}
                  autoComplete="off"
                  className="pr-9 font-mono"
                />
                <button
                  onClick={() => setKeyVisible((v) => !v)}
                  title={keyVisible ? 'Скрыть ключ' : 'Показать ключ'}
                  className="absolute top-1/2 right-2 -translate-y-1/2 text-ws-faint hover:text-ws"
                >
                  {keyVisible ? <EyeOff size={13} /> : <Eye size={13} />}
                </button>
              </div>
              <Button variant="secondary" onClick={handleSaveKey}>
                {savedFlash ? <Check size={13} className="text-ws-success" /> : 'Сохранить'}
              </Button>
            </div>
          </Field>

          <Field label="Базовый URL API" hint="Можно указать свой шлюз или локальный сервер, совместимый с OpenAI.">
            <Input
              value={settings.baseUrl}
              onChange={(e) => onUpdateSettings({ baseUrl: e.target.value })}
              placeholder={preset.baseUrl}
              spellCheck={false}
              className="font-mono"
            />
          </Field>

          <div className="flex items-center justify-between rounded-lg border border-ws bg-ws-elevated px-3 py-2">
            <div className="flex items-center gap-2">
              <Button variant="secondary" onClick={handleTest} disabled={testing}>
                {testing ? <Loader2 size={13} className="animate-spin" /> : <RefreshCw size={13} />}
                Проверить подключение
              </Button>
              {preset.docsUrl && (
                <a
                  href={preset.docsUrl}
                  target="_blank"
                  rel="noreferrer"
                  className="flex items-center gap-1 text-[11px] text-ws-accent hover:underline"
                >
                  Получить ключ
                  <ExternalLink size={10} />
                </a>
              )}
            </div>
            <span className="text-[11px] text-ws-faint">{preset.label}</span>
          </div>

          {testResult && (
            <div
              className={cn(
                'flex items-start gap-2 rounded-lg border px-3 py-2 text-[11px]',
                testResult.ok
                  ? 'border-emerald-500/40 bg-emerald-500/10 text-ws-success'
                  : 'border-red-500/40 bg-red-500/10 text-ws-danger'
              )}
            >
              {testResult.ok ? (
                <Check size={13} className="mt-0.5 shrink-0" />
              ) : (
                <AlertTriangle size={13} className="mt-0.5 shrink-0" />
              )}
              <span className="select-text">{testResult.message}</span>
            </div>
          )}
        </div>
      )}

      {/* ---- Модель и параметры ---- */}
      {tab === 'model' && (
        <div className="space-y-4">
          <div className="space-y-1.5">
            <div className="flex items-center justify-between">
              <label className="text-xs font-semibold text-ws">Модель ({preset.label})</label>
              <button
                type="button"
                onClick={() => void handleFetchModels()}
                disabled={loadingModels}
                className="flex items-center gap-1.5 text-[11px] text-ws-accent hover:underline disabled:opacity-50"
              >
                {loadingModels ? (
                  <Loader2 size={11} className="animate-spin" />
                ) : (
                  <RefreshCw size={11} />
                )}
                <span>{loadingModels ? 'Запрос к API...' : 'Обновить список из API'}</span>
              </button>
            </div>

            {providerModels.length > 8 && (
              <Input
                value={modelSearch}
                onChange={(e) => setModelSearch(e.target.value)}
                placeholder="Поиск модели по названию..."
                spellCheck={false}
                className="text-xs font-mono mb-1.5"
              />
            )}

            {providerModels.length > 0 ? (
              <Select
                value={settings.model}
                onChange={(e) => onUpdateSettings({ model: e.target.value })}
              >
                {filteredModels.map((m) => (
                  <option key={m} value={m}>
                    {m}
                  </option>
                ))}
                {filteredModels.length === 0 && (
                  <option value={settings.model} disabled>
                    Ничего не найдено (по фильтру «{modelSearch}»)
                  </option>
                )}
                {!filteredModels.includes(settings.model) && settings.model && (
                  <option value={settings.model}>{settings.model} (текущая)</option>
                )}
              </Select>
            ) : (
              <Input
                value={settings.model}
                onChange={(e) => onUpdateSettings({ model: e.target.value })}
                placeholder="например, gpt-4o или deepseek-chat"
                spellCheck={false}
                className="font-mono"
              />
            )}

            {modelFetchResult && (
              <div
                className={cn(
                  'flex items-center gap-2 rounded border px-3 py-1.5 text-[11px] mt-1',
                  modelFetchResult.ok
                    ? 'border-emerald-500/40 bg-emerald-500/10 text-ws-success'
                    : 'border-amber-500/40 bg-amber-500/10 text-ws-warning'
                )}
              >
                {modelFetchResult.ok ? (
                  <Check size={12} className="shrink-0" />
                ) : (
                  <AlertTriangle size={12} className="shrink-0" />
                )}
                <span>{modelFetchResult.message}</span>
              </div>
            )}

            {!modelFetchResult && (modelsCache?.[settings.provider]?.length ?? 0) > 0 && (
              <p className="text-[11px] text-ws-faint mt-1">
                ✓ Загружено {modelsCache![settings.provider].length} моделей от API ({preset.label})
              </p>
            )}
            {!modelFetchResult && (modelsCache?.[settings.provider]?.length ?? 0) === 0 && (
              <p className="text-[11px] text-ws-faint mt-1">
                Показаны стандартные пресеты. Нажмите «Обновить список из API», чтобы загрузить все доступные модели вашего тарифа/ключа.
              </p>
            )}
          </div>

          <Field label="Своя модель" hint="Добавьте идентификатор, если его нет в списке выше.">
            <div className="flex gap-2">
              <Input
                value={customModel}
                onChange={(e) => setCustomModel(e.target.value)}
                placeholder="org/model-name"
                spellCheck={false}
                className="font-mono"
              />
              <Button
                variant="secondary"
                disabled={!customModel.trim()}
                onClick={() => {
                  onUpdateSettings({ model: customModel.trim() });
                  setCustomModel('');
                }}
              >
                Применить
              </Button>
            </div>
          </Field>

          <Field label={`Температура: ${settings.temperature.toFixed(1)}`} hint="Ниже — точнее и стабильнее, выше — креативнее.">
            <input
              type="range"
              min={0}
              max={2}
              step={0.1}
              value={settings.temperature}
              onChange={(e) => onUpdateSettings({ temperature: Number(e.target.value) })}
              className="w-full accent-[var(--ws-accent)]"
            />
            <div className="flex justify-between text-[10px] text-ws-faint">
              <span>Точно (0.0)</span>
              <span>Баланс (1.0)</span>
              <span>Креативно (2.0)</span>
            </div>
          </Field>

          <Field label="Максимум токенов в ответе" hint="Ограничивает длину ответа модели.">
            <Input
              type="number"
              min={256}
              max={32000}
              step={256}
              value={settings.maxTokens}
              onChange={(e) => onUpdateSettings({ maxTokens: Number(e.target.value) || 4096 })}
            />
          </Field>
        </div>
      )}

      {/* ---- Системный промт ---- */}
      {tab === 'prompt' && (
        <div className="space-y-4">
          <Field label="Роль ассистента" hint="Быстрый выбор специализации — подставит готовый системный промт.">
            <div className="grid grid-cols-1 gap-2">
              {SYSTEM_PRESETS.map((p) => (
                <button
                  key={p.id}
                  onClick={() =>
                    onUpdateSettings({ systemPrompt: p.prompt, activePresetId: p.id })
                  }
                  className={cn(
                    'rounded-lg border px-3 py-2 text-left transition-colors',
                    settings.activePresetId === p.id
                      ? 'border-ws-accent bg-ws-accent-soft'
                      : 'border-ws bg-ws-elevated hover:border-ws-strong'
                  )}
                >
                  <div
                    className={cn(
                      'text-xs font-medium',
                      settings.activePresetId === p.id ? 'text-ws-accent' : 'text-ws'
                    )}
                  >
                    {p.name}
                  </div>
                  <div className="mt-0.5 text-[10px] text-ws-faint">{p.description}</div>
                </button>
              ))}
            </div>
          </Field>

          <Field label="Системный промт" hint="Задаёт роль и стиль ответов. Применяется ко всем диалогам.">
            <Textarea
              value={settings.systemPrompt}
              onChange={(e) =>
                onUpdateSettings({ systemPrompt: e.target.value, activePresetId: undefined })
              }
              rows={7}
              placeholder="Ты — профессиональный ассистент…"
            />
          </Field>
        </div>
      )}

      {/* ---- Хранилище ключей ---- */}
      {tab === 'keys' && (
        <div className="space-y-4">
          <div className="flex items-start gap-2 rounded-lg border border-ws bg-ws-elevated px-3 py-2.5">
            <ShieldCheck size={15} className="mt-0.5 shrink-0 text-ws-success" />
            <div className="text-[11px] leading-relaxed text-ws-muted">
              Все ключи хранятся <strong className="text-ws">только на этом компьютере</strong> в
              зашифрованном виде. WorkSpace не отправляет их никуда, кроме API выбранного вами
              провайдера.
            </div>
          </div>

          <Field label="Сохранённые ключи">
            {configuredProviders.length === 0 ? (
              <p className="text-[11px] text-ws-faint">Ключи ещё не сохранены.</p>
            ) : (
              <div className="space-y-1.5">
                {configuredProviders.map((p) => (
                  <div
                    key={p}
                    className="flex items-center justify-between rounded-lg border border-ws bg-ws-elevated px-3 py-2"
                  >
                    <div className="flex items-center gap-2">
                      <Check size={13} className="text-ws-success" />
                      <span className="text-xs text-ws">{PROVIDER_PRESETS[p].label}</span>
                      {p === settings.provider && (
                        <span className="rounded bg-ws-accent-soft px-1.5 py-0.5 text-[10px] text-ws-accent">
                          активный
                        </span>
                      )}
                    </div>
                    <button
                      onClick={() => void handleClearKey(p)}
                      title="Удалить сохранённый ключ"
                      className="text-ws-faint transition-colors hover:text-ws-danger"
                    >
                      <Trash2 size={13} />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </Field>
        </div>
      )}

      {tab === 'integrations' && <IntegrationsSettings onNotify={onNotify || (() => undefined)} />}
      {tab === 'system' && <SystemSettings onNotify={onNotify} />}
    </Modal>
  );
};
