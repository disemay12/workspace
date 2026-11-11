/**
 * Настройки внешнего вида и уведомлений.
 *
 * Держим отдельной вкладкой: SystemSettings отвечает за обновления и ярлыки,
 * а оформление меняется часто и не должно требовать прокрутки того экрана.
 */
import React from 'react';
import { Check } from 'lucide-react';
import { Button } from '../ui';
import { usePreferences } from '../../state/preferences';
import { applyAppearance, clampScale, normalizeAccent, UI_SCALE_MAX, UI_SCALE_MIN } from '../../lib/appearance';

interface Props {
  onNotify?: (message: string) => void;
}

const THEMES: { id: 'system' | 'dark' | 'light'; label: string }[] = [
  { id: 'system', label: 'Как в системе' },
  { id: 'dark', label: 'Тёмная' },
  { id: 'light', label: 'Светлая' },
];

export const AppearanceSettings: React.FC<Props> = ({ onNotify }) => {
  const [preferences, updatePreferences] = usePreferences();
  const appearance = preferences.appearance;
  const notifications = preferences.notifications;

  const patchAppearance = (patch: Partial<typeof appearance>) => {
    updatePreferences({ appearance: { ...appearance, ...patch } });
    const systemDark = window.matchMedia?.('(prefers-color-scheme: dark)').matches ?? false;
    applyAppearance({ ...appearance, ...patch }, systemDark);
  };

  return (
    <div className="space-y-4 text-xs text-ws">
      <section className="rounded-lg border border-ws bg-ws-surface p-3">
        <div className="mb-2 text-[11px] font-semibold tracking-wider text-ws-muted uppercase">Тема</div>
        <div className="flex flex-wrap gap-1">
          {THEMES.map((theme) => (
            <button
              key={theme.id}
              onClick={() => patchAppearance({ theme: theme.id })}
              className={
                appearance.theme === theme.id
                  ? 'flex items-center gap-1 rounded border border-ws-accent bg-ws-accent-soft px-2 py-1 text-ws-accent'
                  : 'flex items-center gap-1 rounded border border-ws px-2 py-1 text-ws-muted hover:bg-ws-hover'
              }
            >
              {appearance.theme === theme.id && <Check size={12} />}
              {theme.label}
            </button>
          ))}
        </div>
        <div className="mt-3 grid grid-cols-2 gap-3">
          <label className="flex flex-col gap-1">
            <span className="text-ws-muted">Акцент</span>
            <div className="flex items-center gap-2">
              <input
                type="color"
                value={normalizeAccent(appearance.accent) || '#3B82F6'}
                onChange={(event) => patchAppearance({ accent: event.target.value })}
                className="h-7 w-10 cursor-pointer rounded border border-ws bg-ws-surface"
              />
              <input
                value={appearance.accent}
                onChange={(event) => patchAppearance({ accent: event.target.value })}
                placeholder="по умолчанию"
                className="min-w-0 flex-1 rounded border border-ws bg-ws-editor px-2 py-1 font-mono text-ws"
              />
            </div>
          </label>
          <label className="flex flex-col gap-1">
            <span className="text-ws-muted">Плотность</span>
            <select
              value={appearance.density}
              onChange={(event) => patchAppearance({ density: event.target.value as typeof appearance.density })}
              className="rounded border border-ws bg-ws-editor px-2 py-1 text-ws"
            >
              <option value="comfortable">Обычная</option>
              <option value="compact">Компактная</option>
            </select>
          </label>
          <label className="flex flex-col gap-1">
            <span className="text-ws-muted">
              Масштаб интерфейса: {Math.round(appearance.uiScale * 100)}%
            </span>
            <input
              type="range"
              min={UI_SCALE_MIN}
              max={UI_SCALE_MAX}
              step={0.05}
              value={appearance.uiScale}
              onChange={(event) => patchAppearance({ uiScale: clampScale(Number(event.target.value)) })}
              className="accent-[var(--ws-accent)]"
            />
          </label>
          <label className="flex items-center gap-2 pt-5">
            <input
              type="checkbox"
              checked={appearance.reduceMotion}
              onChange={(event) => patchAppearance({ reduceMotion: event.target.checked })}
            />
            <span className="text-ws-muted">Меньше анимации</span>
          </label>
        </div>
      </section>

      <section className="rounded-lg border border-ws bg-ws-surface p-3">
        <div className="mb-2 text-[11px] font-semibold tracking-wider text-ws-muted uppercase">
          Уведомления
        </div>
        <div className="grid grid-cols-2 gap-3">
          <label className="flex items-center gap-2">
            <input
              type="checkbox"
              checked={notifications.toasts}
              onChange={(event) =>
                updatePreferences({ notifications: { ...notifications, toasts: event.target.checked } })
              }
            />
            <span className="text-ws-muted">Показывать тосты</span>
          </label>
          <label className="flex flex-col gap-1">
            <span className="text-ws-muted">Сигнал</span>
            <select
              value={notifications.sound}
              onChange={(event) =>
                updatePreferences({
                  notifications: {
                    ...notifications,
                    sound: event.target.value as typeof notifications.sound,
                  },
                })
              }
              className="rounded border border-ws bg-ws-editor px-2 py-1 text-ws"
            >
              <option value="off">Без звука</option>
              <option value="knock">Тихий стук</option>
              <option value="soft">Мягкий сигнал</option>
            </select>
          </label>
          <label className="flex flex-col gap-1">
            <span className="text-ws-muted">Громкость: {Math.round(notifications.volume * 100)}%</span>
            <input
              type="range"
              min={0.05}
              max={0.6}
              step={0.05}
              value={notifications.volume}
              onChange={(event) =>
                updatePreferences({
                  notifications: { ...notifications, volume: Number(event.target.value) },
                })
              }
              className="accent-[var(--ws-accent)]"
            />
          </label>
          <div className="flex items-end gap-2">
            <label className="flex flex-1 flex-col gap-1">
              <span className="text-ws-muted">Тишина с</span>
              <input
                type="time"
                value={notifications.quietFrom}
                onChange={(event) =>
                  updatePreferences({
                    notifications: { ...notifications, quietFrom: event.target.value },
                  })
                }
                className="rounded border border-ws bg-ws-editor px-2 py-1 text-ws"
              />
            </label>
            <label className="flex flex-1 flex-col gap-1">
              <span className="text-ws-muted">по</span>
              <input
                type="time"
                value={notifications.quietTo}
                onChange={(event) =>
                  updatePreferences({ notifications: { ...notifications, quietTo: event.target.value } })
                }
                className="rounded border border-ws bg-ws-editor px-2 py-1 text-ws"
              />
            </label>
          </div>
        </div>
        <div className="mt-3 flex flex-wrap gap-3">
          {(
            [
              ['save', 'Сохранение'],
              ['export', 'Экспорт'],
              ['ai', 'Ответ AI'],
              ['update', 'Обновления'],
              ['error', 'Ошибки'],
            ] as const
          ).map(([key, label]) => (
            <label key={key} className="flex items-center gap-1.5">
              <input
                type="checkbox"
                checked={notifications.events[key]}
                onChange={(event) =>
                  updatePreferences({
                    notifications: {
                      ...notifications,
                      events: { ...notifications.events, [key]: event.target.checked },
                    },
                  })
                }
              />
              <span className="text-ws-muted">{label}</span>
            </label>
          ))}
        </div>
        <Button
          className="mt-3"
          onClick={() => onNotify?.('Проверка уведомления: тихий сигнал и тост включены')}
        >
          Проверить уведомление
        </Button>
      </section>
    </div>
  );
};
