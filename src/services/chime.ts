/**
 * Тихий сигнал уведомления: короткий синтезированный звук вместо системных
 * звуков Windows (они громкие и раздражают). Пресеты: knock и soft.
 */
import type { NotificationKind } from '../lib/notifyRules';
import { readPreferences } from '../state/preferences';
import { shouldNotify, shouldPlaySound } from '../lib/notifyRules';

let audioContext: AudioContext | null = null;

function context(): AudioContext | null {
  if (typeof window === 'undefined') return null;
  try {
    if (!audioContext) {
      const Ctor = window.AudioContext ?? (window as unknown as { webkitAudioContext?: typeof AudioContext }).webkitAudioContext;
      if (!Ctor) return null;
      audioContext = new Ctor();
    }
    if (audioContext.state === 'suspended') void audioContext.resume();
    return audioContext;
  } catch {
    return null;
  }
}

/** Короткий деревянный «стук» или мягкий тон с затуханием. */
export function playChime(preset: 'knock' | 'soft', volume: number): void {
  const ctx = context();
  if (!ctx) return;
  const gain = ctx.createGain();
  const level = Math.min(0.6, Math.max(0.01, Number(volume) || 0.2));
  gain.connect(ctx.destination);
  const now = ctx.currentTime;

  if (preset === 'soft') {
    const osc = ctx.createOscillator();
    osc.type = 'sine';
    osc.frequency.setValueAtTime(660, now);
    osc.frequency.exponentialRampToValueAtTime(520, now + 0.16);
    gain.gain.setValueAtTime(0, now);
    gain.gain.linearRampToValueAtTime(level, now + 0.012);
    gain.gain.exponentialRampToValueAtTime(0.0001, now + 0.28);
    osc.connect(gain);
    osc.start(now);
    osc.stop(now + 0.3);
    return;
  }

  // knock: короткий щелчок из двух затухающих импульсов.
  const osc = ctx.createOscillator();
  osc.type = 'triangle';
  osc.frequency.setValueAtTime(180, now);
  osc.frequency.exponentialRampToValueAtTime(120, now + 0.05);
  gain.gain.setValueAtTime(0, now);
  gain.gain.linearRampToValueAtTime(level, now + 0.004);
  gain.gain.exponentialRampToValueAtTime(0.0001, now + 0.09);
  osc.connect(gain);
  osc.start(now);
  osc.stop(now + 0.1);
}

/** Проигрывает сигнал, если правила уведомлений это разрешают. */
export function playChimeFor(kind: NotificationKind): void {
  const prefs = readPreferences().notifications;
  if (!shouldNotify({ kind, prefs, now: new Date() })) return;
  if (!shouldPlaySound(prefs, kind)) return;
  if (prefs.sound === 'off') return;
  playChime(prefs.sound, prefs.volume);
}
