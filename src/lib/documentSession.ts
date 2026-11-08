/**
 * Реестр открытых документов.
 *
 * Даёт единая точка, где приложение и агент узнают тип документа, путь,
 * доступные инструменты и состояние прав. Контроллеры редакторов остаются
 * источниками живых данных; сессия описывает документ и его возможности.
 */
import type { FileType } from '../types';

export type DocumentKind = 'word' | 'markdown' | 'excel' | 'powerpoint' | 'text';

export interface DocumentCapabilities {
  read: boolean;
  write: boolean;
  tools: string[];
}

export interface DocumentHistoryState {
  canUndo: boolean;
  canRedo: boolean;
  size: number;
}

export interface DocumentSession {
  id: string;
  kind: DocumentKind;
  path: string;
  name: string;
  capabilities: DocumentCapabilities;
  model: unknown;
  selection: unknown;
  history: DocumentHistoryState;
  updatedAt: number;
}

export interface DocumentSessionInput {
  id: string;
  kind: DocumentKind;
  path: string;
  name: string;
  writeAccess: boolean;
  tools?: string[];
}

const sessions = new Map<string, DocumentSession>();

export function kindForTab(type: FileType | string): DocumentKind {
  if (type === 'word') return 'word';
  if (type === 'excel') return 'excel';
  if (type === 'powerpoint') return 'powerpoint';
  if (type === 'markdown' || type === 'richtext') return 'markdown';
  return 'text';
}

export function registerDocumentSession(input: DocumentSessionInput): DocumentSession {
  const existing = sessions.get(input.id);
  const session: DocumentSession = {
    id: input.id,
    kind: input.kind,
    path: input.path,
    name: input.name,
    capabilities: {
      read: true,
      write: input.writeAccess,
      tools: input.tools ?? existing?.capabilities.tools ?? [],
    },
    model: existing?.model ?? null,
    selection: existing?.selection ?? null,
    history: existing?.history ?? { canUndo: false, canRedo: false, size: 0 },
    updatedAt: Date.now(),
  };
  sessions.set(input.id, session);
  return session;
}

export function getDocumentSession(id: string): DocumentSession | undefined {
  return sessions.get(id);
}

export function updateDocumentSession(
  id: string,
  patch: Partial<Pick<DocumentSession, 'selection' | 'history' | 'path' | 'name'>> & {
    capabilities?: Partial<DocumentCapabilities>;
  }
): DocumentSession | undefined {
  const current = sessions.get(id);
  if (!current) return undefined;
  const next: DocumentSession = {
    ...current,
    ...patch,
    capabilities: { ...current.capabilities, ...patch.capabilities },
    history: { ...current.history, ...patch.history },
    updatedAt: Date.now(),
  };
  sessions.set(id, next);
  return next;
}

export function removeDocumentSession(id: string): void {
  sessions.delete(id);
}

export function listDocumentSessions(): DocumentSession[] {
  return [...sessions.values()];
}

export function resetDocumentSessions(): void {
  sessions.clear();
}
