import React, { useEffect, useRef } from 'react';
import { cn } from '../../lib/utils';

export interface ContextMenuItem {
  id: string;
  label: string;
  icon?: React.ReactNode;
  danger?: boolean;
  disabled?: boolean;
  separatorBefore?: boolean;
}

export const ContextMenu: React.FC<{
  x: number;
  y: number;
  items: ContextMenuItem[];
  onSelect: (id: string) => void;
  onClose: () => void;
}> = ({ x, y, items, onSelect, onClose }) => {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const outside = (e: MouseEvent) => { if (!ref.current?.contains(e.target as Node)) onClose(); };
    const key = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose(); };
    window.addEventListener('mousedown', outside);
    window.addEventListener('keydown', key);
    return () => { window.removeEventListener('mousedown', outside); window.removeEventListener('keydown', key); };
  }, [onClose]);

  const left = Math.max(8, Math.min(x, window.innerWidth - 240));
  const top = Math.max(8, Math.min(y, window.innerHeight - items.length * 34 - 16));
  return (
    <div ref={ref} className="fixed z-[100] min-w-[220px] rounded-lg border border-ws bg-ws-surface p-1 shadow-ws" style={{left,top}} role="menu">
      {items.map((item) => (
        <React.Fragment key={item.id}>
          {item.separatorBefore && <div className="my-1 border-t border-ws" />}
          <button role="menuitem" disabled={item.disabled} onClick={() => { onSelect(item.id); onClose(); }} className={cn('flex w-full items-center gap-2 rounded px-2.5 py-1.5 text-left text-xs transition-colors disabled:opacity-40', item.danger ? 'text-ws-danger hover:bg-red-500/10' : 'text-ws hover:bg-ws-hover')}>
            <span className="w-4 shrink-0">{item.icon}</span><span>{item.label}</span>
          </button>
        </React.Fragment>
      ))}
    </div>
  );
};
