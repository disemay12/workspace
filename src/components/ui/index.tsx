import React from 'react';
import {
  AlertTriangle,
  Check,
  Info,
  Loader2,
  Sparkles,
  X,
  Search,
  FileCode2,
  BookOpen,
  Table,
  ScrollText,
  PenLine,
  Bot,
} from 'lucide-react';
import { cn } from '../../lib/utils';

/* ============================================================
   Базовые UI-примитивы
   ============================================================ */

export const Button: React.FC<
  React.ButtonHTMLAttributes<HTMLButtonElement> & {
    variant?: 'primary' | 'secondary' | 'ghost' | 'danger';
    size?: 'sm' | 'md';
  }
> = ({ variant = 'secondary', size = 'sm', className, children, ...rest }) => {
  const base =
    'inline-flex items-center justify-center gap-1.5 rounded font-medium transition-colors disabled:cursor-not-allowed disabled:opacity-45';
  const sizes = { sm: 'px-2.5 py-1 text-xs', md: 'px-3.5 py-1.5 text-sm' };
  const variants = {
    primary: 'bg-ws-accent text-white hover:bg-ws-accent hover:opacity-90',
    secondary: 'bg-ws-panel-alt text-ws border border-ws hover:bg-ws-hover',
    ghost: 'text-ws-muted hover:bg-ws-hover hover:text-ws',
    danger: 'bg-red-600 text-white hover:bg-red-500',
  };
  return (
    <button className={cn(base, sizes[size], variants[variant], className)} {...rest}>
      {children}
    </button>
  );
};

export const IconButton: React.FC<
  React.ButtonHTMLAttributes<HTMLButtonElement> & { active?: boolean; title: string }
> = ({ active, className, children, ...rest }) => (
  <button
    className={cn(
      'rounded p-1.5 transition-colors hover:bg-ws-hover',
      active ? 'text-ws-accent' : 'text-ws-faint hover:text-ws',
      className
    )}
    {...rest}
  >
    {children}
  </button>
);

export const Input: React.FC<React.InputHTMLAttributes<HTMLInputElement>> = ({
  className,
  ...rest
}) => (
  <input
    className={cn(
      'w-full rounded border border-ws bg-ws-elevated px-2.5 py-1.5 text-xs text-ws placeholder:text-ws-faint',
      'focus:border-ws-accent focus:outline-none transition-colors',
      className
    )}
    {...rest}
  />
);

export const Textarea = React.forwardRef<
  HTMLTextAreaElement,
  React.TextareaHTMLAttributes<HTMLTextAreaElement>
>(({ className, ...rest }, ref) => (
  <textarea
    ref={ref}
    className={cn(
      'w-full resize-none rounded border border-ws bg-ws-elevated px-2.5 py-1.5 text-xs text-ws placeholder:text-ws-faint',
      'focus:border-ws-accent focus:outline-none transition-colors',
      className
    )}
    {...rest}
  />
));
Textarea.displayName = 'Textarea';

export const Select: React.FC<React.SelectHTMLAttributes<HTMLSelectElement>> = ({
  className,
  children,
  ...rest
}) => (
  <select
    className={cn(
      'w-full rounded border border-ws bg-ws-elevated px-2 py-1.5 text-xs text-ws',
      'focus:border-ws-accent focus:outline-none transition-colors',
      className
    )}
    {...rest}
  >
    {children}
  </select>
);

export const Field: React.FC<{
  label: string;
  hint?: string;
  children: React.ReactNode;
}> = ({ label, hint, children }) => (
  <div className="space-y-1.5">
    <label className="block text-[11px] font-semibold tracking-wide text-ws-muted uppercase">
      {label}
    </label>
    {children}
    {hint && <p className="text-[11px] leading-relaxed text-ws-faint">{hint}</p>}
  </div>
);

export const Spinner: React.FC<{ size?: number; className?: string }> = ({
  size = 14,
  className,
}) => <Loader2 size={size} className={cn('animate-spin', className)} />;

/** Всплывающее уведомление в нижней части окна. */
export const Toast: React.FC<{
  message: string;
  tone?: 'info' | 'success' | 'error';
}> = ({ message, tone = 'info' }) => {
  const icons = {
    info: <Info size={14} className="text-ws-accent" />,
    success: <Check size={14} className="text-ws-success" />,
    error: <AlertTriangle size={14} className="text-ws-danger" />,
  };
  return (
    <div className="animate-fade-in pointer-events-none fixed bottom-8 left-1/2 z-50 -translate-x-1/2">
      <div className="flex items-center gap-2 rounded-lg border border-ws bg-ws-surface px-3.5 py-2 text-xs text-ws shadow-ws">
        {icons[tone]}
        <span>{message}</span>
      </div>
    </div>
  );
};

/** Модальное окно с затемнением. */
export const Modal: React.FC<{
  open: boolean;
  onClose: () => void;
  title: React.ReactNode;
  icon?: React.ReactNode;
  children: React.ReactNode;
  footer?: React.ReactNode;
  width?: string;
}> = ({ open, onClose, title, icon, children, footer, width = 'max-w-lg' }) => {
  if (!open) return null;
  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/55 p-4 backdrop-blur-[2px]"
      onMouseDown={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        className={cn(
          'animate-scale-in flex max-h-[85vh] w-full flex-col overflow-hidden rounded-xl border border-ws bg-ws-surface shadow-ws',
          width
        )}
      >
        <div className="flex items-center justify-between border-b border-ws px-4 py-3">
          <div className="flex items-center gap-2 text-sm font-semibold text-ws">
            {icon}
            {title}
          </div>
          <IconButton title="Закрыть" onClick={onClose}>
            <X size={14} />
          </IconButton>
        </div>
        <div className="flex-1 overflow-y-auto p-4">{children}</div>
        {footer && (
          <div className="flex items-center justify-end gap-2 border-t border-ws px-4 py-3">
            {footer}
          </div>
        )}
      </div>
    </div>
  );
};

/* ============================================================
   AI-компоненты
   ============================================================ */

export const AIBadge: React.FC<{ label?: string; className?: string }> = ({
  label = 'AI',
  className,
}) => (
  <span
    className={cn(
      'inline-flex items-center gap-1 rounded bg-ws-accent-soft px-1.5 py-0.5 text-[10px] font-semibold text-ws-accent',
      className
    )}
  >
    <Sparkles size={10} />
    {label}
  </span>
);

/** Пресеты ролей ассистента с иконками. */
export const PRESET_ICONS: Record<string, React.ReactNode> = {
  general: <Bot size={13} />,
  excel: <Table size={13} />,
  logs: <ScrollText size={13} />,
  copywriter: <PenLine size={13} />,
  coder: <FileCode2 size={13} />,
};

export const EmptyState: React.FC<{
  icon: React.ReactNode;
  title: string;
  description?: string;
  action?: React.ReactNode;
}> = ({ icon, title, description, action }) => (
  <div className="flex flex-1 flex-col items-center justify-center gap-3 p-8 text-center">
    <div className="text-ws-faint">{icon}</div>
    <p className="text-sm font-medium text-ws">{title}</p>
    {description && <p className="max-w-xs text-xs leading-relaxed text-ws-muted">{description}</p>}
    {action}
  </div>
);

export { Search, BookOpen };
