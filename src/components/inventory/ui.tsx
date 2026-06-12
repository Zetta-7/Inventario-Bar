"use client";

import { ReactNode } from "react";

export function Spinner({ size = "md" }: { size?: "sm" | "md" | "lg" }) {
  const s = { sm: "h-4 w-4 border-2", md: "h-5 w-5 border-2", lg: "h-12 w-12 border-4" }[size];
  return (
    <div className={`animate-spin rounded-full border-inv-accent border-t-transparent ${s}`} />
  );
}

export function Icon({
  name,
  className = "w-5 h-5",
}: {
  name:
    | "package"
    | "chart"
    | "refresh"
    | "settings"
    | "download"
    | "file"
    | "search"
    | "logout"
    | "close"
    | "check"
    | "alert"
    | "arrow-up"
    | "arrow-down"
    | "beer"
    | "plus"
    | "trash"
    | "menu"
    | "info"
    | "x-circle"
    | "trending";
  className?: string;
}) {
  const paths: Record<string, ReactNode> = {
    package: (
      <path strokeLinecap="round" strokeLinejoin="round" d="M20 7l-8-4-8 4m16 0l-8 4m8-4v10l-8 4m0-10L4 7m8 4v10M4 7v10l8 4" />
    ),
    chart: (
      <path strokeLinecap="round" strokeLinejoin="round" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
    ),
    refresh: (
      <path strokeLinecap="round" strokeLinejoin="round" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
    ),
    settings: (
      <path strokeLinecap="round" strokeLinejoin="round" d="M10.325 4.317c.426-1.756 2.924-1.756 3.35 0a1.724 1.724 0 002.573 1.066c1.543-.94 3.31.826 2.37 2.37a1.724 1.724 0 001.065 2.572c1.756.426 1.756 2.924 0 3.35a1.724 1.724 0 00-1.066 2.573c.94 1.543-.826 3.31-2.37 2.37a1.724 1.724 0 00-2.572 1.065c-.426 1.756-2.924 1.756-3.35 0a1.724 1.724 0 00-2.573-1.066c-1.543.94-3.31-.826-2.37-2.37a1.724 1.724 0 00-1.065-2.572c-1.756-.426-1.756-2.924 0-3.35a1.724 1.724 0 001.066-2.573c-.94-1.543.826-3.31 2.37-2.37.996.608 2.296.07 2.572-1.065z" />
    ),
    download: (
      <path strokeLinecap="round" strokeLinejoin="round" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
    ),
    file: (
      <path strokeLinecap="round" strokeLinejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
    ),
    search: (
      <path strokeLinecap="round" strokeLinejoin="round" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
    ),
    logout: (
      <path strokeLinecap="round" strokeLinejoin="round" d="M17 16l4-4m0 0l-4-4m4 4H7m6 4v1a3 3 0 01-3 3H6a3 3 0 01-3-3V7a3 3 0 013-3h4a3 3 0 013 3v1" />
    ),
    close: (
      <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
    ),
    check: (
      <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
    ),
    alert: (
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
    ),
    "arrow-up": (
      <path strokeLinecap="round" strokeLinejoin="round" d="M5 10l7-7m0 0l7 7m-7-7v18" />
    ),
    "arrow-down": (
      <path strokeLinecap="round" strokeLinejoin="round" d="M19 14l-7 7m0 0l-7-7m7 7V3" />
    ),
    beer: (
      <path strokeLinecap="round" strokeLinejoin="round" d="M9 2h6a2 2 0 012 2v1H7V4a2 2 0 012-2zm-2 5h10v8a4 4 0 01-4 4h-2a4 4 0 01-4-4V7zm12 2h1a2 2 0 012 2v3a2 2 0 01-2 2h-1" />
    ),
    plus: (
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 4v16m8-8H4" />
    ),
    trash: (
      <path strokeLinecap="round" strokeLinejoin="round" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
    ),
    menu: (
      <path strokeLinecap="round" strokeLinejoin="round" d="M4 6h16M4 12h16M4 18h16" />
    ),
    info: (
      <path strokeLinecap="round" strokeLinejoin="round" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
    ),
    "x-circle": (
      <path strokeLinecap="round" strokeLinejoin="round" d="M10 14l2-2m0 0l2-2m-2 2l-2-2m2 2l2 2m7-2a9 9 0 11-18 0 9 9 0 0118 0z" />
    ),
    trending: (
      <path strokeLinecap="round" strokeLinejoin="round" d="M13 7h8m0 0v8m0-8l-8 8-4-4-6 6" />
    ),
  };

  return (
    <svg className={className} fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.75}>
      {paths[name]}
    </svg>
  );
}

export function ModalShell({
  children,
  onClose,
  maxWidth = "max-w-4xl",
}: {
  children: ReactNode;
  onClose: () => void;
  maxWidth?: string;
}) {
  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 backdrop-blur-sm p-4 sm:p-6 animate-fade-in"
      onClick={onClose}
    >
      <div
        className={`relative w-full ${maxWidth} max-h-[90vh] overflow-auto inv-card animate-scale-in`}
        onClick={(e) => e.stopPropagation()}
      >
        {children}
      </div>
    </div>
  );
}

export function PageHeader({ title, description }: { title: string; description?: string }) {
  return (
    <div className="mb-6">
      <h2 className="text-xl sm:text-2xl font-semibold text-inv-text tracking-tight">{title}</h2>
      {description && <p className="text-sm text-inv-text-muted mt-1">{description}</p>}
    </div>
  );
}

export function StatCard({
  label,
  value,
  sub,
  color,
  icon,
  onClick,
}: {
  label: string;
  value: number;
  sub: string;
  color: string;
  icon: ReactNode;
  onClick?: () => void;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className="inv-card inv-card-hover p-5 text-left w-full cursor-pointer"
      style={{ borderColor: `${color}30` }}
    >
      <div className="flex items-center justify-between mb-3">
        <span className="text-[11px] font-semibold text-inv-text-muted uppercase tracking-wider">{label}</span>
        <div
          className="w-9 h-9 rounded-lg flex items-center justify-center"
          style={{ background: `${color}18`, color }}
        >
          {icon}
        </div>
      </div>
      <div className="text-2xl sm:text-3xl font-bold text-inv-text mb-0.5 tabular-nums">{value}</div>
      <div className="text-xs text-inv-text-muted">{sub}</div>
    </button>
  );
}

export function NavItem({
  active,
  onClick,
  icon,
  label,
  badge,
  shortcut,
}: {
  active: boolean;
  onClick: () => void;
  icon: ReactNode;
  label: string;
  badge?: number;
  shortcut?: string;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-all ${
        active
          ? "bg-inv-accent/15 text-inv-accent-light border border-inv-accent/25"
          : "text-inv-text-secondary hover:bg-inv-elevated hover:text-inv-text border border-transparent"
      }`}
    >
      <span className="shrink-0 opacity-80">{icon}</span>
      <span className="flex-1 text-left">{label}</span>
      {badge !== undefined && badge > 0 && (
        <span className="bg-red-500/90 text-white text-[10px] font-semibold rounded-full px-1.5 py-0.5 min-w-[1.25rem] text-center">
          {badge}
        </span>
      )}
      {shortcut && <span className="text-[10px] text-inv-text-muted hidden lg:inline">{shortcut}</span>}
    </button>
  );
}

export function CategoryBadge({
  categoria,
  emoji,
  color,
  bg,
  border,
}: {
  categoria: string;
  emoji: string;
  color: string;
  bg: string;
  border: string;
}) {
  return (
    <span
      className="inline-flex items-center gap-1 text-[11px] px-2 py-0.5 rounded-md font-medium"
      style={{ background: bg, color, border: `1px solid ${border}` }}
    >
      <span className="opacity-90">{emoji}</span>
      {categoria}
    </span>
  );
}

export function StockStatusBadge({
  label,
  color,
  bg,
}: {
  label: string;
  color: string;
  bg: string;
}) {
  return (
    <span
      className="inline-flex items-center gap-1.5 text-[11px] px-2.5 py-1 rounded-full font-semibold"
      style={{ background: bg, color, border: `1px solid ${color}40` }}
    >
      <span className="w-1.5 h-1.5 rounded-full" style={{ background: color }} />
      {label}
    </span>
  );
}

export function EmptyState({ message }: { message: string }) {
  return (
    <div className="text-center py-12 text-inv-text-muted text-sm">{message}</div>
  );
}

export function FieldLabel({ children }: { children: ReactNode }) {
  return <label className="block text-sm font-medium text-inv-text-secondary mb-1.5">{children}</label>;
}

export function Pagination({
  pagina,
  total,
  onPrev,
  onNext,
}: {
  pagina: number;
  total: number;
  onPrev: () => void;
  onNext: () => void;
}) {
  if (total <= 1) return null;
  return (
    <div className="flex justify-center items-center gap-3 p-4 border-t border-inv-border-subtle">
      <button type="button" onClick={onPrev} disabled={pagina === 1} className="inv-btn-ghost px-4 py-2 text-sm disabled:opacity-40">
        Anterior
      </button>
      <span className="text-sm text-inv-text-muted tabular-nums">
        {pagina} / {total}
      </span>
      <button type="button" onClick={onNext} disabled={pagina === total} className="inv-btn-ghost px-4 py-2 text-sm disabled:opacity-40">
        Siguiente
      </button>
    </div>
  );
}
