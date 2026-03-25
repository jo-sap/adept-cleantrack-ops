import React, { useCallback, useEffect, useId, useLayoutEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import { Check, ChevronDown } from "lucide-react";

export type AppSelectOption = {
  value: string;
  label: string;
  disabled?: boolean;
};

export type AppSelectProps = {
  label?: React.ReactNode;
  value: string;
  onChange: (value: string) => void;
  options: AppSelectOption[];
  placeholder?: string;
  disabled?: boolean;
  error?: string;
  helperText?: string;
  fullWidth?: boolean;
  className?: string;
  /** Extra classes on the trigger button only */
  triggerClassName?: string;
  id?: string;
  name?: string;
  required?: boolean;
  size?: "md" | "sm";
  "aria-describedby"?: string;
};

function nextEnabledIndex(options: AppSelectOption[], from: number, delta: 1 | -1): number {
  const n = options.length;
  if (n === 0) return 0;
  let i = from;
  for (let step = 0; step < n; step++) {
    i = (i + delta + n) % n;
    if (!options[i]?.disabled) return i;
  }
  return from;
}

/**
 * Custom listbox select aligned with app tokens (so-btn-secondary / so-input family).
 * Uses a portal for the menu to avoid overflow clipping in modals and scroll regions.
 */
export const AppSelect: React.FC<AppSelectProps> = ({
  label,
  value,
  onChange,
  options,
  placeholder = "Select…",
  disabled = false,
  error,
  helperText,
  fullWidth = true,
  className = "",
  triggerClassName = "",
  id: idProp,
  name,
  required,
  size = "md",
  "aria-describedby": ariaDescribedBy,
}) => {
  const reactId = useId();
  const baseId = idProp ?? `app-select-${reactId.replace(/:/g, "")}`;
  const listboxId = `${baseId}-listbox`;
  const labelId = `${baseId}-label`;
  const helperId = `${baseId}-helper`;
  const errorId = `${baseId}-error`;

  const [open, setOpen] = useState(false);
  const [highlightedIndex, setHighlightedIndex] = useState(0);
  const [menuRect, setMenuRect] = useState<{ top: number; left: number; width: number; maxHeight: number } | null>(null);

  const rootRef = useRef<HTMLDivElement>(null);
  const triggerRef = useRef<HTMLButtonElement>(null);
  const listRef = useRef<HTMLDivElement>(null);

  const selected = useMemo(() => options.find((o) => o.value === value), [options, value]);
  const displayLabel = selected?.label ?? (value ? value : null);

  const updateMenuPosition = useCallback(() => {
    const el = triggerRef.current;
    if (!el) return;
    const r = el.getBoundingClientRect();
    const gap = 4;
    const below = r.bottom + gap;
    const spaceBelow = window.innerHeight - below - 12;
    const spaceAbove = r.top - 12;
    const maxH = Math.min(240, Math.max(120, Math.max(spaceBelow, spaceAbove) - 8));
    let top = below;
    if (spaceBelow < 160 && spaceAbove > spaceBelow) {
      top = Math.max(8, r.top - gap - maxH);
    }
    setMenuRect({
      top,
      left: r.left,
      width: r.width,
      maxHeight: maxH,
    });
  }, []);

  useLayoutEffect(() => {
    if (!open) {
      setMenuRect(null);
      return;
    }
    updateMenuPosition();
    const ro = new ResizeObserver(() => updateMenuPosition());
    if (triggerRef.current) ro.observe(triggerRef.current);
    window.addEventListener("scroll", updateMenuPosition, true);
    window.addEventListener("resize", updateMenuPosition);
    return () => {
      ro.disconnect();
      window.removeEventListener("scroll", updateMenuPosition, true);
      window.removeEventListener("resize", updateMenuPosition);
    };
  }, [open, updateMenuPosition]);

  useEffect(() => {
    if (!open) return;
    const onDoc = (e: MouseEvent) => {
      const t = e.target as Node;
      if (rootRef.current?.contains(t)) return;
      if (listRef.current?.contains(t)) return;
      setOpen(false);
    };
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [open]);

  const selectIndex = useCallback(
    (idx: number) => {
      const opt = options[idx];
      if (!opt || opt.disabled) return;
      onChange(opt.value);
      setOpen(false);
      triggerRef.current?.focus();
    },
    [onChange, options]
  );

  const openMenu = useCallback(() => {
    if (disabled) return;
    const selIdx = Math.max(
      0,
      options.findIndex((o) => o.value === value && !o.disabled)
    );
    const start = options[selIdx] && !options[selIdx].disabled ? selIdx : nextEnabledIndex(options, 0, 1);
    setHighlightedIndex(start);
    setOpen(true);
  }, [disabled, options, value]);

  const onTriggerKeyDown = (e: React.KeyboardEvent) => {
    if (disabled) return;
    if (e.key === "ArrowDown" || e.key === "ArrowUp" || e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      if (!open) {
        openMenu();
        return;
      }
    }
    if (!open) return;

    if (e.key === "Escape") {
      e.preventDefault();
      setOpen(false);
      triggerRef.current?.focus();
      return;
    }
    if (e.key === "ArrowDown") {
      e.preventDefault();
      setHighlightedIndex((i) => nextEnabledIndex(options, i, 1));
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      setHighlightedIndex((i) => nextEnabledIndex(options, i, -1));
    } else if (e.key === "Home") {
      e.preventDefault();
      setHighlightedIndex(nextEnabledIndex(options, 0, 1));
    } else if (e.key === "End") {
      e.preventDefault();
      setHighlightedIndex(nextEnabledIndex(options, options.length - 1, -1));
    } else if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      selectIndex(highlightedIndex);
    }
  };

  useEffect(() => {
    if (!open || !listRef.current) return;
    const el = listRef.current.querySelector(`[data-option-index="${highlightedIndex}"]`) as HTMLElement | null;
    el?.scrollIntoView({ block: "nearest" });
  }, [open, highlightedIndex]);

  const describedBy = [ariaDescribedBy, error ? errorId : null, helperText ? helperId : null].filter(Boolean).join(" ") || undefined;

  const sizeClasses =
    size === "sm"
      ? "min-h-[34px] px-2 py-1.5 text-xs"
      : "min-h-[42px] px-3 py-2.5 text-sm";

  const triggerState = error
    ? "border-red-200 bg-white shadow-sm focus:ring-2 focus:ring-red-200 focus:border-red-300"
    : "border-[var(--so-border-subtle)] bg-white shadow-[0_1px_2px_rgba(15,23,42,0.04)] hover:border-[var(--so-border-strong)] hover:bg-[#FAFBFC] focus:outline-none focus:border-[var(--so-accent)] focus:ring-2 focus:ring-[rgba(62,95,106,0.2)]";

  return (
    <div
      ref={rootRef}
      className={`${fullWidth ? "w-full" : ""} ${className}`.trim()}
    >
      {name ? (
        <input type="hidden" name={name} value={value} readOnly aria-hidden />
      ) : null}
      {label != null ? (
        <label
          id={labelId}
          htmlFor={baseId}
          className="mb-1 block text-[10px] font-bold uppercase tracking-widest text-gray-400"
        >
          {label}
          {required ? <span className="text-red-500"> *</span> : null}
        </label>
      ) : null}

      <button
        ref={triggerRef}
        type="button"
        id={baseId}
        disabled={disabled}
        aria-haspopup="listbox"
        aria-expanded={open}
        aria-controls={open ? listboxId : undefined}
        aria-describedby={describedBy}
        aria-invalid={!!error}
        aria-required={required}
        onClick={() => (open ? setOpen(false) : openMenu())}
        onKeyDown={onTriggerKeyDown}
        className={`
          inline-flex w-full items-center justify-between gap-2 rounded-[var(--so-radius-input)] border text-left
          transition-[border-color,box-shadow,background-color] duration-150
          ${sizeClasses}
          ${triggerState}
          ${disabled ? "cursor-not-allowed opacity-50" : "cursor-pointer"}
          ${open ? "border-[var(--so-accent)] ring-2 ring-[rgba(62,95,106,0.2)]" : ""}
          ${triggerClassName}
        `.trim().replace(/\s+/g, " ")}
      >
        <span className={`min-w-0 flex-1 truncate ${!displayLabel ? "text-gray-400" : "text-gray-900 font-medium"}`}>
          {displayLabel ?? placeholder}
        </span>
        <ChevronDown
          size={size === "sm" ? 14 : 16}
          className={`shrink-0 text-gray-400 transition-transform duration-200 ${open ? "rotate-180" : ""}`}
          aria-hidden
        />
      </button>

      {helperText && !error ? (
        <p id={helperId} className="mt-1 text-[11px] text-gray-500">
          {helperText}
        </p>
      ) : null}
      {error ? (
        <p id={errorId} className="mt-1 text-[11px] text-red-600" role="alert">
          {error}
        </p>
      ) : null}

      {open && menuRect && typeof document !== "undefined"
        ? createPortal(
            <div
              ref={listRef}
              id={listboxId}
              role="listbox"
              aria-labelledby={label != null ? labelId : undefined}
              aria-activedescendant={`${baseId}-opt-${highlightedIndex}`}
              className="fixed z-[300] overflow-auto rounded-[var(--so-radius-input)] border border-[var(--so-border-subtle)] bg-[var(--so-page-surface)] py-1 shadow-lg"
              style={{
                top: menuRect.top,
                left: menuRect.left,
                width: menuRect.width,
                maxHeight: menuRect.maxHeight,
              }}
            >
              {options.map((opt, idx) => {
                const isSelected = opt.value === value;
                const isHi = idx === highlightedIndex;
                return (
                  <div
                    key={`${opt.value}-${idx}`}
                    id={`${baseId}-opt-${idx}`}
                    role="option"
                    aria-selected={isSelected}
                    data-option-index={idx}
                    className={`
                      flex cursor-pointer items-center justify-between gap-2 px-3 py-2 text-sm
                      ${opt.disabled ? "cursor-not-allowed opacity-40" : ""}
                      ${isHi && !opt.disabled ? "bg-[var(--so-accent-soft)]" : ""}
                      ${isSelected ? "font-semibold text-gray-900" : "font-normal text-gray-800"}
                    `.trim().replace(/\s+/g, " ")}
                    onMouseEnter={() => !opt.disabled && setHighlightedIndex(idx)}
                    onMouseDown={(e) => e.preventDefault()}
                    onClick={() => !opt.disabled && selectIndex(idx)}
                  >
                    <span className="min-w-0 truncate">{opt.label}</span>
                    {isSelected ? <Check size={14} className="shrink-0 text-[var(--so-accent)]" aria-hidden /> : null}
                  </div>
                );
              })}
            </div>,
            document.body
          )
        : null}
    </div>
  );
};

export default AppSelect;
