import { useState, useRef, useEffect } from "react";
import { ChevronDownIcon, XIcon, SearchIcon } from "lucide-react";
import { cn } from "@/lib/utils";

function useDropdownPosition(triggerRef: React.RefObject<HTMLElement | null>, open: boolean) {
  const [style, setStyle] = useState<React.CSSProperties>({});
  useEffect(() => {
    if (!open || !triggerRef.current) return;
    const rect = triggerRef.current.getBoundingClientRect();
    const viewportW = window.innerWidth;
    const left = Math.max(8, Math.min(rect.left, viewportW - rect.width - 8));
    setStyle({
      position: "fixed",
      top: rect.bottom + 4,
      left,
      width: rect.width,
      zIndex: 9999,
    });
  }, [open, triggerRef]);
  return style;
}

interface MultiSelectSearchProps {
  options: { value: string; label: string }[];
  selected: string[];
  onChange: (selected: string[]) => void;
  placeholder?: string;
  searchPlaceholder?: string;
  className?: string;
}

function MultiSelectSearch({
  options,
  selected,
  onChange,
  placeholder = "Select...",
  searchPlaceholder = "Search...",
  className,
}: MultiSelectSearchProps) {
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState("");
  const containerRef = useRef<HTMLDivElement>(null);
  const triggerRef = useRef<HTMLButtonElement>(null);
  const searchInputRef = useRef<HTMLInputElement>(null);
  const dropdownStyle = useDropdownPosition(triggerRef, open);

  useEffect(() => {
    function handleClickOutside(e: MouseEvent) {
      if (
        containerRef.current &&
        !containerRef.current.contains(e.target as Node)
      ) {
        setOpen(false);
        setSearch("");
      }
    }
    document.addEventListener("mousedown", handleClickOutside);
    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, []);

  useEffect(() => {
    if (open && searchInputRef.current) {
      searchInputRef.current.focus();
    }
  }, [open]);

  const filtered = options.filter(
    (o) =>
      o.value.toLowerCase().includes(search.toLowerCase()) ||
      o.label.toLowerCase().includes(search.toLowerCase()),
  );

  const toggle = (value: string) => {
    onChange(
      selected.includes(value)
        ? selected.filter((s) => s !== value)
        : [...selected, value],
    );
  };

  const removeTag = (value: string, e: React.MouseEvent) => {
    e.stopPropagation();
    onChange(selected.filter((s) => s !== value));
  };

  return (
    <div ref={containerRef} className={cn("relative w-full", className)}>
      <button
        ref={triggerRef}
        type="button"
        onClick={() => setOpen(!open)}
        className={cn(
          "flex w-full min-h-9 items-center justify-between gap-2 rounded-md border border-input bg-transparent px-3 py-1.5 text-sm shadow-xs transition-[color,box-shadow] outline-none",
          "focus-visible:border-ring focus-visible:ring-[3px] focus-visible:ring-ring/50",
          "dark:bg-input/30 dark:hover:bg-input/50",
        )}
      >
        <div className="flex flex-1 flex-wrap gap-1 items-center">
          {selected.length === 0 && (
            <span className="text-muted-foreground">{placeholder}</span>
          )}
          {selected.map((val) => {
            const opt = options.find((o) => o.value === val);
            return (
              <span
                key={val}
                className="inline-flex items-center gap-1 rounded-md bg-gray-100 px-2 py-0.5 text-xs font-medium text-gray-700 dark:bg-gray-800 dark:text-gray-300"
              >
                {opt?.value ?? val}
                <XIcon
                  className="size-3 cursor-pointer hover:text-gray-900 dark:hover:text-gray-100"
                  onClick={(e) => removeTag(val, e)}
                />
              </span>
            );
          })}
        </div>
        <ChevronDownIcon
          className={cn(
            "size-4 shrink-0 opacity-50 transition-transform",
            open && "rotate-180",
          )}
        />
      </button>

      {open && (
        <div style={dropdownStyle} className="rounded-md border bg-popover text-popover-foreground shadow-md">
          <div className="flex items-center gap-2 border-b px-3 py-2">
            <SearchIcon className="size-4 text-muted-foreground shrink-0" />
            <input
              ref={searchInputRef}
              type="text"
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder={searchPlaceholder}
              style={{ fontSize: "16px" }}
              className="flex-1 bg-transparent outline-none placeholder:text-muted-foreground"
            />
          </div>
          <div className="max-h-52 overflow-y-auto p-1">
            {filtered.length === 0 && (
              <p className="py-4 text-center text-sm text-muted-foreground">
                No results found
              </p>
            )}
            {filtered.map((option) => {
              const isSelected = selected.includes(option.value);
              return (
                <button
                  key={option.value}
                  type="button"
                  onClick={() => toggle(option.value)}
                  className={cn(
                    "flex w-full items-center gap-2 rounded-sm px-2 py-1.5 text-sm outline-none select-none",
                    "hover:bg-accent hover:text-accent-foreground",
                    isSelected && "bg-accent/50",
                  )}
                >
                  <span
                    className={cn(
                      "flex size-4 shrink-0 items-center justify-center rounded-sm border",
                      isSelected
                        ? "border-primary bg-primary text-primary-foreground"
                        : "border-input",
                    )}
                  >
                    {isSelected && (
                      <svg
                        width="10"
                        height="10"
                        viewBox="0 0 10 10"
                        fill="none"
                      >
                        <path
                          d="M8.5 2.5L3.8 7.5L1.5 5"
                          stroke="currentColor"
                          strokeWidth="1.5"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        />
                      </svg>
                    )}
                  </span>
                  <span className="font-medium">{option.value}</span>
                  <span className="text-muted-foreground">{option.label}</span>
                </button>
              );
            })}
          </div>
          {selected.length > 0 && (
            <div className="border-t px-3 py-2">
              <button
                type="button"
                onClick={() => onChange([])}
                className="text-xs text-muted-foreground hover:text-foreground"
              >
                Clear all
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

export { MultiSelectSearch };
