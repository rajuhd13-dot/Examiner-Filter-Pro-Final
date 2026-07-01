import React, { useState, useMemo, useRef, useEffect } from "react";
import { Search, ChevronDown, Check, X } from "lucide-react";
import { motion, AnimatePresence } from "framer-motion";
import { List } from "react-window";
import { cn } from "../lib/utils";

interface Option {
  label: string;
  value: string;
}

interface FilterDropdownProps {
  label: string;
  placeholder: string;
  options: (string | Option)[];
  selected: string[];
  onChange: (selected: string[]) => void;
  emptyMeansAll?: boolean;
}

export const FilterDropdown: React.FC<FilterDropdownProps> = ({
  label,
  placeholder,
  options,
  selected,
  onChange,
  emptyMeansAll = false
}) => {
  const [isOpen, setIsOpen] = useState(false);
  const [search, setSearch] = useState("");
  const dropdownRef = useRef<HTMLDivElement>(null);

  const formattedOptions = useMemo(() => {
    if (!Array.isArray(options)) return [];
    return options.map(opt => {
      if (typeof opt === "string") {
        return { label: opt, value: opt === "(Blank)" ? "__BLANK__" : opt };
      }
      return opt;
    });
  }, [options]);

  const filteredOptions = useMemo(() => {
    if (!search.trim()) return formattedOptions;
    const lower = search.toLowerCase();
    return formattedOptions.filter(opt => opt.label.toLowerCase().includes(lower));
  }, [formattedOptions, search]);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const toggleOption = (value: string) => {
    const newSelected = selected.includes(value)
      ? selected.filter(s => s !== value)
      : [...selected, value];
    onChange(newSelected);
  };

  const selectAll = () => {
    onChange(formattedOptions.map(o => o.value));
  };

  const clearAll = () => {
    onChange([]);
  };

  const displayLabel = useMemo(() => {
    if (selected.length === 0) return emptyMeansAll ? "All (show all)" : placeholder;
    if (selected.length === formattedOptions.length) return "All selected";
    if (selected.length === 1) {
      const opt = formattedOptions.find(o => o.value === selected[0]);
      if (!opt) return "1 selected";
      
      // If it's a Training Date, try to parse and format it
      if (label === "Training Report" || label === "Training Date") {
        if (opt.label.match(/\d{4}-\d{2}-\d{2}/)) {
           return opt.label; // Already seems to be YYYY-MM-DD
        }
        // Try parsing if it's a long date string
        const date = new Date(opt.label);
        if (!isNaN(date.getTime())) {
          return date.toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: '2-digit' });
        }
      }
      return opt.label;
    }
    return `${selected.length} selected`;
  }, [selected, formattedOptions, placeholder, emptyMeansAll, label]);

  // Virtualized Row Component
  const Row = ({ index, style }: any) => {
    const option = filteredOptions[index];
    const isSelected = selected.includes(option.value);

    const formattedLabel = useMemo(() => {
        if (label === "Training Report" || label === "Training Date") {
            // Try parsing if it's a long date string
            const date = new Date(option.label);
            if (!isNaN(date.getTime()) && option.label.length > 15) {
                return date.toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: '2-digit' });
            }
        }
        return option.label;
    }, [option.label, label]);

    return (
      <div style={style}>
        <div
          onClick={() => toggleOption(option.value)}
          className={cn(
            "group flex items-center gap-3 px-4 py-2 rounded-xl transition-all cursor-pointer mx-2 mb-0.5",
            isSelected ? "bg-blue-50/80" : "hover:bg-gray-50"
          )}
        >
          <div className={cn(
            "flex-shrink-0 w-5 h-5 border-2 rounded-lg transition-all flex items-center justify-center",
            isSelected 
              ? "bg-blue-500 border-blue-500 text-white shadow-sm shadow-blue-200" 
              : "border-gray-200 group-hover:border-blue-300"
          )}>
            {isSelected && <Check className="w-3.5 h-3.5 stroke-[3px]" />}
          </div>
          <span className={cn(
            "text-[13px] font-bold truncate transition-colors",
            isSelected ? "text-blue-700" : "text-gray-600"
          )}>
            {formattedLabel}
          </span>
        </div>
      </div>
    );
  };

  return (
    <div className="flex flex-col gap-1.5 w-full" ref={dropdownRef}>
      <label className="text-[10px] font-black text-gray-400 capitalize tracking-wider ml-1">
        {label}
      </label>
      <div className="relative">
        <button
          type="button"
          onClick={() => setIsOpen(!isOpen)}
          className={cn(
            "w-full flex items-center justify-between gap-2 px-4 py-3 bg-white border border-gray-100 rounded-2xl transition-all duration-300 text-left text-sm shadow-sm",
            isOpen ? "border-blue-400 ring-4 ring-blue-50/50" : "hover:border-blue-200 hover:shadow-md"
          )}
        >
          <span className="truncate text-gray-700 font-bold">{displayLabel}</span>
          <ChevronDown className={cn("w-4 h-4 text-gray-300 transition-transform duration-300", isOpen && "rotate-180 text-blue-500")} />
        </button>

        <AnimatePresence>
          {isOpen && (
            <motion.div
              initial={{ opacity: 0, y: 8, scale: 0.95 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, y: 8, scale: 0.95 }}
              transition={{ duration: 0.2, ease: "circOut" }}
              className="absolute z-[110] w-full mt-3 bg-white border border-gray-100 rounded-3xl shadow-[0_20px_50px_-12px_rgba(0,0,0,0.1)] overflow-hidden min-w-[280px]"
            >
              <div className="p-4 border-b border-gray-50 bg-gray-50/30">
                <div className="relative mb-3">
                  <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 w-4 h-4 text-gray-400" />
                  <input
                    type="text"
                    autoFocus
                    placeholder="Search options..."
                    value={search}
                    onChange={(e) => setSearch(e.target.value)}
                    className="w-full pl-10 pr-4 py-2.5 bg-white border border-gray-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-blue-100 focus:border-blue-400 transition-all font-medium"
                  />
                </div>
                <div className="flex gap-2">
                  <button
                    onClick={selectAll}
                    className="flex-1 py-1.5 px-2 text-[11px] font-black text-blue-600 bg-blue-50 hover:bg-blue-100 rounded-lg transition-colors uppercase tracking-wider"
                  >
                    Select All
                  </button>
                  <button
                    onClick={clearAll}
                    className="flex-1 py-1.5 px-2 text-[11px] font-black text-gray-400 bg-gray-100 hover:bg-gray-200 hover:text-gray-600 rounded-lg transition-colors uppercase tracking-wider"
                  >
                    Clear
                  </button>
                </div>
              </div>

              <div className="max-h-[300px] overflow-hidden p-0 custom-scrollbar">
                {filteredOptions.length === 0 ? (
                  <div className="px-4 py-10 text-center flex flex-col items-center gap-2">
                    <Search className="w-8 h-8 text-gray-100" />
                    <span className="text-xs text-gray-400 font-bold">No matches found</span>
                  </div>
                ) : (
                  <List
                    style={{ height: Math.min(filteredOptions.length * 44, 280) }}
                    rowCount={filteredOptions.length}
                    rowHeight={44}
                    rowComponent={Row}
                    rowProps={{} as any}
                  />
                )}
              </div>


              <div className="p-2.5 border-t border-gray-50 bg-gray-50/50 text-[9px] text-gray-400 text-center font-black uppercase tracking-[2px]">
                {filteredOptions.length} Options Found
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </div>
    </div>
  );
};


