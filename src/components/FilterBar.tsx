'use client';
import { useCallback } from 'react';

interface Props {
  label: string;
  options: string[];
  selected: Set<string>;
  onChange: (s: Set<string>) => void;
}

export default function FilterBar({ label, options, selected, onChange }: Props) {
  const toggle = useCallback((v: string) => {
    const next = new Set(selected);
    if (next.has(v)) next.delete(v); else next.add(v);
    onChange(next);
  }, [selected, onChange]);

  const all = useCallback(() => onChange(new Set(options)), [options, onChange]);
  const none = useCallback(() => onChange(new Set()), [onChange]);

  return (
    <div className="min-w-[180px]">
      <div className="flex items-center gap-2 mb-1">
        <span className="text-[10px] font-semibold text-gray-500 uppercase tracking-wide">{label}</span>
        <button onClick={all} className="text-[9px] px-2 py-0.5 border rounded-full hover:bg-gray-100">All</button>
        <button onClick={none} className="text-[9px] px-2 py-0.5 border rounded-full hover:bg-gray-100">None</button>
      </div>
      <div className="flex flex-wrap gap-1 max-h-[80px] overflow-y-auto scrollbar-thin">
        {options.map(v => (
          <button
            key={v}
            onClick={() => toggle(v)}
            className={`text-[10px] px-2 py-1 rounded-full border transition-all ${
              selected.has(v) ? 'bg-brand-900 text-white border-brand-900' : 'bg-white text-gray-600 border-gray-300 hover:border-brand-500'
            }`}
          >
            {v}
          </button>
        ))}
      </div>
    </div>
  );
}
