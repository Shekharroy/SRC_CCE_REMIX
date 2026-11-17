import { ArrowDown, ArrowUp, ArrowUpDown } from 'lucide-react';

export type SortDirection = 'asc' | 'desc';

export interface SortState<K extends string = string> {
  key: K;
  direction: SortDirection;
}

export function toggleSortState<K extends string>(
  prev: SortState<K> | null,
  key: K,
  defaultDirection: SortDirection = 'asc'
): SortState<K> | null {
  if (!prev || prev.key !== key) {
    return { key, direction: defaultDirection };
  }
  if (prev.direction === defaultDirection) {
    return { key, direction: defaultDirection === 'asc' ? 'desc' : 'asc' };
  }
  return null;
}

export function sortItems<T, K extends string>(
  items: readonly T[],
  sort: SortState<K> | null,
  accessors: Record<K, (item: T) => string | number | boolean | null | undefined>
): T[] {
  if (!sort) return [...items];
  const getter = accessors[sort.key];
  if (!getter) return [...items];

  const dir = sort.direction === 'asc' ? 1 : -1;
  return [...items].sort((a, b) => {
    const va = getter(a);
    const vb = getter(b);

    if (va === vb) return 0;
    if (va === null || va === undefined || va === '') return 1;
    if (vb === null || vb === undefined || vb === '') return -1;

    if (typeof va === 'number' && typeof vb === 'number') {
      return (va - vb) * dir;
    }
    if (typeof va === 'boolean' && typeof vb === 'boolean') {
      return (Number(va) - Number(vb)) * dir;
    }
    return String(va).localeCompare(String(vb), undefined, {
      numeric: true,
      sensitivity: 'base',
    }) * dir;
  });
}

interface SortableHeaderProps<K extends string> {
  label: string;
  sortKey: K;
  sort: SortState<K> | null;
  onSort: (key: K, defaultDir?: SortDirection) => void;
  align?: 'left' | 'right';
  defaultDirection?: SortDirection;
  className?: string;
}

export function SortableHeader<K extends string>({
  label,
  sortKey,
  sort,
  onSort,
  align = 'left',
  defaultDirection = 'asc',
  className = '',
}: SortableHeaderProps<K>) {
  const isActive = sort?.key === sortKey;

  return (
    <th
      className={`py-2.5 px-3 select-none ${
        align === 'right' ? 'text-right' : 'text-left'
      } ${className}`}
      aria-sort={
        isActive
          ? sort.direction === 'asc'
            ? 'ascending'
            : 'descending'
          : 'none'
      }
    >
      <button
        type="button"
        onClick={() => onSort(sortKey, defaultDirection)}
        className={`inline-flex items-center gap-1 font-medium transition-colors whitespace-nowrap cursor-pointer focus:outline-none focus-visible:ring-2 focus-visible:ring-blue-600 rounded-xs ${
          align === 'right' ? 'flex-row-reverse' : ''
        } ${
          isActive
            ? 'text-blue-700'
            : 'text-slate-600 hover:text-slate-900'
        }`}
        title={`Sort by ${label}`}
      >
        <span>{label}</span>
        {isActive ? (
          sort.direction === 'asc' ? (
            <ArrowUp className="w-3 h-3 shrink-0 text-blue-700" />
          ) : (
            <ArrowDown className="w-3 h-3 shrink-0 text-blue-700" />
          )
        ) : (
          <ArrowUpDown className="w-3 h-3 shrink-0 text-slate-400 opacity-75" />
        )}
      </button>
    </th>
  );
}
