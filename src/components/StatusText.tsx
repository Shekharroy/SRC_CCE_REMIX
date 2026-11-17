import { CheckCircle2, Clock, Flame, Snowflake, Sun, XCircle, CircleDot, BookmarkCheck } from 'lucide-react';
import type { NormalizedStatus } from '../shared/types';

interface StatusTextProps {
  status: NormalizedStatus;
}

export function StatusText({ status }: StatusTextProps) {
  switch (status) {
    case 'Converted':
      return (
        <span className="inline-flex items-center gap-1.5 text-xs font-medium text-emerald-700 whitespace-nowrap">
          <CheckCircle2 className="w-3.5 h-3.5 shrink-0" />
          <span>Converted</span>
        </span>
      );
    case 'Booked':
      return (
        <span className="inline-flex items-center gap-1.5 text-xs font-medium text-emerald-700 whitespace-nowrap">
          <BookmarkCheck className="w-3.5 h-3.5 shrink-0" />
          <span>Booked</span>
        </span>
      );
    case 'Hot':
      return (
        <span className="inline-flex items-center gap-1.5 text-xs font-medium text-amber-700 whitespace-nowrap">
          <Flame className="w-3.5 h-3.5 shrink-0" />
          <span>Hot</span>
        </span>
      );
    case 'Follow-up':
      return (
        <span className="inline-flex items-center gap-1.5 text-xs font-medium text-blue-700 whitespace-nowrap">
          <Clock className="w-3.5 h-3.5 shrink-0" />
          <span>Follow-up</span>
        </span>
      );
    case 'Warm':
      return (
        <span className="inline-flex items-center gap-1.5 text-xs font-medium text-amber-600 whitespace-nowrap">
          <Sun className="w-3.5 h-3.5 shrink-0" />
          <span>Warm</span>
        </span>
      );
    case 'Cold':
      return (
        <span className="inline-flex items-center gap-1.5 text-xs font-medium text-slate-500 whitespace-nowrap">
          <Snowflake className="w-3.5 h-3.5 shrink-0" />
          <span>Cold</span>
        </span>
      );
    case 'Lost':
      return (
        <span className="inline-flex items-center gap-1.5 text-xs font-medium text-red-700 whitespace-nowrap">
          <XCircle className="w-3.5 h-3.5 shrink-0" />
          <span>Lost</span>
        </span>
      );
    default:
      return (
        <span className="inline-flex items-center gap-1.5 text-xs font-medium text-slate-700 whitespace-nowrap">
          <CircleDot className="w-3.5 h-3.5 shrink-0" />
          <span>Open</span>
        </span>
      );
  }
}
