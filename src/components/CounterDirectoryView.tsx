import { COUNTER_DEFINITIONS } from '../shared/reportEngine';
import type { CounterWiseReport } from '../shared/types';

interface CounterDirectoryViewProps {
  counterWiseReport: CounterWiseReport;
}

export function CounterDirectoryView({
  counterWiseReport,
}: CounterDirectoryViewProps) {
  const summaryByName = new Map(
    counterWiseReport.summaries.map((s) => [s.counterName, s])
  );

  return (
    <section
      id="directory"
      className="bg-white border border-slate-200 rounded-lg p-6 space-y-5"
    >
      <div className="border-b border-slate-200 pb-4">
        <h2 className="text-base font-semibold text-slate-900">
          Counter &amp; DSE Master Mapping Directory
        </h2>
        <p className="text-xs text-slate-600 mt-1">
          Canonical assignment rules used to resolve Counter Names automatically when uploaded sheets omit or abbreviate locations. Rows assigned to Web Service or Seibel Administrator automatically default Caller Name to Chandrashekhar Kumar unless another caller name is explicitly mentioned.
        </p>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
        {COUNTER_DEFINITIONS.map((counter) => {
          const liveSummary = summaryByName.get(counter.name);
          const dseMap = new Map(
            (liveSummary?.dseBreakdowns || []).map((d) => [d.dseName, d])
          );

          return (
            <div
              key={counter.id}
              className="border border-slate-200 rounded-md p-4 bg-slate-50/40 flex flex-col justify-between"
            >
              <div>
                <div className="flex items-baseline justify-between gap-2 pb-2.5 border-b border-slate-200">
                  <h3 className="text-sm font-semibold text-slate-900">
                    {counter.id}. {counter.name}
                  </h3>
                  <span className="text-xs font-mono tabular-nums text-slate-600">
                    {liveSummary?.totalEnquiries || 0} enq
                  </span>
                </div>

                <ul className="mt-3 space-y-1.5 text-xs">
                  {counter.dseNames.map((dse) => {
                    const stat = dseMap.get(dse);
                    return (
                      <li
                        key={dse}
                        className="flex items-center justify-between gap-2 text-slate-700"
                      >
                        <span>{dse}</span>
                        <span className="font-mono tabular-nums text-slate-500">
                          {stat ? `${stat.totalEnquiries} (${stat.convertedCount + stat.bookedCount} conv)` : '0'}
                        </span>
                      </li>
                    );
                  })}
                </ul>
              </div>

              <div className="mt-4 pt-2.5 border-t border-slate-200 flex items-center justify-between text-xs text-slate-500">
                <span>Conversion Rate</span>
                <span className="font-mono tabular-nums font-medium text-slate-900">
                  {liveSummary?.conversionRate || 0}%
                </span>
              </div>
            </div>
          );
        })}
      </div>
    </section>
  );
}
