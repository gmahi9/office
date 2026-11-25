import type { PointerEvent as ReactPointerEvent } from 'react';
import { CHART_COLORS, readChartData } from './chartData';
import type { ChartModel } from '../../domain/types';
import type { WorkbookSession } from '../../engine/session';

export function ChartCard({
  chart,
  session,
  sheetId,
  originX,
  originY,
  onMove,
  onRemove,
}: {
  chart: ChartModel;
  session: WorkbookSession;
  sheetId: string;
  originX: number;
  originY: number;
  onMove: (id: string, x: number, y: number) => void;
  onRemove: (id: string) => void;
}) {
  const data = readChartData(session, sheetId, chart.range);
  function onPointerDown(event: ReactPointerEvent<HTMLElement>) {
    event.stopPropagation();
    const startX = event.clientX;
    const startY = event.clientY;
    const baseX = chart.x;
    const baseY = chart.y;
    function move(next: globalThis.PointerEvent) {
      onMove(chart.id, Math.max(0, baseX + next.clientX - startX), Math.max(0, baseY + next.clientY - startY));
    }
    function up() {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
    }
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  }
  return (
    <figure
      className="chart-card"
      style={{ left: chart.x - originX, top: chart.y - originY, width: chart.w, height: chart.h }}
      onPointerDown={(event) => event.stopPropagation()}
    >
      <header onPointerDown={onPointerDown}>
        <strong>{chart.title}</strong>
        <button type="button" onClick={() => onRemove(chart.id)} aria-label={`Remove ${chart.title}`}>
          Remove
        </button>
      </header>
      {data.series.length === 0 ? (
        <p className="chart-empty">This range has no numbers yet.</p>
      ) : chart.type === 'pie' ? (
        <Pie data={data} />
      ) : chart.type === 'line' ? (
        <Line data={data} />
      ) : chart.type === 'scatter' ? (
        <Scatter data={data} />
      ) : (
        <Bars data={data} />
      )}
    </figure>
  );
}

function Bars({ data }: { data: ReturnType<typeof readChartData> }) {
  const width = 360;
  const height = 160;
  const max = Math.max(1, ...data.series.flatMap((series) => series.values.map((value) => value ?? 0)));
  const group = width / Math.max(1, data.categories.length);
  return (
    <svg viewBox={`0 0 ${width} ${height + 22}`} role="img" aria-label="Bar chart">
      {data.categories.map((label, index) => (
        <text key={label + index} x={group * index + group / 2} y={height + 14} textAnchor="middle" className="chart-label">
          {(label.split(' ')[0] ?? label).slice(0, 12)}
        </text>
      ))}
      {data.series.map((series, s) =>
        series.values.map((value, index) => {
          const h = ((value ?? 0) / max) * (height - 8);
          const barW = Math.max(4, (group - 10) / data.series.length);
          const x = group * index + 6 + s * barW;
          return <rect key={`${s}-${index}`} x={x} y={height - h} width={barW - 2} height={h} fill={CHART_COLORS[s % CHART_COLORS.length]} rx="2" />;
        }),
      )}
    </svg>
  );
}

function Line({ data }: { data: ReturnType<typeof readChartData> }) {
  const width = 360;
  const height = 160;
  const max = Math.max(1, ...data.series.flatMap((series) => series.values.map((value) => value ?? 0)));
  const step = width / Math.max(1, data.categories.length - 1);
  return (
    <svg viewBox={`0 0 ${width} ${height}`} role="img" aria-label="Line chart">
      {data.series.map((series, s) => {
        const points = series.values
          .map((value, index) => `${index * step},${height - 8 - ((value ?? 0) / max) * (height - 16)}`)
          .join(' ');
        return <polyline key={series.name} points={points} fill="none" stroke={CHART_COLORS[s % CHART_COLORS.length]} strokeWidth="2" />;
      })}
    </svg>
  );
}

function Pie({ data }: { data: ReturnType<typeof readChartData> }) {
  const series = data.series[0];
  const values = series?.values.map((value) => Math.max(0, value ?? 0)) ?? [];
  const total = values.reduce((sum, value) => sum + value, 0) || 1;
  let angle = -Math.PI / 2;
  const arcs = values.map((value, index) => {
    const slice = (value / total) * Math.PI * 2;
    const start = angle;
    angle += slice;
    const large = slice > Math.PI ? 1 : 0;
    const r = 58;
    const x1 = 80 + r * Math.cos(start);
    const y1 = 80 + r * Math.sin(start);
    const x2 = 80 + r * Math.cos(start + slice);
    const y2 = 80 + r * Math.sin(start + slice);
    const d = `M 80 80 L ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2} Z`;
    return <path key={index} d={d} fill={CHART_COLORS[index % CHART_COLORS.length]} />;
  });
  return (
    <svg viewBox="0 0 160 160" role="img" aria-label="Pie chart">
      {arcs}
    </svg>
  );
}

function Scatter({ data }: { data: ReturnType<typeof readChartData> }) {
  const series = data.series[0];
  const ys = series?.values ?? [];
  const maxX = Math.max(1, ...data.xs);
  const maxY = Math.max(1, ...ys.map((value) => value ?? 0));
  return (
    <svg viewBox="0 0 360 160" role="img" aria-label="Scatter chart">
      {ys.map((value, index) =>
        value == null ? null : (
          <circle
            key={index}
            cx={(data.xs[index] / maxX) * 340 + 8}
            cy={150 - (value / maxY) * 140}
            r="4"
            fill={CHART_COLORS[0]}
          />
        ),
      )}
    </svg>
  );
}
