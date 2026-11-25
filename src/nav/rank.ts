import type { UsageStat } from '../data/usage';

/** Two weeks. A tool you used last month sinks behind one you used today. */
export const USAGE_HALF_LIFE_MS = 14 * 24 * 60 * 60 * 1000;

export function usageScore(stat: UsageStat | undefined, now: number): number {
  if (!stat || stat.count <= 0) return 0;
  const age = Math.max(0, now - stat.lastUsed);
  return stat.count * Math.pow(0.5, age / USAGE_HALF_LIFE_MS);
}

export function rankCommands<T extends { id: string; baseOrder: number }>(
  commands: readonly T[],
  usage: readonly UsageStat[],
  now: number,
): T[] {
  const byId = new Map(usage.map((stat) => [stat.commandId, stat]));
  return [...commands].sort((a, b) => {
    const delta = usageScore(byId.get(b.id), now) - usageScore(byId.get(a.id), now);
    if (Math.abs(delta) > 1e-9) return delta;
    return a.baseOrder - b.baseOrder;
  });
}
