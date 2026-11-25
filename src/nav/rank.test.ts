import { beforeEach, describe, expect, it } from 'vitest';
import { resetOfficeDb } from '../data/db';
import { readUsage, recordUse, resetUsage } from '../data/usage';
import { resetDeskCache } from '../data/workbooks';
import { COMMANDS } from './commands';
import { rankCommands } from './rank';

const now = 1_700_000_000_000;

describe('adaptive navigation', () => {
  beforeEach(async () => {
    resetDeskCache();
    await resetOfficeDb();
  });

  it('moves frequently used tools forward and keeps a stable base order', () => {
    const tools = COMMANDS.filter((command) => ['bold', 'italic', 'merge'].includes(command.id));
    const ranked = rankCommands(tools, [
      { commandId: 'merge', count: 10, lastUsed: now },
      { commandId: 'bold', count: 2, lastUsed: now },
    ], now);
    expect(ranked.map((command) => command.id)).toEqual(['merge', 'bold', 'italic']);
  });

  it('lets a recent tool pass an old one', () => {
    const tools = COMMANDS.filter((command) => ['bold', 'merge'].includes(command.id));
    const ranked = rankCommands(tools, [
      { commandId: 'merge', count: 12, lastUsed: now - 100 * 24 * 60 * 60 * 1000 },
      { commandId: 'bold', count: 4, lastUsed: now },
    ], now);
    expect(ranked[0]?.id).toBe('bold');
  });

  it('stores usage on this device and can forget it', async () => {
    await recordUse('chart', now);
    await recordUse('chart', now + 10);
    const usage = await readUsage();
    expect(usage).toEqual([{ commandId: 'chart', count: 2, lastUsed: now + 10 }]);
    const ranked = rankCommands(COMMANDS, usage, now + 10);
    expect(ranked[0]?.id).toBe('chart');
    await resetUsage();
    expect(await readUsage()).toEqual([]);
    expect(rankCommands(COMMANDS, [], now).map((command) => command.id)).toEqual(COMMANDS.map((command) => command.id));
  });
});
