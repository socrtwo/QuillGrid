import { describe, it, expect, vi, beforeAll } from 'vitest';
import type { ChatRequest, ChatResult, ChatMessage } from '../../src/ai/client';
import {
  runEnsemble, planEnsemble, parseStance, normalizeEnsemble, outcomeLine, costNote, requestCount, defaultEnsemble,
  MASTER_PROMPT, DRAFT_NOTE, STANCE_FORMAT, REVISE_RULE, type EnsemblePlan, type MemberCall, type Stance,
} from '../../src/ai/ensemble';
import { AiError } from '../../src/ai/errors';

// In-memory localStorage for the settings tests (Node has none).
const mem = new Map<string, string>();
beforeAll(() => {
  vi.stubGlobal('localStorage', {
    getItem: (k: string) => (mem.has(k) ? mem.get(k)! : null),
    setItem: (k: string, v: string) => { mem.set(k, String(v)); },
    removeItem: (k: string) => { mem.delete(k); },
    clear: () => mem.clear(),
    key: (i: number) => [...mem.keys()][i] ?? null,
    get length() { return mem.size; },
  });
});

const lastUser = (req: ChatRequest) => {
  const m = [...req.messages].reverse().find((x) => x.role === 'user');
  return typeof m?.content === 'string' ? m.content : '';
};

interface Logged { provider: string; model: string; kind: 'answer' | 'master' | 'draft' | 'stance' | 'revise'; req: ChatRequest }

/**
 * Deterministic mock providers. Each member answers "answer from <id>"; stances follow `stances`
 * (per provider, one entry per round; the last entry repeats); the clerk's minutes are numbered.
 */
function mockProviders(opts: { fail?: string[]; stances?: Record<string, Stance[]>; hang?: boolean; usage?: boolean; badJson?: string[] } = {}) {
  const log: Logged[] = [];
  let minutes = 0;
  const stanceCalls: Record<string, number> = {};
  const call: MemberCall = async (provider, model, req) => {
    const text = lastUser(req);
    const kind: Logged['kind'] = text.includes(STANCE_FORMAT) ? 'stance' : text.includes(REVISE_RULE) ? 'revise' : text.includes(DRAFT_NOTE) ? 'draft' : text.includes(MASTER_PROMPT) ? 'master' : 'answer';
    log.push({ provider, model, kind, req });
    if (opts.hang) {
      await new Promise((_, reject) => req.signal?.addEventListener('abort', () => reject(new AiError('Stopped', 'aborted'))));
    }
    if (opts.fail?.includes(provider) && kind === 'answer') throw new AiError(`${provider} is down`, 'network');
    const res = (t: string): ChatResult => {
      req.onToken?.(t, t);
      return { text: t, model, provider, providerName: provider.toUpperCase(), ...(opts.usage ? { usage: { in: 10, out: 5 } } : {}) };
    };
    switch (kind) {
      case 'answer': return res(`answer from ${provider}`);
      case 'master': return res('=SUM(A1:A3)');
      case 'draft': return res(`minute ${++minutes}`);
      case 'revise': return res(`minute ${++minutes}`);
      case 'stance': {
        if (opts.badJson?.includes(provider)) return res('I think it is fine.');
        const i = stanceCalls[provider] = (stanceCalls[provider] ?? -1) + 1;
        const list = opts.stances?.[provider] || ['unite'];
        const stance = list[Math.min(i, list.length - 1)];
        return res('```json\n' + JSON.stringify({ stance, reason: `${provider} says ${stance}`, suggested_change: stance === 'block' ? 'use SUMIFS' : '' }) + '\n```');
      }
    }
  };
  return { call, log };
}

const plan = (mode: 'master' | 'consensus', ids = ['a', 'b', 'c'], extra: Partial<EnsemblePlan> = {}): EnsemblePlan => ({
  mode, members: ids.map((p) => ({ provider: p, model: `${p}-model` })), master: { provider: 'm', model: 'm-model' }, maxRounds: 3, memberTimeoutMs: 1000, ...extra,
});
const messages: ChatMessage[] = [{ role: 'system', content: 'You write spreadsheet formulas.' }, { role: 'user', content: 'Total A1:A3. Respond with the formula only.' }];
const reverse = <T,>(xs: T[]) => xs.slice().reverse();

describe('master mode', () => {
  it('calls every member, then the master once with shuffled, labeled responses', async () => {
    const { call, log } = mockProviders({ usage: true });
    const tokens: string[] = [];
    const status: string[] = [];
    const r = await runEnsemble({ messages, temperature: 0.2, onToken: (d) => tokens.push(d), onStatus: (t) => status.push(t) }, plan('master'), call, { shuffle: reverse });
    const answers = log.filter((x) => x.kind === 'answer');
    expect(answers.map((x) => x.provider).sort()).toEqual(['a', 'b', 'c']);
    expect(answers.every((x) => x.req.temperature === 0.2 && x.req.messages === messages)).toBe(true);
    const masters = log.filter((x) => x.kind === 'master');
    expect(masters).toHaveLength(1);
    expect(masters[0].provider).toBe('m');
    // The master sees the original system prompt and messages, then the instruction and the labeled responses.
    const mm = masters[0].req.messages;
    expect(mm.slice(0, 2)).toEqual(messages);
    const prompt = lastUser(masters[0].req);
    expect(prompt.startsWith(MASTER_PROMPT)).toBe(true);
    // reversed order: Response A is c's answer, C is a's; no provider names appear
    expect(prompt.indexOf('Response A ===\nanswer from c')).toBeGreaterThan(0);
    expect(prompt.indexOf('Response C ===\nanswer from a')).toBeGreaterThan(0);
    // Only the master's answer streams.
    expect(tokens).toEqual(['=SUM(A1:A3)']);
    expect(r.text).toBe('=SUM(A1:A3)');
    expect(r.provider).toBe('m');
    expect(r.model).toBe('m-model');
    expect(r.usage).toEqual({ in: 40, out: 20 });
    expect(r.ensemble.mode).toBe('master');
    expect(r.ensemble.outcome).toBe('synthesized');
    expect(r.ensemble.members.map((m) => m.text)).toEqual(['answer from a', 'answer from b', 'answer from c']);
    expect(status[0]).toBe('Asking 3 AIs…');
    expect(outcomeLine(r.ensemble, (x) => x.toUpperCase())).toBe('Synthesized by M');
  });

  it('skips and reports a failing member', async () => {
    const { call, log } = mockProviders({ fail: ['b'] });
    const r = await runEnsemble({ messages }, plan('master'), call, { shuffle: (x) => x });
    const b = r.ensemble.members.find((m) => m.provider === 'b')!;
    expect(b.error).toMatch(/b is down/);
    expect(b.text).toBeUndefined();
    const prompt = lastUser(log.find((x) => x.kind === 'master')!.req);
    expect(prompt).toContain('answer from a');
    expect(prompt).toContain('answer from c');
    expect(prompt).not.toContain('Response C');
    expect(r.ensemble.note).toMatch(/b did not answer/);
  });

  it('returns the only answer when one member answers, and fails when none do', async () => {
    const one = mockProviders({ fail: ['b', 'c'] });
    const r = await runEnsemble({ messages }, plan('master'), one.call);
    expect(r.text).toBe('answer from a');
    expect(one.log.some((x) => x.kind === 'master')).toBe(false);
    expect(r.ensemble.note).toMatch(/Only a answered/);
    const none = mockProviders({ fail: ['a', 'b', 'c'] });
    await expect(runEnsemble({ messages }, plan('master'), none.call)).rejects.toThrow(/None of the 3 AIs answered/);
  });

  it('shows a member answer when the master fails', async () => {
    const { call } = mockProviders();
    const failing: MemberCall = (provider, model, req) => (provider === 'm' ? Promise.reject(new AiError('m is down', 'network')) : call(provider, model, req));
    const tokens: string[] = [];
    const r = await runEnsemble({ messages, onToken: (d) => tokens.push(d) }, plan('master'), failing);
    expect(r.text).toBe('answer from a');
    expect(tokens).toEqual(['answer from a']);
    expect(r.ensemble.note).toMatch(/The master, m, failed \(m is down\)/);
  });

  it('leaves out a member that times out', async () => {
    const call: MemberCall = async (provider, model, req) => {
      if (provider === 'b' && !lastUser(req).includes(MASTER_PROMPT)) await new Promise((res) => setTimeout(res, 500));
      return { text: `answer from ${provider}`, model, provider, providerName: provider };
    };
    const r = await runEnsemble({ messages }, plan('master', ['a', 'b', 'c'], { memberTimeoutMs: 50 }), call);
    expect(r.ensemble.members.find((m) => m.provider === 'b')!.error).toMatch(/No answer within/);
  });

  it('runs the on-device model one request at a time and the others in parallel', async () => {
    let running = 0, maxLocal = 0;
    const call: MemberCall = async (provider, model) => {
      if (provider.startsWith('local')) { running++; maxLocal = Math.max(maxLocal, running); }
      await new Promise((res) => setTimeout(res, 10));
      if (provider.startsWith('local')) running--;
      return { text: `answer from ${provider}`, model, provider, providerName: provider };
    };
    const p: EnsemblePlan = { mode: 'master', members: [{ provider: 'local', model: 'x', sequential: true }, { provider: 'cloud', model: 'y' }], master: { provider: 'local', model: 'x', sequential: true }, maxRounds: 3, memberTimeoutMs: 1000 };
    // The same sequential provider as member and master never overlaps with itself.
    await runEnsemble({ messages }, p, call);
    expect(maxLocal).toBe(1);
  });
});

describe('consensus mode', () => {
  it('reaches unity in round 1 when all unite', async () => {
    const { call, log } = mockProviders();
    const tokens: string[] = [];
    const r = await runEnsemble({ messages, onToken: (d) => tokens.push(d) }, plan('consensus'), call);
    expect(log.filter((x) => x.kind === 'draft')).toHaveLength(1);
    expect(log.filter((x) => x.kind === 'stance')).toHaveLength(3);
    expect(log.some((x) => x.kind === 'revise')).toBe(false);
    expect(r.ensemble.outcome).toBe('unity');
    expect(r.ensemble.rounds).toHaveLength(1);
    expect(r.ensemble.rounds![0].minute).toBe('minute 1');
    expect(r.text).toBe('minute 1');
    expect(tokens).toEqual(['minute 1']);
    expect(r.provider).toBe('m');
    expect(outcomeLine(r.ensemble)).toBe('United (3 of 3)');
    // Stance prompts carry the anonymized answers and the minute.
    const sp = lastUser(log.find((x) => x.kind === 'stance')!.req);
    expect(sp).toContain('Unite if the minute is a correct, complete answer you can support.');
    expect(sp).toContain('minute 1');
    expect(sp).toContain('Response A');
  });

  it('records stand-asides and counts unreadable replies as unite', async () => {
    const { call } = mockProviders({ stances: { b: ['stand-aside'] }, badJson: ['c'] });
    const r = await runEnsemble({ messages }, plan('consensus'), call);
    expect(r.ensemble.outcome).toBe('unity-with-stand-asides');
    expect(r.ensemble.standAsides).toEqual([{ provider: 'b', reason: 'b says stand-aside' }]);
    expect(r.ensemble.rounds![0].responses.find((x) => x.provider === 'c')).toMatchObject({ stance: 'unite', reason: expect.stringMatching(/could not be read/) });
    expect(outcomeLine(r.ensemble)).toBe('United; 1 stood aside');
  });

  it('a block triggers a revision and a second round', async () => {
    const { call, log } = mockProviders({ stances: { b: ['block', 'unite'] } });
    const status: string[] = [];
    const r = await runEnsemble({ messages, onStatus: (t) => status.push(t) }, plan('consensus'), call);
    expect(log.filter((x) => x.kind === 'revise')).toHaveLength(1);
    const rev = lastUser(log.find((x) => x.kind === 'revise')!.req);
    expect(rev).toContain(REVISE_RULE);
    expect(rev).toContain('b says block');
    expect(rev).toContain('use SUMIFS');
    expect(r.ensemble.rounds!.map((x) => x.minute)).toEqual(['minute 1', 'minute 2']);
    expect(r.ensemble.outcome).toBe('unity');
    expect(r.text).toBe('minute 2');
    expect(status).toContain('Round 1 of 3: 1 objection');
  });

  it('after maxRounds with a persistent block yields no-unity and still returns the minute', async () => {
    const { call, log } = mockProviders({ stances: { a: ['block'] } });
    const tokens: string[] = [];
    const r = await runEnsemble({ messages, onToken: (d) => tokens.push(d) }, plan('consensus', ['a', 'b', 'c'], { maxRounds: 2 }), call);
    expect(r.ensemble.outcome).toBe('no-unity');
    expect(r.ensemble.rounds).toHaveLength(2);
    expect(log.filter((x) => x.kind === 'revise')).toHaveLength(1); // no revision after the last round
    expect(r.text).toBe('minute 2');
    expect(tokens).toEqual(['minute 2']);
    expect(r.ensemble.blocks).toEqual([{ provider: 'a', reason: 'a says block Suggested change: use SUMIFS' }]);
    expect(outcomeLine(r.ensemble)).toBe('No unity after 2 rounds: 1 objection');
    // N×(rounds+1)+rounds requests at most
    expect(log.length).toBeLessThanOrEqual(requestCount('consensus', 3, 2));
  });
});

describe('abort and dispatch', () => {
  it('abort cancels all calls', async () => {
    const { call, log } = mockProviders({ hang: true });
    const ctl = new AbortController();
    const p = runEnsemble({ messages, signal: ctl.signal }, plan('master'), call);
    await new Promise((res) => setTimeout(res, 5));
    expect(log).toHaveLength(3);
    ctl.abort();
    await expect(p).rejects.toMatchObject({ code: 'aborted' });
    expect(log.every((x) => x.req.signal?.aborted)).toBe(true);
    // Nothing more is asked after Stop.
    await new Promise((res) => setTimeout(res, 5));
    expect(log).toHaveLength(3);
  });

  it('plans only for normal requests with two or more usable members', () => {
    const lookup = (id: string) => (['a', 'b', 'local'].includes(id) ? { model: `${id}-q`, sequential: id === 'local' } : null);
    const ens = { ...defaultEnsemble(), mode: 'master' as const, members: [{ provider: 'a' }, { provider: 'b', model: 'b-custom' }, { provider: 'gone' }], master: 'b' };
    const p = planEnsemble({}, ens, lookup)!;
    expect(p.members).toEqual([{ provider: 'a', model: 'a-q', sequential: false }, { provider: 'b', model: 'b-custom', sequential: false }]);
    expect(p.master).toEqual({ provider: 'b', model: 'b-custom', sequential: false });
    expect(planEnsemble({ provider: 'a' }, ens, lookup)).toBeNull();
    expect(planEnsemble({}, { ...ens, mode: 'single' }, lookup)).toBeNull();
    expect(planEnsemble({}, { ...ens, members: [{ provider: 'a' }, { provider: 'gone' }] }, lookup)).toBeNull();
    // A master that is not set up falls back to the first member.
    expect(planEnsemble({}, { ...ens, master: 'gone' }, lookup)!.master.provider).toBe('a');
  });

  it('key tests and model lists bypass the ensemble', async () => {
    const { saveAiSettings } = await import('../../src/ai/store');
    const { chat, listModels } = await import('../../src/ai/client');
    const { resolveProvider } = await import('../../src/ai/store');
    const hits: string[] = [];
    vi.stubGlobal('fetch', async (url: string, init?: RequestInit) => {
      hits.push(`${init?.method || 'GET'} ${url}`);
      if (String(url).endsWith('/models')) return new Response(JSON.stringify({ data: [{ id: 'm1' }] }), { status: 200 });
      const body = JSON.parse(String(init?.body));
      const last = body.messages[body.messages.length - 1].content;
      const text = last.includes(MASTER_PROMPT) ? 'combined' : `from ${url}`;
      return new Response(JSON.stringify({ choices: [{ message: { content: text }, finish_reason: 'stop' }], usage: { prompt_tokens: 3, completion_tokens: 2 } }), { status: 200 });
    });
    saveAiSettings((s) => {
      s.providers = {
        ollama: { id: 'ollama', model: 'o1', baseUrl: 'http://localhost:4901/v1' },
        lmstudio: { id: 'lmstudio', model: 'l1', baseUrl: 'http://localhost:4902/v1' },
      };
      s.active = 'ollama';
      s.ensemble = normalizeEnsemble({ mode: 'master', members: [{ provider: 'ollama' }, { provider: 'lmstudio' }], master: 'lmstudio' });
    });
    // A request naming a provider (the key test, "Try it") goes to that provider alone.
    hits.length = 0;
    const one = await chat({ provider: 'ollama', messages: [{ role: 'user', content: 'Reply with OK.' }] });
    expect(hits).toEqual(['POST http://localhost:4901/v1/chat/completions']);
    expect(one.ensemble).toBeUndefined();
    // Model lists never chat.
    hits.length = 0;
    expect(await listModels((await resolveProvider('lmstudio'))!)).toEqual(['m1']);
    expect(hits).toEqual(['GET http://localhost:4902/v1/models']);
    // Normal commands use the ensemble: two members, then the master.
    hits.length = 0;
    const all = await chat({ messages: [{ role: 'user', content: 'Total A1:A3' }] });
    expect(hits).toHaveLength(3);
    expect(all.text).toBe('combined');
    expect(all.provider).toBe('lmstudio');
    expect(all.usage).toEqual({ in: 9, out: 6 });
    expect(all.ensemble?.members.map((m) => m.provider)).toEqual(['ollama', 'lmstudio']);
    saveAiSettings((s) => { s.ensemble.mode = 'single'; });
    hits.length = 0;
    await chat({ messages: [{ role: 'user', content: 'Total A1:A3' }] });
    expect(hits).toEqual(['POST http://localhost:4901/v1/chat/completions']);
    vi.unstubAllGlobals();
  });
});

describe('settings', () => {
  it('settings export never contains keys', async () => {
    vi.stubGlobal('localStorage', {
      getItem: (k: string) => (mem.has(k) ? mem.get(k)! : null),
      setItem: (k: string, v: string) => { mem.set(k, String(v)); },
      removeItem: (k: string) => { mem.delete(k); },
    });
    const { saveAiSettings, aiSettings, setApiKey, reloadAiSettings } = await import('../../src/ai/store');
    const secret = 'sk-test-0123456789abcdef0123456789abcdef';
    await setApiKey('custom', secret);
    // Even settings text that carries a key next to a member is cleaned on load.
    mem.set('quillgrid.ai.v1', JSON.stringify({ ...aiSettings(), ensemble: { mode: 'consensus', members: [{ provider: 'custom', model: 'x', key: secret, apiKey: secret }, { provider: 'groq' }], master: 'groq', maxRounds: 9 } }));
    reloadAiSettings();
    const s = aiSettings();
    expect(s.ensemble).toEqual({ mode: 'consensus', members: [{ provider: 'custom', model: 'x' }, { provider: 'groq' }], master: 'groq', maxRounds: 5, showDeliberation: true, memberTimeoutMs: 90000 });
    saveAiSettings();
    expect(JSON.stringify(aiSettings())).not.toContain(secret);
    expect(mem.get('quillgrid.ai.v1')).not.toContain(secret);
    expect(mem.get('quillgrid.ai.v1')).toContain('"ensemble"');
    vi.unstubAllGlobals();
  });

  it('cost note and stance parsing', () => {
    expect(costNote(3, 3)).toBe('Each request is sent to 3 services (about 4 requests for master mode, up to 15 for consensus). Paid services charge for each.');
    expect(parseStance('{"stance":"Stand aside","reason":"prefer XLOOKUP"}')).toEqual({ stance: 'stand-aside', reason: 'prefer XLOOKUP' });
    expect(parseStance('<think>hm</think> {"stance":"block","reason":"wrong range","suggested_change":"A1:A4"}')).toEqual({ stance: 'block', reason: 'wrong range', suggestedChange: 'A1:A4' });
    expect(parseStance('sure').unreadable).toBe(true);
    expect(normalizeEnsemble(null)).toEqual(defaultEnsemble());
  });
});
