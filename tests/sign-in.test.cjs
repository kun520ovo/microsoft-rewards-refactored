const assert = require('node:assert/strict');
const { webcrypto } = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const scriptPath = process.env.REWARDS_TEST_SCRIPT ||
  path.resolve(__dirname, '..', '微软积分商城签到（重构版）1.0.4.js');
const script = fs.readFileSync(scriptPath, 'utf8');
const today = 20261023;

// Synthetic Flight records contain only the fields read from the Rewards page.
function page({ app = true, pc = true, appEnabled = true, pcEnabled = true } = {}) {
  const tasks = [
    { partner: 'dailyset', activitiesCompleted: 3, activitiesTotal: 3,
      isCurrentDayCompleted: true, isEnabled: true, completedDays: 4 },
  ];
  for (const [partner, done, enabled] of [['bing', pc, pcEnabled], ['bingapp', app, appEnabled]]) {
    if (done === null) continue;
    tasks.push({ partner, activitiesCompleted: done ? 1 : 0, activitiesTotal: 1,
      isCurrentDayCompleted: done, isEnabled: enabled, completedDays: 4 });
  }
  const record = '0:' + JSON.stringify({ streakTasks: tasks }) + '\n';
  return '<html><script>self.__next_f.push([1,' + JSON.stringify(record) + '])</script></html>';
}

function harness({ html = page(), storage = {}, appResponse, finalUrl, pageError } = {}) {
  const values = new Map(Object.entries(structuredClone(storage)));
  const requests = [], logs = [], delays = [];
  const context = {
    URL, URLSearchParams, setTimeout, clearTimeout, setInterval, clearInterval, crypto: webcrypto,
    location: new URL('https://www.bing.com/'),
    navigator: { userAgent: 'Rewards regression test' },
    GM_info: { script: { name: 'Rewards regression test', version: '1.0.4' } },
    GM_getValue: (key, fallback) => values.has(key) ? structuredClone(values.get(key)) : fallback,
    GM_setValue: (key, value) => values.set(key, structuredClone(value)),
    GM_deleteValue: key => values.delete(key),
    GM_listValues: () => [...values.keys()],
    GM_log: value => logs.push(String(value)),
    GM_notification: () => {},
    GM_openInTab: () => { throw new Error('Unexpected tab opening'); },
    GM_xmlhttpRequest: () => { throw new Error('Unexpected real network request'); },
    console: { log: () => {}, debug: () => {}, error: () => {} },
  };
  const entry = '  AppEntry.start();';
  assert.equal(script.split(entry).length, 2, 'Locate exactly one startup call');
  const instrumented = script.replace(entry,
    '  globalThis.testApi = { DataParsers, TaskState, TaskRecords, TaskActions, TaskManager, ' +
    'RunState, RunSummary, AppSignIn, HttpClient, RunGuard, Wait, Utils, WebSession };');
  vm.runInNewContext(instrumented, context, { filename: scriptPath, timeout: 1000 });
  const api = context.testApi;
  api.Utils.getTodayNum = () => today;
  api.RunState.dateNowNum = today;
  api.RunState.token = 'synthetic-app-token';
  api.RunGuard.assertActive = () => {};
  api.WebSession.ensure = async () => true;
  api.Wait.randomDelay = async () => {};
  api.Wait.delay = async ms => { delays.push(ms); };
  api.TaskRecords.load(api.TaskState);
  const fixture = { html, finalUrl: finalUrl || 'https://rewards.bing.com/earn', pageError };
  api.HttpClient.text = async options => {
    requests.push(options);
    const url = new URL(options.url);
    if (url.origin === 'https://rewards.bing.com' && url.pathname === '/earn') {
      if (fixture.pageError) throw fixture.pageError;
      return options.fullResponse
        ? { status: 200, finalUrl: fixture.finalUrl, responseText: fixture.html }
        : fixture.html;
    }
    if (url.origin === 'https://prod.rewardsplatform.microsoft.com' && url.pathname === '/dapi/me/activities') {
      return JSON.stringify(appResponse || { response: { activity: { p: 5 } } });
    }
    throw new Error('Unexpected request: ' + url.origin + url.pathname);
  };
  return {
    api, values, requests, logs, delays, fixture,
    posts: () => requests.filter(request => request.method === 'POST'),
    summary: () => api.RunSummary.format({
      info: { balance: 28286, pc: { progress: 60, max: 60 }, searchStatus: 'available', readProgress: 30, readMax: 30 },
      startBalance: 28119, endBalance: 28286, today,
      tasks: api.TaskState, read: { progress: 30, max: 30 }, promosEnabled: true,
    }),
  };
}

test('官网两端已完成：跳过重复签到，不请求旧令牌，不误报 PC 失败', async () => {
  const h = harness();
  const result = await h.api.TaskActions.doSign();
  assert.equal(result.status, 'completed');
  assert.equal(h.posts().length, 0);
  assert.ok(h.requests.every(request => new URL(request.url).pathname === '/earn'));
  assert.equal(h.api.TaskState.appSignDate, today);
  assert.equal(h.api.TaskState.pcSignDate, today);
  assert.equal(h.api.TaskState.signDate, today);
  assert.equal(h.values.get('Config.tasks').sign, today);
  assert.ok(h.logs.every(line => !/PC签入失败|RequestVerificationToken/.test(line)));
  assert.match(h.summary(), /^签到\s+✅/);
});

test('PC 未完成时先完成 App，并让搜索流程继续而不是原地重试', async () => {
  const h = harness({ html: page({ app: false, pc: false }) });
  const result = await h.api.TaskManager.runTask(() => h.api.TaskActions.doSign(), '签到');
  assert.equal(result.status, 'pending');
  assert.equal(result.retryable, false);
  assert.equal(h.posts().length, 1);
  assert.equal(h.api.TaskState.appSignDate, today);
  assert.equal(h.api.TaskState.pcSignDate, 0);
  assert.equal(h.api.TaskState.signDate, 0);
  assert.equal(h.api.TaskState.signPoint, 5);
  assert.ok(!h.delays.includes(60000));
  assert.match(h.summary(), /^签到\s+⏳ App 已完成 \/ PC 待确认/);
});

test('搜索后的连签复核确认 PC，保留 App 积分且不重复提交或计分', async () => {
  const h = harness({ html: page({ app: false, pc: false }) });
  await h.api.TaskActions.doSign();
  h.fixture.html = page();
  const result = await h.api.TaskActions.doStreak();
  assert.equal(result.status, 'completed');
  assert.equal(h.api.TaskState.signDate, today);
  assert.equal(h.api.TaskState.signPoint, 5);
  assert.equal(h.posts().length, 1);
  assert.equal(h.values.get('Config.tasks').signPC, today);
  assert.ok(h.logs.some(line => line === '💻 PC签入成功！'));
  assert.match(h.summary(), /^签到\s+✅/);
});

test('刷新后保留 App 的确认结果，PC 未完成时不会重复签到', async () => {
  const first = harness({ html: page({ app: false, pc: false }) });
  await first.api.TaskActions.doSign();
  const h = harness({ html: page({ app: true, pc: false }), storage: Object.fromEntries(first.values) });
  const result = await h.api.TaskActions.doSign();
  assert.equal(result.status, 'pending');
  assert.equal(h.posts().length, 0);
  assert.equal(h.api.TaskState.signPoint, 5);
  assert.equal(h.api.TaskState.signDate, 0);
});

test('仅 PC 完成不能将失败或未确认的 App 也判为完成', async () => {
  const h = harness({ html: page({ app: false }), appResponse: { error: 'unavailable' } });
  const result = await h.api.TaskActions.doSign();
  assert.equal(result.status, 'unknown');
  assert.equal(h.api.TaskState.pcSignDate, today);
  assert.equal(h.api.TaskState.appSignDate, 0);
  assert.equal(h.api.TaskState.signDate, 0);
  assert.match(h.summary(), /^签到\s+⏳ App 待确认 \/ PC 已完成/);
});

test('旧版任一端成功的缓存不能遮盖未完成的 PC 打卡', async () => {
  const h = harness({ html: page({ pc: false }), storage: {
    'Config.tasks': { sign: today }, 'Config.signPoint': 5,
  } });
  assert.equal(h.api.TaskState.signDate, 0);
  const result = await h.api.TaskActions.doSign();
  assert.equal(result.status, 'pending');
  assert.equal(h.values.get('Config.tasks').sign, 0);
  assert.equal(h.values.get('Config.tasks').signApp, today);
});

test('页面只有其他已完成任务时，PC 和 App 保持未知，不调用旧接口', async () => {
  const h = harness({ html: page({ app: null, pc: null }), appResponse: { error: 'unavailable' } });
  const result = await h.api.TaskActions.doSign();
  assert.equal(result.status, 'unknown');
  assert.equal(h.api.TaskState.signDate, 0);
  assert.equal(h.api.TaskState.pcSignDate, 0);
  assert.ok(h.requests.every(request => !request.url.includes('reportactivity')));
  assert.ok(h.logs.some(line => line.includes('PC打卡状态暂未确认')));
});

test('重定向到登录页时不采信页面内的完成标记', async () => {
  const h = harness({ finalUrl: 'https://login.live.com/oauth20_authorize.srf', appResponse: { error: 'unavailable' } });
  await h.api.TaskActions.doSign();
  assert.equal(h.api.TaskState.appSignDate, 0);
  assert.equal(h.api.TaskState.pcSignDate, 0);
  assert.equal(h.api.TaskState.signDate, 0);
});

test('跨日不复用昨天的完成记录或签到积分', async () => {
  const h = harness({ html: page({ app: false, pc: false }), storage: {
    'Config.tasks': { sign: today - 1, signApp: today - 1, signPC: today - 1 },
    'Config.signPoint': 10,
  } });
  assert.equal(h.api.TaskState.signDate, 0);
  assert.equal(h.api.TaskState.signPoint, -1);
  await h.api.TaskActions.doSign();
  assert.equal(h.api.TaskState.appSignDate, today);
  assert.equal(h.api.TaskState.signDate, 0);
  assert.equal(h.api.TaskState.signPoint, 5);
});

test('官网未启用的 PC 卡片不能作为打卡完成证据', async () => {
  const h = harness({ html: page({ pcEnabled: false }) });
  const result = await h.api.TaskActions.doSign();
  assert.equal(result.status, 'pending');
  assert.equal(h.api.TaskState.pcSignDate, 0);
  assert.equal(h.api.TaskState.signDate, 0);
});

test('当日已核验的完成记录在临时查询失败时仍保留', async () => {
  const h = harness({ pageError: new Error('Temporary network failure'), storage: {
    'Config.tasks': { sign: today, signApp: today, signPC: today },
    'Config.signPoint': 5,
  } });
  const result = await h.api.TaskActions.doSign();
  assert.equal(result.status, 'completed');
  assert.equal(h.posts().length, 0);
  assert.equal(h.api.TaskState.signPoint, 5);
});

test('官网明确显示 PC 尚未完成时撤回过时的完成记录', async () => {
  const h = harness({ html: page({ pc: false }), storage: {
    'Config.tasks': { sign: today, signApp: today, signPC: today },
    'Config.signPoint': 5,
  } });
  const result = await h.api.TaskActions.doSign();
  assert.equal(result.status, 'pending');
  assert.equal(h.api.TaskState.pcSignDate, 0);
  assert.equal(h.values.get('Config.tasks').sign, 0);
});

test('会话中止信号必须继续上抛，不能保存签到完成结果', async () => {
  const stop = Object.assign(new Error('Session stopped'), { code: 'WEB_SESSION_STOPPED' });
  const h = harness({ pageError: stop });
  await assert.rejects(() => h.api.TaskActions.doSign(), /Session stopped/);
  assert.equal(h.values.has('Config.tasks'), false);
});

test('关闭签到时不请求页面或提交签到', async () => {
  const h = harness({ storage: { 'Tasks.sign': false } });
  const result = await h.api.TaskActions.doSign();
  assert.equal(result.status, 'skipped');
  assert.equal(h.requests.length, 0);
});
