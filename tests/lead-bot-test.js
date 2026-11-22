// Mock tests for Code.gs (fake Apps Script services, no real API calls).
// Run from the repo root:  node tests/lead-bot-test.js Code.gs
const fs = require('fs');
const vm = require('vm');
const crypto = require('crypto');
const assert = require('assert');

const code = fs.readFileSync(process.argv[2], 'utf8');
const ALL_KEYS = { GEMINI_API_KEY: 'k1', WA_TOKEN: 't1', WA_PHONE_NUMBER_ID: '123456' };

function chain(calls, name, overrides = {}) {
  const proxy = new Proxy({}, {
    get: (_t, prop) => (prop in overrides ? overrides[prop] : (...args) => { calls.push(name + '.' + String(prop)); return proxy; }),
  });
  return proxy;
}

function makeEnv({ props = { ...ALL_KEYS }, fetchHandler = () => ({ code: 200, body: {} }), hasLog = true, mailThrows = null, whatsapp = false, ownerEmail = 'owner@gmail.com' } = {}) {
  const rows = [], fetches = [], logs = [], cache = {}, calls = [], mails = [];
  let triggers = [];
  const logSheet = chain(calls, 'log', {
    appendRow: (r) => rows.push(r),
    getLastRow: () => rows.length,
    getLastColumn: () => (rows[0] || []).length,
    getRange: (...args) => chain(calls, 'range', { getValues: () => [rows[0] ? Array.from(rows[0]) : []] }),
  });
  let logExists = hasLog;
  const ss = {
    getSheetByName: (n) => (n === 'Lead Log' && logExists ? logSheet : null),
    insertSheet: () => { logExists = true; return logSheet; },
    getId: () => 'ss1',
  };
  const form = chain(calls, 'form', { getId: () => 'form1', getPublishedUrl: () => 'https://forms/pub', getEditUrl: () => 'https://forms/edit' });
  const ctx = {
    Logger: { log: (m) => logs.push(String(m)) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => props[k] ?? null, setProperty: (k, v) => { props[k] = v; } }) },
    SpreadsheetApp: { getActiveSpreadsheet: () => ss },
    UrlFetchApp: {
      fetch: (url, opts = {}) => {
        fetches.push({ url, opts });
        const r = fetchHandler(url, opts);
        return { getResponseCode: () => r.code, getContentText: () => (typeof r.body === 'string' ? r.body : JSON.stringify(r.body)) };
      },
    },
    MailApp: {
      sendEmail: (msg) => { if (mailThrows) throw new Error(mailThrows); mails.push(msg); },
      getRemainingDailyQuota: () => 99,
    },
    Session: { getEffectiveUser: () => ({ getEmail: () => ownerEmail }) },
    CacheService: { getScriptCache: () => ({ get: (k) => cache[k] ?? null, put: (k, v) => { cache[k] = v; } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Utilities: {
      computeDigest: (_a, s) => Array.from(crypto.createHash('sha256').update(s).digest()).map((b) => (b > 127 ? b - 256 : b)),
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      Charset: { UTF_8: 'UTF_8' },
      base64EncodeWebSafe: (bytes) => Buffer.from(bytes.map((b) => b & 0xff)).toString('base64url'),
      sleep: () => {},
    },
    FormApp: {
      create: () => { calls.push('FormApp.create'); return form; },
      openById: (id) => { if (id === 'form1') return form; throw new Error('gone'); },
      createTextValidation: () => chain(calls, 'validation'),
      DestinationType: { SPREADSHEET: 'SPREADSHEET' },
    },
    ScriptApp: {
      getProjectTriggers: () => triggers,
      deleteTrigger: (t) => { triggers = triggers.filter((x) => x !== t); },
      newTrigger: (fn) => ({ forSpreadsheet: () => ({ onFormSubmit: () => ({ create: () => { triggers.push({ getHandlerFunction: () => fn }); } }) }) }),
    },
  };
  vm.createContext(ctx);
  vm.runInContext(code, ctx);
  if (whatsapp) vm.runInContext('CONFIG.SEND_WHATSAPP = true', ctx);
  return { ctx, rows, fetches, logs, props, calls, mails, triggers: () => triggers };
}

const aiOk = (text) => ({ code: 200, body: { candidates: [{ content: { role: 'model', parts: [{ text: 'thinking...', thought: true }, { text }] }, finishReason: 'STOP' }] } });
const waOk = { code: 200, body: { messages: [{ id: 'wamid.1' }] } };
const isAi = (url) => url.includes('generativelanguage.googleapis.com');
const happy = (url) => (isAi(url) ? aiOk('Hi Priya! ...') : waOk);
const event = (nv, uid) => ({ namedValues: nv, ...(uid ? { triggerUid: uid } : {}) });
const nv = (name, phone, interest, key = 'What specifically are you looking for?', email = 'priya@example.com') => ({ Name: [name], 'Which service do you need?': ['Lead automation'], Email: [email], 'WhatsApp Number': [phone], [key]: [interest] });

let passed = 0;
function test(title, fn) { fn(); passed++; console.log('PASS', title); }

test('default (email on, WhatsApp off): Gemini reply is emailed and logged, no WhatsApp call', () => {
  const env = makeEnv({ fetchHandler: happy });
  env.ctx.onFormSubmit(event(nv('Priya Shah', '98765 43210', 'A CRM for my clinic')));
  assert.strictEqual(env.rows.length, 1);
  const [ts, name, email, phone, service, interest, msg, status] = env.rows[0];
  assert.ok(Object.prototype.toString.call(ts) === '[object Date]');
  assert.deepStrictEqual([name, email, phone, service, interest, msg, status], ['Priya Shah', 'priya@example.com', '919876543210', 'Lead automation', 'A CRM for my clinic', 'Hi Priya! ...', 'Email sent']);

  assert.strictEqual(env.fetches.length, 1);
  const ai = env.fetches[0];
  assert.strictEqual(ai.url, 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash-lite:generateContent');
  assert.strictEqual(ai.opts.headers['x-goog-api-key'], 'k1');
  const aBody = JSON.parse(ai.opts.payload);
  assert.ok(aBody.systemInstruction.parts[0].text.includes('PPS Anantams'));
  assert.strictEqual(aBody.contents[0].parts[0].text, ['First name: Priya', 'Service: Lead automation', 'Details: A CRM for my clinic'].join(String.fromCharCode(10)));

  assert.strictEqual(env.mails.length, 1);
  assert.deepStrictEqual({ ...env.mails[0] }, { to: 'priya@example.com', subject: 'Thanks for contacting PPS Anantams', body: 'Hi Priya! ...', name: 'PPS Anantams' });
});

test('repeated sheet columns (old empty + new answer) still give the real answers', () => {
  const env = makeEnv({ fetchHandler: happy });
  env.ctx.onFormSubmit(event({
    Name: ['', 'Manav'],
    Email: ['', 'manav@example.com'],
    'WhatsApp Number': ['', ''],
    'What are you looking for?': [''],
    'Which service do you need?': ['Website design'],
    'What specifically are you looking for?': ['A site for my gym'],
  }, 'uid9'));
  assert.deepStrictEqual(Array.from(env.rows[0]).slice(1), ['Manav', 'manav@example.com', '', 'Website design', 'A site for my gym', 'Hi Priya! ...', 'Email sent']);
  assert.strictEqual(env.mails[0].to, 'manav@example.com');
});

test('only the Gemini key is needed while WhatsApp is off', () => {
  const env = makeEnv({ props: { GEMINI_API_KEY: 'k1' }, fetchHandler: happy });
  env.ctx.onFormSubmit(event(nv('A', '919876543210', 'x')));
  assert.strictEqual(env.rows[0][7], 'Email sent');
});

test('both channels on: email and WhatsApp both sent', () => {
  const env = makeEnv({ fetchHandler: happy, whatsapp: true });
  env.ctx.onFormSubmit(event(nv('Priya Shah', '98765 43210', 'A CRM')));
  assert.strictEqual(env.rows[0][7], 'Email sent; WhatsApp sent');
  const wa = env.fetches[1];
  assert.strictEqual(wa.url, 'https://graph.facebook.com/v25.0/123456/messages');
  assert.strictEqual(wa.opts.headers.Authorization, 'Bearer t1');
  const wBody = JSON.parse(wa.opts.payload);
  assert.deepStrictEqual([wBody.to, wBody.type, wBody.text.body], ['919876543210', 'text', 'Hi Priya! ...']);
});

test('WhatsApp error does not stop the email, and shows a hint', () => {
  const env = makeEnv({ whatsapp: true, fetchHandler: (url) => (isAi(url) ? aiOk('Hello A') : { code: 400, body: { error: { message: 'Business Account locked', code: 131031 } } }) });
  env.ctx.onFormSubmit(event(nv('A', '919876543210', 'x')));
  assert.strictEqual(env.rows[0][6], 'Hello A');
  assert.match(env.rows[0][7], /^Email sent; WhatsApp failed: 400: Business Account locked -> Meta locked/);
  assert.strictEqual(env.mails.length, 1);
});

test('email error is logged with the AI message kept', () => {
  const env = makeEnv({ fetchHandler: happy, mailThrows: 'Service invoked too many times for one day: email.' });
  env.ctx.onFormSubmit(event(nv('A', '919876543210', 'x')));
  assert.strictEqual(env.rows[0][6], 'Hi Priya! ...');
  assert.strictEqual(env.rows[0][7], 'Email failed: Service invoked too many times for one day: email.');
});

test('bad or injected email -> logged failure, no API calls, no mail', () => {
  for (const bad of ['not-an-email', 'a@b.com, c@d.com', 'a@b.com\nBcc: x@y.com']) {
    const env = makeEnv({ fetchHandler: happy });
    env.ctx.onFormSubmit(event(nv('A', '919876543210', 'x', undefined, bad)));
    assert.match(env.rows[0][7], /^Failed: Email ".*" is not valid/);
    assert.strictEqual(env.fetches.length, 0);
    assert.strictEqual(env.mails.length, 0);
  }
});

test('bad phone only matters when WhatsApp is on', () => {
  let env = makeEnv({ fetchHandler: happy });
  env.ctx.onFormSubmit(event(nv('A', '12345', 'x')));
  assert.strictEqual(env.rows[0][7], 'Email sent');
  assert.strictEqual(env.rows[0][3], '12345');
  env = makeEnv({ fetchHandler: happy, whatsapp: true });
  env.ctx.onFormSubmit(event(nv('A', '12345', 'x')));
  assert.match(env.rows[0][7], /^Failed: WhatsApp number "12345" is not valid/);
  assert.strictEqual(env.fetches.length, 0);
});

test('old "Interest" title from a hand-made form still works', () => {
  const env = makeEnv({ fetchHandler: happy });
  env.ctx.onFormSubmit(event(nv('Ravi', '+91 98765-43210', 'Chatbot', 'Interest')));
  assert.strictEqual(env.rows[0][5], 'Chatbot');
  assert.strictEqual(env.rows[0][7], 'Email sent');
});

test('Gemini problems are reported with hints and nothing is sent', () => {
  const cases = [
    [{ code: 200, body: { candidates: [{ content: { parts: [{ text: 'hmm', thought: true }] }, finishReason: 'MAX_TOKENS' }] } }, /^Failed: Gemini returned no text \(finishReason: MAX_TOKENS\)/],
    [{ code: 200, body: { promptFeedback: { blockReason: 'SAFETY' } } }, /^Failed: Gemini blocked this request \(SAFETY\)$/],
    [{ code: 400, body: { error: { code: 400, message: 'API key not valid.', status: 'INVALID_ARGUMENT' } } }, /^Failed: Gemini 400: API key not valid\. -> Check GEMINI_API_KEY/],
    [{ code: 429, body: { error: { message: 'Resource exhausted' } } }, /^Failed: Gemini 429: Resource exhausted -> Free limit reached/],
    [{ code: 502, body: '<html>Bad gateway</html>' }, /^Failed: Gemini 502: no details$/],
  ];
  for (const [reply, expected] of cases) {
    const env = makeEnv({ fetchHandler: () => reply });
    env.ctx.onFormSubmit(event(nv('A', '919876543210', 'x')));
    assert.match(env.rows[0][7], expected);
    assert.strictEqual(env.mails.length, 0);
  }
});

test('busy model (503) -> next model is used and the email still goes out', () => {
  const env = makeEnv({ fetchHandler: (url) => (url.includes('gemini-3.5-flash-lite') ? { code: 503, body: { error: { code: 503, message: 'This model is currently experiencing high demand.' } } } : aiOk('Hi from backup')) });
  env.ctx.onFormSubmit(event(nv('A', '919876543210', 'x')));
  assert.deepStrictEqual(env.fetches.map((f) => f.url.split('/models/')[1]), ['gemini-3.5-flash-lite:generateContent', 'gemini-3.6-flash:generateContent']);
  assert.strictEqual(env.rows[0][6], 'Hi from backup');
  assert.strictEqual(env.rows[0][7], 'Email sent');
});

test('all models busy -> tries each once, logs the last error; bad key stops at once', () => {
  let env = makeEnv({ fetchHandler: () => ({ code: 503, body: { error: { message: 'high demand' } } }) });
  env.ctx.onFormSubmit(event(nv('A', '919876543210', 'x')));
  assert.strictEqual(env.fetches.length, 3);
  assert.strictEqual(env.rows[0][7], 'Failed: Gemini 503: high demand -> Google is busy right now. Wait a minute and try again.');
  env = makeEnv({ fetchHandler: () => ({ code: 400, body: { error: { message: 'API key not valid.' } } }) });
  env.ctx.onFormSubmit(event(nv('A', '919876543210', 'x')));
  assert.strictEqual(env.fetches.length, 1);
});

test('missing Gemini key -> logged failure, no API calls', () => {
  const env = makeEnv({ props: {}, fetchHandler: happy });
  env.ctx.onFormSubmit(event(nv('A', '919876543210', 'x')));
  assert.strictEqual(env.rows[0][7], 'Failed: Script property GEMINI_API_KEY is missing');
  assert.strictEqual(env.fetches.length, 0);
});

test('formula text from the form is stored as plain text', () => {
  const env = makeEnv({ fetchHandler: happy });
  env.ctx.onFormSubmit(event(nv('=HYPERLINK("http://x","y")', '919876543210', '+cmd')));
  assert.strictEqual(env.rows[0][1], '\'=HYPERLINK("http://x","y")');
  assert.strictEqual(env.rows[0][5], "'+cmd");
});

test('duplicate trigger fire is skipped; empty event ignored', () => {
  const env = makeEnv({ fetchHandler: happy });
  env.ctx.onFormSubmit(event(nv('A', '919876543210', 'x'), 'uid1'));
  env.ctx.onFormSubmit(event(nv('A', '919876543210', 'x'), 'uid1'));
  env.ctx.onFormSubmit(event({}, 'uid1'));
  assert.strictEqual(env.rows.length, 1);
  assert.strictEqual(env.mails.length, 1);
});

test('phone normalisation cases', () => {
  const n = makeEnv().ctx.normalizePhone_;
  assert.strictEqual(n('9876543210'), '919876543210');
  assert.strictEqual(n('09876543210'), '919876543210');
  assert.strictEqual(n('0091 98765 43210'), '919876543210');
  assert.strictEqual(n('+1 (415) 555-0100'), '14155550100');
  assert.strictEqual(n('123'), '');
});

test('setup makes tab with Email column, form with email question, one trigger (run twice)', () => {
  const env = makeEnv({ props: {}, hasLog: false });
  env.ctx.setup();
  env.ctx.setup();
  assert.deepStrictEqual(Array.from(env.rows[0]), ['Timestamp', 'Name', 'Email', 'WhatsApp', 'Service', 'Details', 'AI Message', 'Status']);
  assert.strictEqual(env.rows.length, 1);
  assert.strictEqual(env.props.FORM_ID, 'form1');
  assert.ok(env.calls.includes('form.setDestination'));
  assert.ok(env.calls.includes('validation.requireTextIsEmail'));
  assert.ok(env.calls.includes('form.addListItem') && env.calls.includes('form.setChoiceValues'));
  assert.strictEqual(env.calls.filter((c) => c === 'FormApp.create').length, 1);
  assert.ok(env.calls.includes('validation.requireTextMatchesPattern'));
  assert.strictEqual(env.triggers().length, 1);
  assert.ok(env.logs.some((l) => l.includes('https://forms/pub')));
});

test('old Lead Log with different columns is kept aside and a fresh tab is made', () => {
  const env = makeEnv({ props: {} });
  env.rows.push(['Timestamp', 'Name', 'Email', 'WhatsApp', 'Interest', 'AI Message', 'Status']);
  env.ctx.setup();
  assert.ok(env.calls.includes('log.setName'));
});

test('testRun emails the owner by default and logs a row', () => {
  const env = makeEnv({ fetchHandler: happy });
  env.ctx.testRun();
  assert.strictEqual(env.mails[0].to, 'owner@gmail.com');
  assert.strictEqual(env.rows[0][2], 'owner@gmail.com');
  assert.strictEqual(env.rows[0][3], '91XXXXXXXXXX');
  assert.strictEqual(env.rows[0][7], 'Email sent');
  assert.ok(env.logs.some((l) => l.includes('Check the Lead Log tab and your email')));
});

test('testRun guards a missing email and a placeholder phone', () => {
  let env = makeEnv({ fetchHandler: happy, ownerEmail: '' });
  env.ctx.testRun();
  assert.strictEqual(env.rows.length, 0);
  assert.ok(env.logs[0].includes('CONFIG.TEST_EMAIL'));
  env = makeEnv({ fetchHandler: happy, whatsapp: true });
  vm.runInContext("CONFIG.TEST_PHONE = '91XXXXXXXXXX'", env.ctx);
  env.ctx.testRun();
  assert.strictEqual(env.rows.length, 0);
  assert.ok(env.logs[0].includes('CONFIG.TEST_PHONE'));
});

test('checkSetup (WhatsApp off): asks only for the Gemini key, checks email, never prints keys', () => {
  const env = makeEnv({ props: { GEMINI_API_KEY: 'AIzaSecret' }, fetchHandler: (url) => (isAi(url) ? aiOk('Hi Asha!') : waOk) });
  env.ctx.setup();
  env.ctx.checkSetup();
  const out = env.logs.join('\n');
  assert.ok(!out.includes('AIzaSecret'));
  assert.match(out, /GEMINI_API_KEY: set/);
  assert.ok(!out.includes('WA_TOKEN'));
  assert.match(out, /Gemini: OK\. Sample reply: Hi Asha!/);
  assert.match(out, /Email: OK\. Emails left today: 99/);
  assert.match(out, /WhatsApp: off/);
  assert.ok(!env.fetches.some((f) => f.url.includes('graph.facebook.com')));
  assert.strictEqual(env.mails.length, 0);
});

test('checkSetup (WhatsApp on): checks the phone number id with a GET', () => {
  const env = makeEnv({ whatsapp: true, props: { GEMINI_API_KEY: 'AIzaSecret', WA_TOKEN: 'EAAsecret', WA_PHONE_NUMBER_ID: '123456' }, fetchHandler: (url) => (isAi(url) ? aiOk('Hi Asha!') : { code: 200, body: { display_phone_number: '+1 555 0100' } }) });
  env.ctx.setup();
  env.ctx.checkSetup();
  const out = env.logs.join('\n');
  assert.ok(!out.includes('AIzaSecret') && !out.includes('EAAsecret'));
  assert.match(out, /WhatsApp: OK\. Sending from \+1 555 0100/);
  const waCheck = env.fetches.find((f) => f.url.includes('graph.facebook.com'));
  assert.strictEqual(waCheck.url, 'https://graph.facebook.com/v25.0/123456?fields=display_phone_number,verified_name');
});

console.log('\nAll ' + passed + ' tests passed');
