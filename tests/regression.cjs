// 실행: node tests/regression.cjs — 실제 Google 서비스나 유료 API를 호출하지 않음.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const source = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
new vm.Script(source, {filename: 'Code.gs'});
let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('PASS ' + name); }
  catch (e) { console.error('FAIL ' + name); throw e; }
}
function env(options = {}) {
  const properties = {PROVIDER: 'gemini', GEMINI_API_KEY: 'fake-test-key', GEMINI_MODEL: 'gemini-3.8-flash',
    KOREAN_BYTE: '3', ...options.properties};
  const grid = new Map(), writes = [], alerts = [], requests = [], sleeps = [], cache = new Map(), history = [];
  let selected = [{getRow: () => 2, getNumRows: () => 1}];
  let html = '', now = Date.UTC(2026, 9, 11), released = 0;
  const historySheet = {getSheetId: () => 10,
    getRange: () => ({getValues: () => [history[0]]}),
    appendRow(row) { if (options.backupFails) throw new Error('백업 저장 실패'); history.push(row); }};
  const set = (row, col, value, formula = null) => grid.set(row + ':' + col, {value, formula});
  const get = (row, col) => grid.get(row + ':' + col) || {value: '', formula: null};
  const sheet = {getSheetId: () => 9, getName: () => '학생 자료', getMaxRows: () => 300,
    getRange(row, col, height, width) {
      return {
        getValues: () => Array.from({length: height}, (_, r) =>
          Array.from({length: width}, (_, c) => get(row + r, col + c).value)),
        setValues(values) {
          writes.push({row, col, height, width, values});
          for (let r = 0; r < height; r++) for (let c = 0; c < width; c++) set(row + r, col + c, values[r][c]);
        }
      };
    }};
  const store = {getProperty: key => properties[key] ?? null,
    setProperty(key, value) { properties[key] = value; },
    setProperties(values) { Object.assign(properties, values); },
    deleteProperty(key) { delete properties[key]; }};
  const ui = {ButtonSet: {OK: 'OK'}, alert: (...args) => alerts.push(args), showModalDialog: () => {}};
  class MockDate extends Date {
    constructor(...args) { super(...(args.length ? args : [now])); }
    static now() { return now; }
  }
  const c = {
    Date: MockDate, PropertiesService: {getDocumentProperties: () => store},
    LockService: {getDocumentLock: () => ({tryLock: () => true, releaseLock: () => released++})},
    CacheService: {getScriptCache: () => ({get: key => cache.get(key) || null, put: (key, value) => cache.set(key, value)})},
    SpreadsheetApp: {getUi: () => ui, getActiveSheet: () => sheet, flush: () => {},
      getActiveSpreadsheet: () => ({toast: () => {}, getSheetById: id => id === 9 ? sheet : null,
        getActiveSheet: () => sheet, getActiveRangeList: () => ({getRanges: () => selected}),
        getSheetByName: () => history.length ? historySheet : null,
        insertSheet: () => historySheet, setActiveSheet: () => {}})},
    HtmlService: {createHtmlOutput(value) {
      html = value; return {setWidth() {return this;}, setHeight() {return this;}};
    }},
    Utilities: {sleep(ms) {sleeps.push(ms); now += ms;}, formatDate(date) {return date.toISOString();}},
    UrlFetchApp: {fetch(url, opts) {
      requests.push({url, opts, payload: opts.payload ? JSON.parse(opts.payload) : null});
      const reply = options.fetch ? options.fetch(url, opts, requests.length) : gemini(modelText(opts));
      if (reply.throw) throw new Error(reply.throw);
      return {getResponseCode: () => reply.code ?? 200,
        getContentText: () => reply.body ?? JSON.stringify(reply.json)};
    }}
  };
  vm.createContext(c); vm.runInContext(source, c);
  function student(row, fields = {}) {
    const columns = {M: 13, N: 14, O: 15, P: 16, Q: 17, R: 18, S: 19, T: 20, U: 21, V: 22};
    const defaults = {M: '', N: 0, O: original,
      P: 120, Q: 3, R: 2, S: 6, T: '', U: '이전 세특', V: 99};
    for (const [column, value] of Object.entries({...defaults, ...fields})) set(row, columns[column], value);
  }
  return {c, sheet, properties, grid, writes, alerts, requests, sleeps, cache, history, set, get, student,
    select: rows => { selected = rows.map(row => ({getRow: () => row, getNumRows: () => 1})); },
    advance: ms => {now += ms;}, getHtml: () => html, released: () => released};
}
const first = '영작 수행평가에서 시민과학과 연계한 대기질 관측 연구 계획서를 영작함.';
const goodText = first + ' 시민의 관측 자료를 비교하는 계획을 구체적인 어휘로 설명함.';
const original = 'Citizens collect and compare air quality data for our research project.';
const datum = {writing: original, wordNum: 120, targetScore: 3,
  grammarScore: 2, sentence: 6, contentScores: [], krByte: 3, maxBytes: 1500};
function draft(text = goodText, changes = {}) {
  const body = text.slice(first.length).trim();
  return JSON.stringify({participation_only: false, topic: '대기질 관측', topic_evidence: 'air quality',
    sentences: (body.match(/[^.]+\./g) || []).map(s => ({text: s.trim(), kind: 'content',
      quotes: ['collect and compare air quality data'], score_fields: []})), ...changes});
}
function audit(prompt, supported = true) {
  const claims = JSON.parse(prompt.match(/검사 문장, JSON 배열: (.+)/)[1]);
  return JSON.stringify({checks: claims.map(c => ({index: c.index, supported, reason: supported ? '' : '원문에 없는 주장'}))});
}
function modelText(opts) {
  const payload = JSON.parse(opts.payload);
  const schema = payload.generationConfig?.responseJsonSchema || payload.text?.format?.schema;
  if (!schema) return goodText;
  return schema.properties.checks ? audit(payload.input || payload.contents[0].parts[0].text) : draft();
}
function gemini(text, extra = {}) {return {json: {modelVersion: 'gemini-3.8-flash',
  candidates: [{finishReason: 'STOP', content: {parts: [{text}]}}],
  usageMetadata: {promptTokenCount: 800, candidatesTokenCount: 100, thoughtsTokenCount: 20}, ...extra}};}
const literal = value => JSON.parse(JSON.stringify(value));

test('실제 오류 사례의 단어·조사가 그대로 보존됨', () => {
  const {c} = env();
  const examples = ['글의 흐름이 자연스럽고 메시지가 명확함.', '자원봉사자와 지역 사회를 연결함.',
    '봉사자들이 쉽게 참여함.', '경쟁 업체 대비 자사 서비스의 우수성을 부각함.',
    '부모들의 관심을 기업의 핵심 강점이자 시장의 기회로 제시함.',
    '글로벌 트렌드 분석 능력을 자사의 강점으로 설명함.'];
  for (const text of examples) assert.equal(c.cleanup(text), text);
  assert.equal(c.cleanup('글의\n흐름이 자연스럽고\t명확함.'), '글의 흐름이 자연스럽고 명확함.');
});
test('단어 수 언급은 치환하지 않고 위반으로 검출', () => {
  const {c} = env();
  const text = first + ' 50단어를 작성함.';
  assert.equal(c.cleanup(text), text);
  assert.ok(c.validateSetuk_(text, {...datum, wordNum: 50}, true).includes('170단어 미만 학생의 단어·글자 수 노출'));
  assert.deepEqual(literal(c.validateSetuk_(first + ' 자원봉사자와 자사의 자료를 비교함.', datum, true)), []);
});
test('첫 문장의 활동·연구 주제·완결 골격 검사', () => {
  const {c} = env();
  assert.equal(c.firstSentenceOK(goodText), true);
  assert.equal(c.firstSentenceOK('영작 수행평가 투자 내용을 분석함.'), false);
  assert.equal(c.firstSentenceOK('영작 수행평가에서 시민과학과 연계한 [연구 주제] 연구 계획서를 영작함.'), false);
  assert.deepEqual(literal(c.validateSetuk_(goodText, datum, true)), []);
});
test('완결된 문장만 남기며 마지막 단어를 자르지 않음', () => {
  const {c} = env();
  const next = ' 시민의 관측 자료를 비교하는 구체적인 방법을 설명함.';
  const cap = c.byteLength(first, 3) + 20;
  assert.equal(c.fitCompleteSentences_(first + next, cap, 3), first);
  assert.equal(c.fitCompleteSentences_(first, 20, 3), '');
  assert.equal(c.splitSentences_('측정값 1.5도를 비교함. 연구 계획을 작성함.').sentences.length, 2);
});
test('길이 초과는 완결 문장만 남기고 실제 저장 문장을 검증', () => {
  const {c} = env(); let calls = 0;
  const long = first + ' 관측 자료를 영어로 구체적으로 설명함.'.repeat(15);
  c.callLLM = (prompt, schema, name) => {calls++; return name === 'setuk_audit' ? audit(prompt) : draft(long);};
  const result = c.generateSetuk({...datum, maxBytes: c.byteLength(first, 3) + 50});
  assert.equal(result, first); assert.equal(calls, 2); assert.ok(result.endsWith('함.'));
});
test('검증 API가 실패하면 확인되지 않은 초안을 저장하지 않음', () => {
  const {c} = env(); let calls = 0;
  c.callLLM = () => { if (++calls === 1) return draft(); throw new Error('검증 연결 오류'); };
  assert.throws(() => c.generateSetuk(datum), /검증 연결 오류/);
});
test('형식이 잘못된 JSON 초안은 한 번 보정 후 근거 검증', () => {
  const {c} = env(); let calls = 0;
  c.callLLM = (prompt, schema, name) => ++calls === 1 ? '연구 계획을 작성함.' : name === 'setuk_audit' ? audit(prompt) : draft();
  assert.equal(c.generateSetuk(datum), goodText); assert.equal(calls, 3);
});
test('첫 문장조차 못 들어가면 불완전한 출력 저장 금지', () => {
  const {c} = env(); c.callLLM = () => draft();
  assert.throws(() => c.generateSetuk({...datum, maxBytes: 20}), /검증 실패/);
});
test('평가 불가능한 원문은 참가 문장 허용, 69단어를 자동 참가 처리하지 않음', () => {
  const {c} = env();
  const participation = vm.runInContext('CONFIG.PARTICIPATION_TEXT', c);
  assert.deepEqual(literal(c.validateSetuk_(participation, datum, true)), []);
  assert.equal(c.wordGuidance(69, 1000).levelLabel, '기본미달');
  assert.equal(c.wordGuidance(300, 1000).levelLabel, '최극상');
});
test('분수 점수와 1.5점 구분, 잘못된 점수 검출', () => {
  const {c} = env();
  assert.equal(c.parseNumber_('2/3'), 2); assert.equal(c.parseNumber_('3/3'), 3);
  assert.equal(c.parseScore_('1.5점', [1, 1.5, 2], 'R'), 1.5);
  assert.equal(c.parseScore_('', [1, 2, 3], 'Q'), null);
  assert.throws(() => c.parseScore_('23', [1, 2, 3], 'Q'), /점수/);
});
test('새 채점표와 과제 반영, 무오류 단정과 과거 목표 구문 제거', () => {
  const {c} = env(); const prompt = c.buildPrompt(datum);
  for (const token of ['시민과학', 'If it had not been for', '복합 관계대명사', '오류 3개 이하', '4~8개', '5문장 이하', '70~99단어']) assert.ok(prompt.includes(token));
  assert.ok(!prompt.includes('현재완료 진행형')); assert.ok(!prompt.includes('기업 투자 제안서'));
  assert.ok(c.validateSetuk_(first + ' 문법 오류 없이 내용을 표현함.', datum, true).length);
  assert.deepEqual(literal(c.validateSetuk_(first + ' 자료 부족에 대한 해결 방안을 서술함.', datum, true)), []);
});
test('전체 수식은 보존하고 U·V만 저장', () => {
  const e = env(); e.student(2); e.set(2, 14, 0, '=LEN(M2)*3');
  e.set(2, 16, 120, '=WORDCOUNT(O2)'); e.set(2, 19, 6, '=SENTENCES(O2)');
  e.c.processRows([2], e.sheet);
  assert.ok(e.writes.length); assert.ok(e.writes.every(w => w.col === 21 && w.width === 2));
  assert.equal(e.get(2, 16).formula, '=WORDCOUNT(O2)');
  assert.equal(e.get(2, 19).formula, '=SENTENCES(O2)');
  assert.equal(e.get(2, 21).value, goodText);
  assert.equal(e.properties.PENDING_SETUK_RUN, undefined);
  assert.equal(e.released(), 1);
});
test('M=1,101바이트에서도 통합 결과는 1,500바이트 이내', () => {
  const e = env(); e.student(2, {M: '가'.repeat(367)});
  e.c.generateSetuk = d => {assert.equal(d.maxBytes, 398); return goodText;};
  e.c.processRows([2]);
  assert.equal(e.get(2, 21).value, goodText + ' ' + '가'.repeat(367));
  assert.ok(e.get(2, 22).value <= 1500);
});
test('M이 지나치게 길면 API 요청 없이 기존 결과 유지', () => {
  const e = env(); e.student(2, {M: '가'.repeat(499)}); e.c.processRows([2]);
  assert.equal(e.requests.length, 0); assert.equal(e.get(2, 21).value, '이전 세특');
  assert.deepEqual(JSON.parse(e.properties.PENDING_SETUK_RUN).rows, [2]);
});
test('부정행위·미제출은 API 설정 없이도 강제 비움, 참가함은 고정 문장', () => {
  const e = env({properties: {PROVIDER: '', GEMINI_API_KEY: ''}});
  e.student(2, {T: '부정행위', M: '기존 말듣'}); e.student(3, {T: '미제출'});
  e.student(4, {T: '참가함', M: '기존 말듣'}); e.c.processRows([2, 3, 4]);
  assert.equal(e.get(2, 21).value, ''); assert.equal(e.get(2, 22).value, 0);
  assert.equal(e.get(3, 21).value, ''); assert.match(e.get(4, 21).value, /영작 수행평가에 참가함\.$/);
  assert.ok(!e.get(4, 21).value.includes('기존 말듣')); assert.equal(e.requests.length, 0);
});
test('원문 빈칸은 기존 결과 보존', () => {
  const e = env(); e.student(2, {O: ''}); e.c.processRows([2]);
  assert.equal(e.get(2, 21).value, '이전 세특'); assert.equal(e.requests.length, 0);
  assert.equal(JSON.parse(e.properties.LAST_SETUK_USAGE).skipped, 1);
});
test('API 빈 본문은 실패로 처리하고 재실행 행 기록', () => {
  const e = env({fetch: () => gemini('')}); e.student(2); e.c.processRows([2]);
  assert.equal(e.get(2, 21).value, '이전 세특'); assert.equal(e.writes.length, 0);
  assert.deepEqual(JSON.parse(e.properties.PENDING_SETUK_RUN).rows, [2]);
  assert.ok(e.alerts.some(args => args.join(' ').includes('빈 본문')));
});
test('대량 처리 중단에도 앞선 행은 저장되며 원래 시트에서 재개', () => {
  const e = env(); e.student(2); e.student(3); e.student(4);
  e.c.generateSetuk = () => {e.advance(241000); return goodText;};
  e.c.processRows([2, 3, 4]); assert.equal(e.get(2, 21).value, goodText);
  assert.equal(e.get(3, 21).value, '이전 세특');
  assert.deepEqual(JSON.parse(e.properties.PENDING_SETUK_RUN), {sheetId: 9, rows: [3, 4], profile: 'gemini'});
  e.c.generateSetuk = () => goodText; e.c.resumePendingRows();
  assert.equal(e.get(3, 21).value, goodText); assert.equal(e.get(4, 21).value, goodText);
  assert.equal(e.properties.PENDING_SETUK_RUN, undefined);
});
test('Gemini 3.8에는 LOW 사고 수준과 공통 토큰 상한 적용', () => {
  const e = env(); e.student(2); e.c.processRows([2]);
  const payload = e.requests[0].payload;
  assert.equal(payload.generationConfig.maxOutputTokens, 3000);
  assert.deepEqual(payload.generationConfig.thinkingConfig, {thinkingLevel: 'LOW'});
  assert.equal(payload.generationConfig.temperature, undefined);
  const stats = JSON.parse(e.properties.LAST_SETUK_USAGE);
  assert.equal(stats.inputTokens, 1600); assert.equal(stats.outputTokens, 240);
  assert.equal(stats.thinkingTokens, 40); assert.equal(stats.groundingChecks, 1);
  assert.deepEqual(stats.actualModels, ['gemini-3.8-flash']);
});
test('Gemini 2.5 Flash의 사고 끄기 요청은 별도로 유지', () => {
  const e = env(); e.c.callGemini('test', {key: 'fake', model: 'gemini-2.5-flash'});
  assert.deepEqual(e.requests[0].payload.generationConfig.thinkingConfig, {thinkingBudget: 0});
});
test('HTTP 429·503은 재시도, 인증 실패는 재시도하지 않음', () => {
  const e = env({fetch: (_, __, n) => n < 3 ? {code: n === 1 ? 429 : 503, json: {error: {message: 'busy'}}} : gemini(goodText)});
  assert.equal(e.c.callGemini('test', {key: 'fake', model: 'gemini-3.8-flash'}), goodText);
  assert.equal(e.requests.length, 3); assert.deepEqual(e.sleeps, [2000, 4000]);
  const denied = env({fetch: () => ({code: 401, json: {error: {message: 'key fake-test-key is invalid'}}})});
  assert.throws(() => denied.c.callGemini('test', {key: 'fake-test-key', model: 'gemini-3.8-flash'}), /API 키 숨김/);
  assert.equal(denied.requests.length, 1);
});
test('MAX_TOKENS로 잘린 API 응답을 성공으로 저장하지 않음', () => {
  const e = env({fetch: () => gemini(goodText, {candidates: [{finishReason: 'MAX_TOKENS', content: {parts: [{text: goodText}]}}]})});
  e.student(2); e.c.processRows([2]); assert.equal(e.writes.length, 0);
  assert.equal(e.get(2, 21).value, '이전 세특');
});
test('인증·할당량 오류는 뒤 학생의 반복 호출을 중단하고 남은 행 보존', () => {
  for (const code of [401, 429]) {
    const e = env({fetch: () => ({code, json: {error: {message: 'request denied'}}})});
    e.student(2); e.student(3); e.c.processRows([2, 3]);
    assert.equal(e.requests.length, code === 401 ? 1 : 3);
    assert.equal(e.get(2, 21).value, '이전 세특'); assert.equal(e.get(3, 21).value, '이전 세특');
    assert.deepEqual(JSON.parse(e.properties.PENDING_SETUK_RUN).rows, [2, 3]);
  }
});
test('Gemini 모델 목록은 다음 페이지와 최신 3.8까지 조회', () => {
  const e = env({fetch: (_, __, n) => ({json: n === 1 ? {models: [
    {name: 'models/gemini-3.5-flash', supportedGenerationMethods: ['generateContent']},
    {name: 'models/gemini-3.8-flash-tts', supportedGenerationMethods: ['generateContent']}], nextPageToken: 'p2'}
    : {models: [{name: 'models/gemini-3.8-flash', supportedGenerationMethods: ['generateContent']}]}})});
  assert.deepEqual(literal(e.c.listModels('fake', 'gemini')), ['gemini-3.8-flash', 'gemini-3.5-flash']);
  assert.ok(e.requests[1].url.includes('pageToken=p2'));
});
test('Claude 모델 목록은 페이지와 은퇴 여부 처리', () => {
  const e = env({fetch: (_, __, n) => ({json: n === 1 ? {data: [{id: 'claude-sonnet-4-5', lifecycle: 'retired'}], has_more: true, last_id: 'p1'}
    : {data: [{id: 'claude-sonnet-5-5', lifecycle: 'active'}], has_more: false}})});
  assert.deepEqual(literal(e.c.listModels('fake', 'anthropic')), ['claude-sonnet-5-5']);
  assert.ok(e.requests[1].url.includes('after_id=p1'));
});
test('OpenAI Responses API의 출력·추론 사용량과 토큰 상한', () => {
  const e = env({properties: {PROVIDER: 'openai', OPENAI_API_KEY: 'fake', OPENAI_MODEL: 'gpt-6-luna'}, fetch: (_, opts) => ({json: {
    model: 'gpt-6-sol', status: 'completed', output: [{type: 'reasoning'},
      {type: 'message', content: [{type: 'output_text', text: modelText(opts)}]}],
    usage: {input_tokens: 500, input_tokens_details: {cached_tokens: 50}, output_tokens: 200,
      output_tokens_details: {reasoning_tokens: 30}}}})});
  e.student(2); e.c.processRows([2], e.sheet, 'sol');
  assert.equal(e.requests[0].url, 'https://api.openai.com/v1/responses');
  assert.equal(e.requests[0].payload.max_output_tokens, 3000); assert.equal(e.requests[0].payload.store, false);
  assert.deepEqual(e.requests[0].payload.reasoning, {effort: 'low'});
  assert.equal(e.get(2, 21).value, goodText);
  const stats = JSON.parse(e.properties.LAST_SETUK_USAGE);
  assert.equal(stats.outputTokens, 400); assert.equal(stats.thinkingTokens, 60);
  assert.equal(e.requests[0].payload.model, 'gpt-6-sol');
  assert.equal(e.requests[0].payload.text.format.type, 'json_schema');
  assert.equal(e.requests[0].payload.text.format.strict, true);
});
test('OpenAI 미완료 응답은 저장하지 않음', () => {
  const e = env({fetch: () => ({json: {status: 'incomplete', incomplete_details: {reason: 'max_output_tokens'}}})});
  assert.throws(() => e.c.callOpenAI('test', {model: 'gpt-6-luna', key: 'fake'}), /max_output_tokens/);
});
test('Claude 사고 끄기 지원 메타데이터와 최신 모델 요청', () => {
  const e = env({properties: {PROVIDER: 'anthropic', ANTHROPIC_API_KEY: 'fake', ANTHROPIC_MODEL: 'claude-sonnet-5-5'},
    fetch: (_, opts) => opts.method === 'get' ? {json: {capabilities: {thinking: {types: {disabled: {supported: true}}}}}}
      : {json: {model: 'claude-sonnet-5-5', stop_reason: 'end_turn', content: [{type: 'text', text: goodText}],
        usage: {input_tokens: 300, cache_read_input_tokens: 20, output_tokens: 100}}}});
  e.c.callAnthropic('test', {key: 'fake', model: 'claude-sonnet-5-5'});
  assert.equal(e.requests[1].payload.max_tokens, 3000);
  assert.deepEqual(e.requests[1].payload.thinking, {type: 'disabled'});
  assert.equal(e.requests[1].payload.temperature, undefined);
});
test('항상 사고가 켜진 Claude는 supported 메타데이터에 따라 LOW로 요청', () => {
  const e = env({fetch: (_, opts) => opts.method === 'get' ? {json: {capabilities: {
    thinking: {types: {disabled: {supported: false}, adaptive: {supported: true}}}, effort: {low: {supported: true}}}}}
    : {json: {stop_reason: 'end_turn', content: [{type: 'text', text: goodText}]}}});
  e.c.callAnthropic('test', {key: 'fake', model: 'claude-opus-5-5'});
  assert.deepEqual(e.requests[1].payload.thinking, {type: 'adaptive'});
  assert.deepEqual(e.requests[1].payload.output_config, {effort: 'low'});
});
test('보고되지 않은 사고 토큰은 0으로 확정하지 않고 별도 표시', () => {
  const e = env({fetch: (_, opts) => gemini(modelText(opts), {usageMetadata: {promptTokenCount: 100, candidatesTokenCount: 80}})});
  e.student(2); e.c.processRows([2]);
  assert.equal(JSON.parse(e.properties.LAST_SETUK_USAGE).unknownThinkingResponses, 2);
  e.c.showLastUsage();
  assert.ok(e.alerts.some(args => args.join(' ').includes('세부 수치를 보고하지 않은 응답: 2회')));
});
test('설정창 스크립트 문법과 상태 JSON 안전성', () => {
  const e = env({properties: {GEMINI_MODEL: '</script><script>alert(1)</script>'}});
  e.c.showSettings(); const html = e.getHtml();
  const script = html.match(/<script>([\s\S]*?)<\/script>/)[1]; new vm.Script(script);
  assert.ok(script.includes('\\u003c/script>')); assert.ok(!script.includes('fake-test-key'));
});
test('설정 저장은 기존 키 유지, 바이트·내용 열 검증, 리셋 후 기본 제공사 복구 안 함', () => {
  const e = env(); e.c.saveSettings({provider: 'gemini', model: 'gemini-3.8-flash', apiKey: '', krByte: 3, contentCols: ''});
  assert.equal(e.properties.GEMINI_API_KEY, 'fake-test-key');
  assert.throws(() => e.c.saveSettings({provider: 'gemini', model: 'gemini-3.8-flash', krByte: 0}), /바이트/);
  assert.throws(() => e.c.parseContentColumns_('U,W,X,Y,Z'), /결과 열/);
  e.c.resetApiSettings(); assert.equal(e.c.getProvider(), ''); assert.equal(e.properties.KOREAN_BYTE, '3');
});
test('범위는 전체 일치와 유효 행만 허용', () => {
  const {c} = env();
  assert.deepEqual(literal(c.parseRowsFromInput('U3:U5')), [3, 4, 5]);
  assert.deepEqual(literal(c.parseRowsFromInput('3')), [3]);
  assert.deepEqual(literal(c.parseRowsFromInput('0:3')), []);
  assert.deepEqual(literal(c.parseRowsFromInput('noise U3:U5 extra')), []);
  assert.deepEqual(literal(c.parseRowsFromInput('1:999999999')), []);
});
function openaiResponse(opts) {
  return {json: {model: 'gpt-6-sol', status: 'completed',
    output: [{type: 'message', content: [{type: 'output_text', text: modelText(opts)}]}],
    usage: {input_tokens: 500, output_tokens: 100, output_tokens_details: {reasoning_tokens: 10}}}};
}
test('기본 메뉴는 저장된 Claude·과거 Gemini 모델과 무관하게 Gemini 3.8 사용', () => {
  const e = env({properties: {PROVIDER: 'anthropic', GEMINI_MODEL: 'gemini-3.5-flash'}});
  e.student(2); e.c.runOnSelection();
  assert.equal(e.requests.length, 2);
  assert.ok(e.requests.every(r => r.url.includes('/gemini-3.8-flash:generateContent')));
  assert.equal(e.properties.PROVIDER, 'anthropic');
  assert.equal(e.properties.GEMINI_MODEL, 'gemini-3.5-flash');
});
test('Sol은 비연속 선택 행만 재작성하며 이전 결과·새 결과를 백업', () => {
  const e = env({properties: {OPENAI_API_KEY: 'fake-sol', OPENAI_MODEL: 'gpt-6-luna'},
    fetch: (_, opts) => openaiResponse(opts)});
  for (const row of [2, 3, 4, 5]) e.student(row);
  e.select([3, 5, 3]); e.c.runSelectedWithSol();
  assert.equal(e.requests.length, 4); assert.ok(e.requests.every(r => r.payload.model === 'gpt-6-sol'));
  assert.equal(e.get(2, 21).value, '이전 세특'); assert.equal(e.get(4, 21).value, '이전 세특');
  assert.equal(e.get(3, 21).value, goodText); assert.equal(e.get(5, 21).value, goodText);
  assert.equal(e.history.length, 3);
  assert.deepEqual(e.history.slice(1).map(r => r[3]), [3, 5]);
  assert.equal(e.history[1][5], '이전 세특'); assert.equal(e.history[1][7], goodText);
  assert.equal(e.properties.PROVIDER, 'gemini'); assert.equal(e.properties.OPENAI_MODEL, 'gpt-6-luna');
  // AI가 평가 불가능한 제출물로 판정한 참가 결과도 재작성 백업에 포함.
  e.student(6, {O: 'hello', P: 1});
  const participation = vm.runInContext('CONFIG.PARTICIPATION_TEXT', e.c);
  e.c.generateSetuk = () => participation; e.c.processRows([6], e.sheet, 'sol');
  assert.equal(e.history[3][7], participation);
});
test('Sol 재작성 후 기본 메뉴를 누르면 다시 Gemini 3.8 사용', () => {
  const e = env({properties: {PROVIDER: 'openai', OPENAI_API_KEY: 'fake-sol'},
    fetch: (url, opts) => url.includes('openai') ? openaiResponse(opts) : gemini(modelText(opts))});
  e.student(2); e.student(3); e.select([2]); e.c.runSelectedWithSol();
  e.select([3]); e.c.runOnSelection();
  assert.ok(e.requests.slice(2).every(r => r.url.includes('gemini-3.8-flash')));
});
test('Sol 작업을 중단·재개해도 다른 모델로 바뀌지 않음', () => {
  const e = env({properties: {OPENAI_API_KEY: 'fake-sol'}, fetch: (_, opts) => openaiResponse(opts)});
  e.student(2); e.student(3);
  e.c.generateSetuk = () => {e.advance(241000); return goodText;};
  e.c.processRows([2, 3], e.sheet, 'sol');
  assert.deepEqual(JSON.parse(e.properties.PENDING_SETUK_RUN), {sheetId: 9, rows: [3], profile: 'sol'});
  e.properties.PROVIDER = 'gemini';
  // 새 실행의 컨텍스트를 확인하는 생성 함수를 사용; 실제 API는 위 메뉴 테스트에서 검증.
  e.c.generateSetuk = () => {
    assert.equal(vm.runInContext('RUN_CONTEXT_.model', e.c), 'gpt-6-sol'); return goodText;
  };
  e.c.resumePendingRows(); assert.equal(e.get(3, 21).value, goodText);
  assert.equal(e.properties.PENDING_SETUK_RUN, undefined);
});
test('OpenAI 키가 없으면 Gemini로 몰래 대체하지 않고 기존 결과 보존', () => {
  const e = env(); e.student(2); e.c.processRows([2], e.sheet, 'sol');
  assert.equal(e.requests.length, 0); assert.equal(e.writes.length, 0);
  assert.equal(e.get(2, 21).value, '이전 세특'); assert.equal(e.history.length, 0);
  assert.ok(e.alerts.some(a => a.join(' ').includes('GPT-6 Sol용 API 키')));
});
test('Gemini와 Sol의 작성 프롬프트·검증 프롬프트·문체 규칙은 동일', () => {
  const g = env(), s = env({properties: {OPENAI_API_KEY: 'fake-sol'}, fetch: (_, opts) => openaiResponse(opts)});
  g.student(2); s.student(2); g.c.processRows([2]); s.c.processRows([2], s.sheet, 'sol');
  for (let i = 0; i < 2; i++) {
    assert.equal(g.requests[i].payload.contents[0].parts[0].text, s.requests[i].payload.input);
    assert.equal(g.requests[i].payload.systemInstruction.parts[0].text, s.requests[i].payload.instructions);
    assert.deepEqual(g.requests[i].payload.generationConfig.responseJsonSchema, s.requests[i].payload.text.format.schema);
  }
  assert.equal(g.get(2, 21).value, s.get(2, 21).value);
  assert.ok(!source.includes('Math.random'));
});
test('원문에 없는 인용·미입력 점수·한 항목의 여러 문장을 거부', () => {
  const {c} = env(); const value = JSON.parse(draft());
  value.topic_evidence = 'Citizens use statistical software.';
  assert.throws(() => c.prepareDraft_(JSON.stringify(value), datum), /인용이 원문에 없음/);
  value.topic_evidence = 'air quality'; value.sentences[0].quotes = ['an invented quotation'];
  assert.throws(() => c.prepareDraft_(JSON.stringify(value), datum), /인용이 원문에 없음/);
  value.sentences[0].quotes = ['air quality']; value.sentences[0].score_fields = ['C1'];
  assert.throws(() => c.prepareDraft_(JSON.stringify(value), datum), /미입력 점수/);
  value.sentences[0].score_fields = []; value.sentences[0].text += ' 결과를 발표함.';
  assert.throws(() => c.prepareDraft_(JSON.stringify(value), datum), /완결 문장 하나/);
});
test('실제 원문 인용을 붙였어도 의미 검증에서 근거 없으면 보정', () => {
  const e = env(); e.student(2); let drafts = 0;
  e.c.callLLM = (prompt, schema, name) => {
    if (name === 'setuk_audit') return audit(prompt, drafts > 1);
    drafts++;
    return drafts === 1 ? draft(first + ' 통계 분석을 완료하고 오염 감소 성과를 입증함.') : draft();
  };
  e.c.processRows([2]); assert.equal(e.get(2, 21).value, goodText);
  const stats = JSON.parse(e.properties.LAST_SETUK_USAGE);
  assert.equal(stats.groundingChecks, 2); assert.equal(stats.groundingRejected, 1); assert.equal(stats.repaired, 1);
});
test('근거 없는 초안이 보정 후에도 남으면 U·V를 쓰지 않음', () => {
  const e = env(); e.student(2);
  e.c.callLLM = (prompt, schema, name) => name === 'setuk_audit' ? audit(prompt, false)
    : draft(first + ' 공동 연구를 수행하고 성과를 발표함.');
  e.c.processRows([2]); assert.equal(e.writes.length, 0);
  assert.equal(e.get(2, 21).value, '이전 세특'); assert.equal(e.get(2, 22).value, 99);
  assert.deepEqual(JSON.parse(e.properties.PENDING_SETUK_RUN).rows, [2]);
  assert.ok(e.alerts.some(a => a.join(' ').includes('검증 실패')));
});
test('검증 응답의 빠진 문장·중복 번호·문자열 판정은 저장 거부', () => {
  const invalid = [{checks: [{index: 0, supported: true, reason: ''}]},
    {checks: [{index: 0, supported: true, reason: ''}, {index: 0, supported: true, reason: ''}]},
    {checks: [{index: 0, supported: 'true', reason: ''}, {index: 1, supported: true, reason: ''}]}];
  for (const value of invalid) {
    const e = env(); e.student(2);
    e.c.callLLM = (_, __, name) => name === 'setuk_audit' ? JSON.stringify(value) : draft();
    e.c.processRows([2]); assert.equal(e.writes.length, 0); assert.equal(e.get(2, 21).value, '이전 세특');
  }
});
test('주제만 확인되는 글은 분량을 채우지 않고 첫 문장만 저장 가능', () => {
  const e = env(); e.student(2, {O: 'My research topic is air quality.', P: 7, Q: '', R: '', S: 1});
  e.c.callLLM = (prompt, _, name) => name === 'setuk_audit' ? audit(prompt) : draft(first, {sentences: []});
  e.c.processRows([2]); assert.equal(e.get(2, 21).value, first);
  assert.equal(e.get(2, 22).value, e.c.byteLength(first, 3));
});
test('원문이 충분한데 참가만 기록하려는 초안도 검증 후 보정', () => {
  const {c} = env(); let count = 0;
  c.callLLM = (prompt, _, name) => {
    count++;
    if (name === 'setuk_audit') return audit(prompt, count !== 2);
    return count === 1 ? JSON.stringify({participation_only: true, topic: '', topic_evidence: '', sentences: []}) : draft();
  };
  assert.equal(c.generateSetuk(datum), goodText); assert.equal(count, 4);
});
test('Sol 백업이 실패하면 검증된 새 결과도 기존 결과를 덮어쓰지 않음', () => {
  const e = env({properties: {OPENAI_API_KEY: 'fake-sol'}, backupFails: true, fetch: (_, opts) => openaiResponse(opts)});
  e.student(2); e.c.processRows([2], e.sheet, 'sol');
  assert.equal(e.writes.length, 0); assert.equal(e.get(2, 21).value, '이전 세특');
  assert.deepEqual(JSON.parse(e.properties.PENDING_SETUK_RUN).rows, [2]);
});
test('Sol 메뉴에서도 부정행위·미제출은 백업·API 설정 없이 강제 비움', () => {
  const e = env({properties: {GEMINI_API_KEY: ''}, backupFails: true});
  e.student(2, {T: '부정행위'}); e.student(3, {T: '미제출'}); e.student(4, {T: '참가함'});
  e.c.processRows([2, 3, 4], e.sheet, 'sol');
  assert.equal(e.get(2, 21).value, ''); assert.equal(e.get(3, 21).value, '');
  assert.equal(e.get(2, 22).value, 0); assert.equal(e.requests.length, 0); assert.equal(e.history.length, 0);
  assert.match(e.get(4, 21).value, /영작 수행평가에 참가함\.$/);
});
test('내용보다 언어를 먼저 쓰거나 과장된 칭찬을 붙이면 거부', () => {
  const {c} = env(); const value = JSON.parse(draft());
  value.sentences[0].kind = 'language';
  value.sentences.push({...value.sentences[0], kind: 'content'});
  assert.throws(() => c.prepareDraft_(JSON.stringify(value), datum), /공통 전개 순서/);
  assert.ok(c.validateSetuk_(first + ' 독보적인 연구 역량을 발휘함.', datum, true).length);
});

console.log('\n' + passed + ' regression checks passed. No external API requests made.');
