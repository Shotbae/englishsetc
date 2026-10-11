/***********************************************************************
 * 영어 영작 세특 자동 생성 (Gemini API - Final Master Version)
 * 주제: "나의 기업 투자 제안서" 영작 수행평가
 *
 * - M열: 말듣어통합 세특 (고정 텍스트)
 * - N열: 말듣어 바이트 수 (참고용, 실제 계산은 M열 직접 측정)
 * - O열: 작문 내용
 * - P열: 사용한 단어 수
 * - Q열: 목표 언어 사용하기 (점수, 만점 3)
 * - R열: 문법과 표기법 (점수, 만점 2)
 * - S열: 문장 수 (문법성 판단 보조 기준)
 * - T열: 세특기입수준 (부정행위/미제출/참가함 등)
 * - U열: 작문세특내용 (← 통합된 최종 결과 출력)
 * - V열: 작문세특 바이트 수 (← 숫자만 출력)
 ***********************************************************************/

/* ====================== 사용자 환경 설정 ====================== */
const CONFIG = {
  COL_MALDDEUT_TEXT: 'M',  // 말듣어통합 세특
  COL_MALDDEUT_BYTE: 'N',  // 말듣어 바이트 수
  COL_WRITING:       'O',  // 작문 내용
  COL_WORDS:         'P',  // 사용한 단어 수
  COL_TARGET:        'Q',  // 목표 언어 사용하기 점수
  COL_GRAMMAR:       'R',  // 문법과 표기법 점수
  COL_SENTENCE:      'S',  // 문장 수
  COL_LEVEL:         'T',  // 세특기입수준
  COL_OUTPUT:        'U',  // 세특 결과 출력 열 (녹색 열)
  COL_BYTES:         'V',  // 세특 바이트 수 출력 열

  TARGET_MAX:    3,    
  GRAMMAR_MAX:   2,    

  KOREAN_BYTE:   3,

  DEFAULT_MODEL: 'gemini-3.5-flash', 
  DEFAULT_PROVIDER: 'gemini',  
  
  API_SLEEP_MS:  500, // 순간 폭주 방지 0.5초 대기        

  TEMPERATURE_DRAFT:  0.75,   
  MAX_OUTPUT_TOKENS:  3000,   

  WORD_TIER_TOP3:   300,      
  WORD_TIER_TOP2:   200,      
  WORD_TIER_TOP1:   170,      
  WORD_RATIO_HIGH:  150,      
  WORD_BASE:        100       
};

/* ====================== 프로바이더 정의 ====================== */
const PROVIDERS = {
  gemini: {
    label: 'Gemini (Google)',
    keyProp: 'GEMINI_API_KEY', modelProp: 'GEMINI_MODEL',
    defaultModel: 'gemini-3.5-flash'
  },
  openai: {
    label: 'OpenAI (GPT)',
    keyProp: 'OPENAI_API_KEY', modelProp: 'OPENAI_MODEL',
    defaultModel: 'gpt-4o'
  },
  anthropic: {
    label: 'Claude (Anthropic)',
    keyProp: 'ANTHROPIC_API_KEY', modelProp: 'ANTHROPIC_MODEL',
    defaultModel: 'claude-sonnet-4-5'
  }
};

function getProvider() {
  return PropertiesService.getDocumentProperties().getProperty('PROVIDER') || CONFIG.DEFAULT_PROVIDER;
}
function providerKey(provider) {
  return PropertiesService.getDocumentProperties().getProperty(PROVIDERS[provider].keyProp) || '';
}
function providerModel(provider) {
  return PropertiesService.getDocumentProperties().getProperty(PROVIDERS[provider].modelProp) || PROVIDERS[provider].defaultModel;
}

/* ====================== 메뉴 생성 ====================== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📝 영작 세특 생성')
    .addItem('① 선택한 범위로 세특 생성 (초고속)', 'runOnSelection')
    .addItem('② 범위를 입력해 세특 생성 (예: U3:U10)', 'runOnRangeInput')
    .addSeparator()
    .addItem('⚙️ 설정 (API 키 / 모델)', 'showSettings')
    .addItem('ℹ️ 현재 설정 확인', 'showCurrentSettings')
    .addToUi();
}

/* ====================== 설정 다이얼로그 ====================== */
function showSettings() {
  const p = PropertiesService.getDocumentProperties();
  const current = getProvider();
  const krByte   = p.getProperty('KOREAN_BYTE') || CONFIG.KOREAN_BYTE;

  const state = {};
  Object.keys(PROVIDERS).forEach(function (pv) {
    const k = providerKey(pv);
    state[pv] = {
      label: PROVIDERS[pv].label,
      hasKey: !!k,
      masked: k ? (k.slice(0, 6) + '••••' + k.slice(-4)) : '',
      model: providerModel(pv)
    };
  });
  const stateJson = JSON.stringify(state);

  let providerOptions = '';
  Object.keys(PROVIDERS).forEach(function (pv) {
    providerOptions += '<option value="' + pv + '"' + (pv === current ? ' selected' : '') + '>' + PROVIDERS[pv].label + '</option>';
  });

  const html = [
    '<div style="font-family:Apple SD Gothic Neo,Malgun Gothic,sans-serif;font-size:13px;padding:6px;">',
    '  <div style="margin-bottom:12px;">',
    '    <label style="font-weight:600;">프로바이더</label><br>',
    '    <select id="provider" style="width:100%;box-sizing:border-box;padding:6px;">' + providerOptions + '</select>',
    '  </div>',
    '  <div style="margin-bottom:12px;">',
    '    <label style="font-weight:600;">API 키</label><br>',
    '    <input id="apiKey" type="text" style="width:100%;box-sizing:border-box;padding:6px;">',
    '    <div id="keyHint" style="color:#888;font-size:11px;margin-top:2px;"></div>',
    '  </div>',
    '  <div style="margin-bottom:12px;">',
    '    <label style="font-weight:600;">모델 선택</label>',
    '    <button onclick="loadModels()" style="float:right;font-size:11px;padding:2px 8px;cursor:pointer;">API 키로 모델 목록 불러오기</button>',
    '    <br>',
    '    <select id="model" style="width:100%;box-sizing:border-box;padding:6px;margin-top:4px;"></select>',
    '  </div>',
    '  <div style="margin-bottom:14px;">',
    '    <label style="font-weight:600;">한글 1자당 바이트</label><br>',
    '    <input id="krByte" type="number" value="' + krByte + '" style="width:100%;box-sizing:border-box;padding:6px;">',
    '  </div>',
    '  <button onclick="save()" style="padding:8px 16px;background:#1a73e8;color:#fff;border:none;border-radius:4px;cursor:pointer;">저장</button>',
    '  <span id="status" style="margin-left:10px;color:#188038;"></span>',
    '</div>',
    '<script>',
    '  var STATE = ' + stateJson + ';',
    '  function setStatus(t,c){ var s=document.getElementById("status"); s.textContent=t; s.style.color=c||"#188038"; }',
    '  function curProvider(){ return document.getElementById("provider").value; }',
    '  function refreshProvider(){',
    '    var pv = curProvider();',
    '    var st = STATE[pv] || {};',
    '    document.getElementById("apiKey").value = "";',
    '    document.getElementById("apiKey").placeholder = st.hasKey ? ("저장됨: " + st.masked + " (변경 시에만 입력)") : "API 키 입력";',
    '    document.getElementById("keyHint").textContent = st.hasKey ? "비워두고 저장하면 기존 키가 유지됩니다." : "이 프로바이더의 키가 아직 없습니다.";',
    '    var sel = document.getElementById("model");',
    '    sel.innerHTML = "";',
    '    if(pv === "gemini") {',
    '      sel.innerHTML = "<option value=\\"gemini-3.5-flash\\">gemini-3.5-flash (Flash: 초고속 / 가성비)</option>" + ',
    '                      "<option value=\\"gemini-3.1-pro-preview\\">gemini-3.1-pro-preview (Pro: 고품질 / 논리력 우수)</option>";',
    '      if (st.model && (st.model === "gemini-3.5-flash" || st.model === "gemini-3.1-pro-preview")) {',
    '          sel.value = st.model;',
    '      } else if (st.model) {',
    '          sel.innerHTML += "<option value=\\"" + st.model + "\\">" + st.model + "</option>";',
    '          sel.value = st.model;',
    '      }',
    '    } else {',
    '      var o = document.createElement("option");',
    '      o.value = st.model || ""; o.textContent = st.model || "(모델 불러오기 필요)"; o.selected = true;',
    '      sel.appendChild(o);',
    '    }',
    '    sel.setAttribute("data-current", st.model || "");',
    '  }',
    '  function loadModels(){',
    '    var pv = curProvider();',
    '    var key = document.getElementById("apiKey").value.trim();',
    '    setStatus("모델 목록 불러오는 중...", "#666");',
    '    google.script.run',
    '      .withSuccessHandler(populate)',
    '      .withFailureHandler(function(e){ setStatus("불러오기 실패: " + e.message, "#d93025"); })',
    '      .listModels(key, pv);',
    '  }',
    '  function populate(models){',
    '    var sel = document.getElementById("model");',
    '    var currentModel = sel.getAttribute("data-current") || "";',
    '    sel.innerHTML = "";',
    '    var found = false;',
    '    models.forEach(function(m){',
    '      var o = document.createElement("option");',
    '      o.value = m; o.textContent = m;',
    '      if (m === currentModel){ o.selected = true; found = true; }',
    '      sel.appendChild(o);',
    '    });',
    '    if (!found && currentModel){',
    '      var o = document.createElement("option");',
    '      o.value = currentModel; o.textContent = currentModel + " (현재)";',
    '      o.selected = true;',
    '      sel.insertBefore(o, sel.firstChild);',
    '    }',
    '    setStatus(models.length + "개 모델 불러옴", "#188038");',
    '  }',
    '  function save(){',
    '    var data = {',
    '      provider: curProvider(),',
    '      apiKey: document.getElementById("apiKey").value.trim(),',
    '      model: document.getElementById("model").value.trim(),',
    '      krByte: document.getElementById("krByte").value',
    '    };',
    '    setStatus("저장 중...", "#666");',
    '    google.script.run',
    '      .withSuccessHandler(function(){ setStatus("저장 완료!"); setTimeout(google.script.host.close, 700); })',
    '      .withFailureHandler(function(e){ setStatus("오류: " + e.message, "#d93025"); })',
    '      .saveSettings(data);',
    '  }',
    '  document.getElementById("provider").addEventListener("change", function(){ refreshProvider(); });',
    '  window.onload = function(){',
    '    refreshProvider();',
    '    var pv = curProvider();',
    '    if (STATE[pv] && STATE[pv].hasKey) loadModels();',
    '  };',
    '</scr' + 'ipt>'
  ].join('\n');

  const out = HtmlService.createHtmlOutput(html).setWidth(400).setHeight(540);
  SpreadsheetApp.getUi().showModalDialog(out, '세특 생성 설정');
}

function saveSettings(data) {
  const p = PropertiesService.getDocumentProperties();
  const pv = (data.provider && PROVIDERS[data.provider]) ? data.provider : CONFIG.DEFAULT_PROVIDER;
  p.setProperty('PROVIDER', pv);
  if (data.apiKey) p.setProperty(PROVIDERS[pv].keyProp, data.apiKey);
  if (data.model)  p.setProperty(PROVIDERS[pv].modelProp, data.model);
  if (data.krByte) p.setProperty('KOREAN_BYTE', String(parseInt(data.krByte, 10)));
}

function listModels(apiKeyInput, provider) {
  provider = (provider && PROVIDERS[provider]) ? provider : getProvider();
  const key = (apiKeyInput && apiKeyInput.trim()) ? apiKeyInput.trim() : providerKey(provider);
  if (!key) throw new Error('API 키를 먼저 입력하세요.');

  if (provider === 'gemini')    return listModelsGemini(key);
  if (provider === 'openai')    return listModelsOpenAI(key);
  if (provider === 'anthropic') return listModelsAnthropic(key);
  throw new Error('알 수 없는 프로바이더: ' + provider);
}

function listModelsGemini(key) {
  const resp = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models', {
    method: 'get', headers: { 'x-goog-api-key': key }, muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error('조회 실패 ' + resp.getResponseCode() + ' - ' + resp.getContentText().slice(0, 200));
  const json = JSON.parse(resp.getContentText());
  const models = (json.models || [])
    .filter(function (m) { return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0; })
    .map(function (m) { return m.name.replace(/^models\//, ''); })
    .sort();
  if (!models.length) throw new Error('사용 가능한 모델이 없습니다.');
  return models;
}

function listModelsOpenAI(key) {
  const resp = UrlFetchApp.fetch('https://api.openai.com/v1/models', {
    method: 'get', headers: { 'Authorization': 'Bearer ' + key }, muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error('조회 실패 ' + resp.getResponseCode() + ' - ' + resp.getContentText().slice(0, 200));
  const json = JSON.parse(resp.getContentText());
  const models = (json.data || [])
    .map(function (m) { return m.id; })
    .filter(function (id) { return /gpt|^o\d|chatgpt/.test(id); })
    .sort();
  if (!models.length) throw new Error('사용 가능한 모델이 없습니다.');
  return models;
}

function listModelsAnthropic(key) {
  const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/models', {
    method: 'get',
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error('조회 실패 ' + resp.getResponseCode() + ' - ' + resp.getContentText().slice(0, 200));
  const json = JSON.parse(resp.getContentText());
  const models = (json.data || [])
    .map(function (m) { return m.id; })
    .sort()
    .reverse(); 
  if (!models.length) throw new Error('사용 가능한 모델이 없습니다.');
  return models;
}

function showCurrentSettings() {
  const p = PropertiesService.getDocumentProperties();
  const pv = getProvider();
  const key = providerKey(pv);
  const masked = key ? key.slice(0, 6) + '••••' + key.slice(-4) : '(없음)';
  const msg =
    '프로바이더: ' + PROVIDERS[pv].label + '\n' +
    'API 키: ' + masked + '\n' +
    '모델: ' + providerModel(pv) + '\n' +
    '한글 1자 바이트: ' + (p.getProperty('KOREAN_BYTE') || CONFIG.KOREAN_BYTE);
  SpreadsheetApp.getUi().alert('현재 설정', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ====================== 실행 진입점 ====================== */
function runOnSelection() {
  const range = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  if (!range) { toast('선택된 범위가 없습니다.'); return; }
  const rows = [];
  for (let r = range.getRow(); r < range.getRow() + range.getNumRows(); r++) rows.push(r);
  processRows(rows);
}

function runOnRangeInput() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('세특 생성 범위 입력',
    '세특을 생성할 범위를 입력하세요. (예: U3:U10 또는 3:10)',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const rows = parseRowsFromInput(res.getResponseText().trim());
  if (!rows.length) { toast('범위를 인식할 수 없습니다. 예: U3:U10'); return; }
  processRows(rows);
}

function parseRowsFromInput(text) {
  let m = text.match(/^(\d+)\s*:\s*(\d+)$/);
  if (m) return rangeArray(parseInt(m[1], 10), parseInt(m[2], 10));
  m = text.match(/[A-Za-z]+(\d+)\s*:\s*[A-Za-z]+(\d+)/);
  if (m) return rangeArray(parseInt(m[1], 10), parseInt(m[2], 10));
  m = text.match(/^[A-Za-z]+(\d+)$/);
  if (m) return [parseInt(m[1], 10)];
  return [];
}
function rangeArray(a, b) {
  const start = Math.min(a, b), end = Math.max(a, b), arr = [];
  for (let i = start; i <= end; i++) arr.push(i);
  return arr;
}

/* ====================== 핵심 처리 (초고속 Array Batch 최적화) ====================== */
function processRows(rows) {
  const p = PropertiesService.getDocumentProperties();
  const provider = getProvider();
  const apiKey = providerKey(provider);
  if (!apiKey) {
    SpreadsheetApp.getUi().alert('먼저 [⚙️ 설정]에서 ' + PROVIDERS[provider].label + ' API 키를 입력하세요.');
    return;
  }
  const krByte = parseInt(p.getProperty('KOREAN_BYTE') || CONFIG.KOREAN_BYTE, 10);
  const sheet = SpreadsheetApp.getActiveSheet();

  const cM = colToNum(CONFIG.COL_MALDDEUT_TEXT);
  const cN = colToNum(CONFIG.COL_MALDDEUT_BYTE);
  const cO = colToNum(CONFIG.COL_WRITING);
  const cP = colToNum(CONFIG.COL_WORDS);
  const cQ = colToNum(CONFIG.COL_TARGET);
  const cR = colToNum(CONFIG.COL_GRAMMAR);
  const cS = colToNum(CONFIG.COL_SENTENCE);
  const cT = colToNum(CONFIG.COL_LEVEL);
  const cU = colToNum(CONFIG.COL_OUTPUT);
  const cV = colToNum(CONFIG.COL_BYTES);

  const minRow = Math.min(...rows);
  const maxRow = Math.max(...rows);
  const numRowsToFetch = maxRow - minRow + 1;
  const numColsToFetch = cV - cM + 1;
  
  const dataRange = sheet.getRange(minRow, cM, numRowsToFetch, numColsToFetch);
  const dataValues = dataRange.getValues();

  // 열 인덱스 매핑
  const idxM = cM - cM; 
  const idxN = cN - cM;
  const idxO = cO - cM;
  const idxP = cP - cM;
  const idxQ = cQ - cM;
  const idxR = cR - cM;
  const idxS = cS - cM;
  const idxT = cT - cM;
  const idxU = cU - cM;
  const idxV = cV - cM;

  let done = 0, skipped = 0, errors = 0;
  const errLog = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (row < 2) { skipped++; continue; } 

    const rIdx = row - minRow; 

    const malddeutText = String(dataValues[rIdx][idxM] || '').trim();
    const writing  = String(dataValues[rIdx][idxO] || '').trim();
    const words    = String(dataValues[rIdx][idxP] || '').trim();
    const target   = String(dataValues[rIdx][idxQ] || '').trim();
    const grammar  = String(dataValues[rIdx][idxR] || '').trim();
    const sentence = String(dataValues[rIdx][idxS] || '').trim();
    const level    = String(dataValues[rIdx][idxT] || '').trim();

    toast((i + 1) + '/' + rows.length + '행(' + row + ') 처리 중...');

    if (/부정행위|미제출/.test(level)) {
      dataValues[rIdx][idxU] = '';
      dataValues[rIdx][idxV] = 0;
      done++;
      continue;
    }
    if (/참가함/.test(level)) {
      const msg = '나의 기업 투자 제안서 영작 수행평가에 참가함.';
      dataValues[rIdx][idxU] = msg;
      dataValues[rIdx][idxV] = byteLength(msg, krByte);
      done++;
      continue;
    }

    if (!writing) { skipped++; continue; }

    // ★ 안전성 대폭 강화: N열의 숫자를 믿지 않고, M열의 말듣어 텍스트 바이트를 스크립트가 직접 계산합니다.
    let mByte = byteLength(malddeutText, krByte);
    let spaceByte = malddeutText ? 1 : 0; 
    
    // AI가 쓸 수 있는 순수 영작 세특의 최대 한도 설정
    let studentMaxBytes = 1500 - mByte - spaceByte;
    if (studentMaxBytes < 500) studentMaxBytes = 500;
    if (studentMaxBytes > 1500) studentMaxBytes = 1500;

    const oScore = toNum(target);
    const pScore = toNum(grammar);
    const isPerfect = (oScore >= CONFIG.TARGET_MAX) && (pScore >= CONFIG.GRAMMAR_MAX);
    const wordNum = toNum(words);

    const d = {
      writing: writing, words: words, wordNum: wordNum,
      oScore: oScore, pScore: pScore, sentence: sentence,
      isPerfect: isPerfect, maxBytes: studentMaxBytes, krByte: krByte
    };

    try {
      const aiResult = generateSetuk(d);
      
      // AI 작문 세특 내용 뒤에 띄어쓰기 한 칸과 함께 말듣어 세특 완벽 병합
      // 이제 generateSetuk 내부에서 분량 초과를 완벽히 통제하므로 하드컷팅은 하지 않습니다.
      let finalSetuk = malddeutText ? (aiResult + ' ' + malddeutText) : aiResult;

      dataValues[rIdx][idxU] = finalSetuk;
      dataValues[rIdx][idxV] = byteLength(finalSetuk, krByte);
      done++;
    } catch (e) {
      errors++;
      errLog.push('행 ' + row + ': ' + e.message);
    }

    if (i < rows.length - 1 && CONFIG.API_SLEEP_MS > 0) {
      Utilities.sleep(CONFIG.API_SLEEP_MS);
    }
  }

  dataRange.setValues(dataValues);

  if (errors > 0) {
    let msg = '⚠️ 세특 생성 완료 (단, ' + errors + '건의 빈칸 발생)\n\n구글 서버 순간 과부하로 인해 텍스트가 출력되지 않은 빈칸이 있습니다.\n1~2분 정도 기다리신 후 빈칸 행만 드래그하여 다시 실행해 주세요.\n\n[오류 내역]\n' + errLog.slice(0, 10).join('\n');
    SpreadsheetApp.getUi().alert('결과 안내 (오류 포함)', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('✅ 세특 생성 완료', '모든 세특 생성이 성공적으로 완료되었습니다 (' + done + '건)', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/* ====================== 본문 전개 다양화 패턴 ====================== */
const STYLE_VARIANTS = [
  '본문은 학생이 설정한 사업 아이템의 구체성과 시장성을 먼저 조명한 뒤, 이를 영어로 풀어낸 표현 능력으로 자연스럽게 이어 서술할 것.',
  '본문은 제안서의 논리 전개(강점·기회·위험·대응)를 중심축으로 삼되, 각 지점에서 드러난 영어 표현의 우수성을 근거와 함께 엮어 서술할 것.',
  '본문은 학생의 영어 표현력(어휘·구문·연결·정확성)을 먼저 부각한 뒤, 그것이 사업 아이디어를 얼마나 설득력 있게 전달했는지로 연결하여 서술할 것.',
  '본문은 작문 내용과 영어 능력을 문장마다 교차로 엮어 하나의 매끄러운 흐름으로 서술할 것.',
  '본문은 위험 요소에 대한 학생의 대응 논리와 결론부의 설득 전략을 부각하고, 그 과정에서 드러난 영어 구사력을 근거와 함께 서술할 것.',
  '본문은 학생이 사업 구상을 영어로 구조화하고 조직한 방식(도입-전개-결론의 짜임새)에 주목하여, 내용과 표현을 균형 있게 서술할 것.'
];
function pickStyle() {
  return STYLE_VARIANTS[Math.floor(Math.random() * STYLE_VARIANTS.length)];
}

/* ====================== 단어 수 기반 동적 안내 계산 ====================== */
function wordGuidance(wordNum, maxBytes) {
  var g = { levelLabel: '', densityRule: '', wordCountClause: '', ratioRule: '', poorSafeguard: false, byteGuide: '' };

  if (wordNum >= CONFIG.WORD_TIER_TOP3) {
    g.levelLabel = '최극상';
    g.densityRule = '예외적으로 탁월한 성취 수준임. 학생이 사용한 어휘나 문법적 강점이 왜 우수한지, 어떤 맥락에서 훌륭한지 기존 근거를 한 겹 더 깊이 파고들어 분량을 매우 풍부하게 확보할 것.';
    g.byteGuide = '- 최종 세특은 ' + Math.round(maxBytes * 0.95) + '~' + maxBytes + '바이트(목표의 95~100%) 범위로 한도를 꽉 채워 대단히 풍성하게 작성할 것.';
  } else if (wordNum >= CONFIG.WORD_TIER_TOP2) {
    g.levelLabel = '최상';
    g.densityRule = '대단히 충실한 성취 수준임. 확인된 근거를 바탕으로 최대한 상세하고 풍부하게 풀어서 서술할 것.';
    g.byteGuide = '- 최종 세특은 ' + Math.round(maxBytes * 0.90) + '~' + Math.round(maxBytes * 0.95) + '바이트 범위로 길고 충실하게 작성할 것.';
  } else if (wordNum >= CONFIG.WORD_TIER_TOP1) {
    g.levelLabel = '상상';
    g.densityRule = '매우 충실한 성취 수준임. 확인된 근거를 구체적으로 풀어내어 풍성하게 서술할 것.';
    g.byteGuide = '- 최종 세특은 ' + Math.round(maxBytes * 0.80) + '~' + Math.round(maxBytes * 0.90) + '바이트 범위로 적절히 여유를 두고 작성할 것.';
  } else if (wordNum >= CONFIG.WORD_RATIO_HIGH) {
    g.levelLabel = '상위';
    g.densityRule = '기본 분량을 확실히 넘어 충실하게 작성한 상위 수준임. 근거를 바탕으로 충실히 서술할 것.';
    g.byteGuide = '- 최종 세특은 ' + Math.round(maxBytes * 0.70) + '~' + Math.round(maxBytes * 0.80) + '바이트 범위로 작성할 것.';
  } else if (wordNum >= CONFIG.WORD_BASE) {
    g.levelLabel = '보통';
    g.densityRule = '기본 분량을 충족한 보통 수준임. 내용을 과장하지 말고 학생이 작성한 사실에만 근거해 객관적이고 담백하게 서술할 것. 절대 분량이 적다거나 부족하다는 뉘앙스를 주지 말 것.';
    g.byteGuide = '- 최종 세특은 ' + Math.round(maxBytes * 0.50) + '~' + Math.round(maxBytes * 0.70) + '바이트 범위로 간결하게 작성할 것. 억지로 문장을 늘리지 말 것.';
  } else {
    g.levelLabel = '기본미달';
    g.densityRule = '내용을 과장하지 말고 학생이 작성한 사실에만 근거해 객관적이고 담백하게 서술할 것. 절대 분량이 적다거나 부족하다는 뉘앙스를 주지 말 것.';
    g.byteGuide = '- 최종 세특은 ' + Math.round(maxBytes * 0.30) + '~' + Math.round(maxBytes * 0.50) + '바이트 수준으로 매우 짧고 담백하게 끝낼 것. 절대 말을 지어내어 길이를 늘리지 말 것.';
  }

  if (wordNum >= CONFIG.WORD_TIER_TOP3) {
    g.wordCountClause = '고등학생 수준에서 300단어가 넘는 대단히 방대한 분량을 완성한 점은 매우 어려운 성취이므로, 이 점을 성실성과 작문 능력의 뚜렷한 근거로 반드시 드러낼 것. (단, 단어 수를 언급할 때는 한글 수사 대신 반드시 아라비아 숫자를 사용할 것. 예: 300단어)';
  } else if (wordNum >= CONFIG.WORD_TIER_TOP2) {
    g.wordCountClause = '고등학생 수준에서 200단어가 넘는 매우 풍부한 분량을 완성한 점은 어려운 성취이므로, 이 점을 성실성과 작문 능력의 뚜렷한 근거로 반드시 드러낼 것. (단, 단어 수를 언급할 때는 한글 수사 대신 반드시 아라비아 숫자를 사용할 것. 예: 200단어)';
  } else if (wordNum >= CONFIG.WORD_TIER_TOP1) {
    g.wordCountClause = '고등학생 수준에서 170단어가 넘는 상당한 분량을 작성한 점을 성실성과 작문 능력의 근거로 드러낼 것. (단, 단어 수를 언급할 때는 한글 수사 대신 반드시 아라비아 숫자를 사용할 것. 예: 170단어)';
  } else {
    g.wordCountClause = '사용한 단어 수나 글자 수를 본문에 객관적 수치로 드러내지 말 것(예: 120단어, 백단어 등 금지).';
  }

  if (wordNum >= CONFIG.WORD_RATIO_HIGH) {
    g.ratioRule = '작문 내용 요약 40% : 교사의 평가·근거 있는 칭찬 60%의 비중으로 서술할 것. (단, 교사의 칭찬 근거로 삼을 만한 요소가 모두 소진되었다면, 학생 작문 내용의 구체적인 요약 비중을 유연하게 늘려서라도 제시된 목표 분량(하한선)을 반드시 채울 것.)';
  } else if (wordNum >= CONFIG.WORD_BASE) {
    g.ratioRule = '작문 내용 요약 50% : 교사의 평가·근거 있는 칭찬 50%의 균형으로 서술할 것. (단, 칭찬 근거가 부족하면 내용 요약 비중을 늘려 제시된 하한선 분량을 채울 것.)';
  } else {
    g.ratioRule = '작문 내용 요약을 위주로 하되(내용 60% : 교사 평가 40% 정도), 확인되는 강점이 적으면 억지로 칭찬을 만들지 말고 담백하게 서술할 것.';
    g.poorSafeguard = true;
  }

  return g;
}

/* ====================== 프롬프트 구성 ====================== */
function coreRules(d, g) {
  const perfectRule = d.isPerfect
    ? '목표 언어 사용하기와 문법·표기법이 모두 만점이므로, 어법상 오류가 없다는 취지의 내용을 근거와 함께 자연스럽게 드러내도 됨.'
    : '어법상 오류가 전혀 없다는 식의 단정적 표현은 쓰지 말 것(만점이 아니므로).';

  return [
'[활동 정의]',
'- 이 활동은 "나의 기업 투자 제안서 영작 수행평가"임. 학생은 자신이 직접 창업한다고 가정한 가상의 기업을 설정하고, 그 기업에 대한 투자를 이끌어내는 투자 제안서를 영어로 작성함.',
'- 이것은 AI 사용이 배제된 실제 시험(수행평가)이므로, 학생 본인의 실력을 보여주는 신빙성 높은 지표임. 서술 시 "수행평가" 성격이 드러나게 할 것.',
'',
'[첫 문장 - 반드시 지킬 것]',
'- 세특의 첫 문장은 반드시 다음 골격으로 시작하여 완결할 것: "영작 수행평가에서 자신이 창업한 [기업 설명] 기업에 대한 투자 제안서를 영작함."',
'- [기업 설명]에는 학생 작문에서 파악되는 "어떤 종류의 기업인지"를 넣을 것(예: 온라인 의료 서비스, 청년 부채 문제 해결 등). 구체적 기업명(고유명사)은 절대 쓰지 말 것.',
'- 첫 문장은 위 골격 그대로 쓰되 [기업 설명]만 학생별로 채우며, 마침표로 끝맺고 그다음 문장부터 본문을 전개할 것.',
'',
'[문체]',
'- 모든 문장은 "~함.", "~했음.", "~임.", "~됨.", "~보임." 등 명사형·개조식(음슴체) 종결어미로 끝맺을 것.',
'- 문장 중간 연결에는 "~하여", "~하며", "~으로써" 등을 자유롭게 사용하여 문장 간 흐름을 부드럽게 할 것.',
'- "~습니다/~합니다/~한다/~이다/~였다" 종결어미는 절대 쓰지 말 것.',
'',
'[명칭 통일 및 동일 명칭 반복 금지 - 매우 중요]',
'- 첫 문장에서 이미 "영작 수행평가에서"라고 활동을 밝혔으므로, 이후 이어지는 본문에서는 "영작 수행평가에서", "이 평가에서", "이번 수행평가에서" 등의 명칭을 절대 반복해서 쓰지 말 것.',
'- 오직 학생의 작문 내용과 영어 표현력에 대한 칭찬 및 서술에만 집중할 것.',
'',
'[문장부호 - 반드시 지킬 것]',
'- 본문에는 쉼표(,)와 온점(.) 두 가지 문장부호만 사용할 것. 그 외 특수기호 절대 금지.',
'- 나열할 때는 가운뎃점이나 슬래시 대신 반드시 쉼표로 연결할 것. 예: "강점, 기회, 위험" (O).',
'',
'[근거 기반 서술 - 매우 중요, 위반 금지]',
'- 우수하다는 평가에는 반드시 구체적 근거를 붙일 것. 근거 없는 막연한 칭찬은 신뢰를 떨어뜨리므로 지양할 것.',
'- 근거의 출처는 오직 아래 두 가지로 한정함: (1) 제공된 점수(목표 언어 점수, 문법 점수), (2) 학생의 실제 작문 원문을 직접 분석해 확인되는 사실.',
'- 작문 원문에서 실제로 확인되지 않는 능력·표현·문법 요소는 절대 지어내어 서술하지 말 것.',
'- ' + perfectRule,
'',
'[비율]',
'- ' + g.ratioRule,
(g.poorSafeguard
  ? '- 진실성을 비율 규칙보다 우선함. 작문이 빈약하면 없는 강점을 만들지 말고, 확인되는 만큼만 사실대로 서술하며 제시된 분량을 모두 채우지 않아도 됨.'
  : '- 다만 작문에 비해 근거가 부족하면 비율에 얽매이지 말고 진실성을 우선할 것. 없는 강점을 지어내지 말 것.'),
'',
'[분량 근거]',
'- ' + g.wordCountClause,
'',
'[기타 규칙]',
'- 문단을 나누지 말고 줄바꿈 없이 하나의 문단으로 작성할 것.',
'- 한국어로만 작성하고 영어 단어·문장을 본문에 넣지 말 것.',
'',
'[부정적 서술 절대 금지 - 가장 높은 우선순위 엄격 규칙]',
'- 학생의 단점, 부족한 점, 아쉬운 점, 한계, 미흡함, 짧은 분량 등 부정적이거나 타협하는 서술은 어떤 형태로도 절대 하지 말 것.',
'- 다음과 같은 표현은 예외 없이 금지함: "비록 ~하지만", "다만", "아쉬움", "아쉬운", "미흡", "부족", "제한이 있", "발전의 여지", "개선이 필요", "다소 짧", "다소 간결", "완성도 면에서", "정확성이 일부", "보통 수준", "그치", "머무름".',
'- 문장을 억지로 늘리기 위해 "비록 분량은 적지만~", "비록 간결하지만~" 식의 병 주고 약 주기(대비 구조)를 절대 쓰지 말 것. 못 쓴 학생이라면 그냥 팩트만 한두 줄 쓰고 짧게 끝낼 것.',
'',
'[사전 안내된 형식 구조는 성취로 평가 금지]',
'- 미리 정해진 형식 구조를 갖추었다는 서술은 하지 말 것: "도입-전개-결론", "서론-본론-결론", "강점-기회-위험-대응" 등.'
  ].join('\n');
}

function buildPrompt(d) {
  const g = wordGuidance(d.wordNum, d.maxBytes); 

  return [
'당신은 고등학교 2학년 영어 교과의 세부능력 및 특기사항(세특)을 작성하는 전문 교사입니다.',
'학생의 영어 작문과 평가 정보를 바탕으로, 아래 규칙을 모두 지켜 세특을 한국어로 작성하세요.',
'',
coreRules(d, g),
'',
'[세특 수준]',
'- 이 학생의 성취 수준: ' + g.levelLabel,
'- ' + g.densityRule,
'',
'[본문 전개 방향]',
'- ' + pickStyle(),
'- 첫 문장(골격)과 본문이 하나의 유기적 흐름이 되도록 자연스럽게 연결할 것.',
'',
'[분량 안내]',
'- 한글 1자=' + d.krByte + '바이트, 영문/숫자/공백 1바이트로 계산함.',
g.byteGuide, 
'',
'[학생 평가 정보]',
'- 사용한 단어 수: ' + (d.words || '미상'),
'- 목표 언어 사용하기 점수: ' + d.oScore + ' / ' + CONFIG.TARGET_MAX + ' (전치사+관계대명사, 현재완료 진행형 사용의 어법 정확성)',
'- 문법과 표기법 점수: ' + d.pScore + ' / ' + CONFIG.GRAMMAR_MAX,
'- 문장 수: ' + (d.sentence || '미상'),
'',
'[학생 작문 원문]',
d.writing,
'',
'위 정보를 종합하여 규칙을 모두 지켜, 다른 머리말이나 설명 없이 세특 본문만 한 문단으로 출력하세요. 첫 문장은 반드시 지정된 골격으로 시작하고, 모든 문장은 개조식(음슴체) 종결어미로 끝나야 합니다.'
  ].join('\n');
}

/* ====================== Gemini 호출 ====================== */
function generateSetuk(d) {
  const maxBytes = d.maxBytes, krByte = d.krByte;

  let result = cleanup(callLLM(buildPrompt(d), {
    temperature: CONFIG.TEMPERATURE_DRAFT, thinking: false
  }), krByte, d.wordNum);

  // [1] 첫 문장 안전장치 유지 (재생성)
  if (!firstSentenceOK(result)) {
    const emph = buildPrompt(d) +
      '\n\n[재강조] 첫 문장은 예외 없이 "영작 수행평가에서 자신이 창업한 [기업 설명] 기업에 대한 투자 제안서를 영작함." 골격으로 시작해야 함.';
    const retry = cleanup(callLLM(emph, {
      temperature: CONFIG.TEMPERATURE_DRAFT, thinking: false
    }), krByte, d.wordNum);
    if (firstSentenceOK(retry)) result = retry;
  }

  // [2] ★ 분량 초과 안전장치 (선생님 제안: 줄여서 재시도)
  if (byteLength(result, krByte) > maxBytes) {
    const lengthEmph = buildPrompt(d) +
      '\n\n[길이 초과 엄중 경고] 작성한 내용이 목표 바이트(' + maxBytes + '바이트)를 초과했습니다! 내용을 압축하여 반드시 분량 내로 다시 작성하세요.';
    result = cleanup(callLLM(lengthEmph, {
      temperature: CONFIG.TEMPERATURE_DRAFT, thinking: false
    }), krByte, d.wordNum);
  }

  // [3] ★ 최후의 수단 (스마트 컷팅: 문장 중간이 아닌 마침표에서 깔끔하게 절삭)
  if (byteLength(result, krByte) > maxBytes) {
    result = trimDangling(truncateToBytes(result, maxBytes, krByte));
  }

  return result;
}

function firstSentenceOK(text) {
  const s = String(text).trim();
  if (!s) return false;
  const dot = s.indexOf('.');
  const first = dot >= 0 ? s.substring(0, dot) : s;
  return /영작\s*수행평가/.test(first) && /(제안서|투자)/.test(first);
}

/* ====================== 후처리 정리 ====================== */
function cleanup(text, krByte, wordNum) {
  let s = normalizeName(String(text));
  s = stripPunctuation(s);
  if (wordNum < CONFIG.WORD_TIER_TOP1) s = maskWordCount(s); 
  s = trimDangling(s);
  return s.trim();
}

function normalizeName(str) {
  let s = String(str);
  s = s.replace(/영작\s*수행\s*평가/g, '영작 수행평가');
  s = s.replace(/영문\s*에세이\s*평가|영어\s*에세이\s*평가|영작\s*에세이\s*평가|영문\s*평가|영작문\s*평가/g, '영작 수행평가');
  s = s.replace(/영작\s*평가/g, '영작 수행평가');
  s = s.replace(/영작 수행평가\s*수행평가/g, '영작 수행평가');
  return s;
}

function trimDangling(str) {
  const s = String(str).trim();
  if (!s) return s;
  let lastGood = -1;
  for (let i = 0; i < s.length; i++) {
    const code = s.codePointAt(i);
    if (code >= 0xAC00 && code <= 0xD7A3 && ((code - 0xAC00) % 28) === 16) {
      let j = i + 1;
      while (j < s.length && s[j] === ' ') j++;
      if (s[j] === '.') lastGood = j + 1;
    }
  }
  if (lastGood > 0 && lastGood < s.length) return s.substring(0, lastGood).trim();
  return s;
}

function stripPunctuation(str) {
  let s = String(str);
  s = s.replace(/[·・‧∙•/\\\-–—―~]+/g, ',');
  s = s.replace(/[^가-힣ㄱ-ㅎㅏ-ㅣA-Za-z0-9,.\s]/g, ' ');
  s = s.replace(/\s+([,.])/g, '$1');
  s = s.replace(/,\s*,/g, ',').replace(/,{2,}/g, ',');
  s = s.replace(/,\s*\./g, '.');
  s = s.replace(/\.\s*\./g, '.').replace(/\.{2,}/g, '.');
  s = s.replace(/,(?=[가-힣A-Za-z])/g, ', ');
  s = s.replace(/\s{2,}/g, ' ');
  s = s.replace(/[,\s]+$/g, '');
  return s.trim();
}

function maskWordCount(str) {
  var s = String(str);
  var COUNT = '(?:약\\s*|총\\s*|무려\\s*|자그마치\\s*)?' +
              '(?:[0-9][0-9,]*|[일이삼사오육칠팔구십백천만몇]+)' +
              '\\s*여?\\s*개?\\s*(?:의\\s*)?' +
              '(?:단어|글자|자)';
  var VNOUN = '(?:글|내용|분량|에세이|텍스트|문장|작문|호흡|글자)';
  var LOOKAHEAD = '(?=[^.]{0,15}?' + VNOUN + ')';
  var MODTAIL = '(?:가\\s*넘는|이\\s*넘는|를\\s*넘는|가\\s*넘게|이상|가량의?|정도의?|의|로\\s*구성된|로\\s*이루어진|로\\s*이뤄진|짜리)';

  var reMod = new RegExp(COUNT + '\\s*' + MODTAIL + '\\s*' + LOOKAHEAD, 'g');
  s = s.replace(reMod, '');

  var TAIL2 = '(?:가\\s*넘는|이상|가량|정도|를|로|가|는|으로|의)?';
  var reRep = new RegExp(COUNT + '\\s*' + TAIL2, 'g');
  s = s.replace(reRep, function (m) {
    return /넘|이상|가량|정도/.test(m) ? '상당한 분량으로' : '풍부한 분량으로';
  });

  var reB = new RegExp('(?:단어|글자|자)\\s*(?:약|총|무려)?\\s*(?:[0-9][0-9,]*|[일이삼사오육칠팔구십백천만몇]+)\\s*여?\\s*개?\\s*' + TAIL2, 'g');
  s = s.replace(reB, '풍부한 분량으로');

  s = s.replace(/\s{2,}/g, ' ').replace(/\s+([,.])/g, '$1');
  return s.trim();
}

/* ====================== LLM 호출 ====================== */
function callLLM(promptText, opts) {
  opts = opts || {};
  const temperature = (typeof opts.temperature === 'number') ? opts.temperature : CONFIG.TEMPERATURE_DRAFT;
  const provider = getProvider();
  if (provider === 'openai')    return callOpenAI(promptText, temperature);
  if (provider === 'anthropic') return callAnthropic(promptText, temperature);
  return callGemini(promptText, opts);
}

function callOpenAI(promptText, temperature) {
  const key = providerKey('openai');
  const model = providerModel('openai');
  if (!key) throw new Error('OpenAI API 키가 없습니다.');

  const resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + key },
    payload: JSON.stringify({
      model: model,
      temperature: temperature,
      max_tokens: CONFIG.MAX_OUTPUT_TOKENS,
      messages: [{ role: 'user', content: promptText }]
    }),
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode(), body = resp.getContentText();
  if (code !== 200) throw new Error('OpenAI 응답 ' + code + ' - ' + body.slice(0, 300));
  const json = JSON.parse(body);
  const out = json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content;
  return String(out || '').trim();
}

function callAnthropic(promptText, temperature) {
  const key = providerKey('anthropic');
  const model = providerModel('anthropic');
  if (!key) throw new Error('Claude API 키가 없습니다.');

  function doFetch(includeTemperature) {
    const payload = {
      model: model,
      max_tokens: CONFIG.MAX_OUTPUT_TOKENS,
      messages: [{ role: 'user', content: promptText }]
    };
    if (includeTemperature) payload.temperature = temperature;
    const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    return { code: resp.getResponseCode(), body: resp.getContentText() };
  }

  let r = doFetch(true);
  if (r.code === 400 && /temperature/i.test(r.body) && /deprecat|not.*support|unrecognized|unknown.*param/i.test(r.body)) {
    r = doFetch(false);
  }
  if (r.code !== 200) throw new Error('Claude 응답 ' + r.code + ' - ' + r.body.slice(0, 300));

  const json = JSON.parse(r.body);
  const out = (json.content || [])
    .filter(function (b) { return b.type === 'text'; })
    .map(function (b) { return b.text; })
    .join('');
  return String(out || '').trim();
}

function callGemini(promptText, opts) {
  opts = opts || {};
  const temperature = (typeof opts.temperature === 'number') ? opts.temperature : CONFIG.TEMPERATURE_DRAFT;
  const wantThinking = opts.thinking === true; 

  const apiKey = providerKey('gemini');
  const model  = providerModel('gemini');
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent';

  function doFetch(disableThinking) {
    const gen = { temperature: temperature, maxOutputTokens: 8192 };
    if (disableThinking) gen.thinkingConfig = { thinkingBudget: 0 };
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': apiKey },
      payload: JSON.stringify({
        contents: [{ role: 'user', parts: [{ text: promptText }] }],
        generationConfig: gen
      }),
      muteHttpExceptions: true
    };
    const resp = UrlFetchApp.fetch(url, options);
    return { code: resp.getResponseCode(), body: resp.getContentText() };
  }

  let attempt = 0;
  let maxRetries = 3;
  let r, code, body;

  while (attempt < maxRetries) {
    r = doFetch(!wantThinking);
    code = r.code;
    body = r.body;

    if (code === 400 && !wantThinking && /thinking|thinkingConfig|thinking_budget/i.test(body)) {
      r = doFetch(false);
      code = r.code;
      body = r.body;
    }

    if (code === 200) break; 

    if (code === 503 || code === 429) {
      attempt++;
      if (attempt >= maxRetries) {
        throw new Error('API 서버 과부하: 3회 재시도 실패');
      }
      Utilities.sleep(3000 * attempt);
      continue;
    }
    throw new Error('API 응답 ' + code + ' - ' + body.slice(0, 300));
  }

  const json = JSON.parse(body);
  if (!json.candidates || !json.candidates.length) {
    throw new Error('응답에 candidates 없음: ' + body.slice(0, 300));
  }
  const cand = json.candidates[0];
  const parts = (cand.content && cand.content.parts) || [];
  const out = parts.filter(function (pt) { return pt.text && !pt.thought; })
                   .map(function (pt) { return pt.text; })
                   .join('');
  return out.trim();
}

/* ====================== 바이트 유틸 ====================== */
function charBytes(ch, krByte) {
  const code = ch.codePointAt(0);
  if (code <= 0x7F) return 1;                          
  if (code >= 0xAC00 && code <= 0xD7A3) return krByte; 
  if (code >= 0x1100 && code <= 0x11FF) return krByte; 
  if (code >= 0x3130 && code <= 0x318F) return krByte; 
  return krByte;                                       
}
function byteLength(str, krByte) {
  let b = 0;
  for (const ch of String(str)) b += charBytes(ch, krByte);
  return b;
}
function truncateToBytes(str, maxBytes, krByte) {
  let acc = '', b = 0;
  for (const ch of str) {
    const cb = charBytes(ch, krByte);
    if (b + cb > maxBytes) break;
    acc += ch;
    b += cb;
  }
  return acc.trim();
}

/* ====================== 기타 유틸 ====================== */
function colToNum(letters) {
  let n = 0;
  const s = letters.toUpperCase();
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n;
}
function toNum(v) {
  const m = String(v).replace(/[^0-9.\-]/g, '');
  const n = parseFloat(m);
  return isNaN(n) ? 0 : n;
}
function toast(msg) {
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, '세특 생성', 5);
}
