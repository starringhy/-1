/**
 * 지구마불세계여행 프로젝트 - 최종 차시 실시간 탐구 종합 전광판
 * 다중 사용자 실시간 서버 동기화 (Express REST + WebSocket + Google Apps Script 호환)
 */

const STORAGE_KEY_TEAMS = 'earth_marble_teams_v1';
const STORAGE_KEY_SCRIPT_URL = 'earth_marble_apps_script_url';
const DEFAULT_TEAMS_COUNT = 14;

// State
let teamsData = [];
let currentViewMode = 'input'; // 'input' | 'teacher'
let activeZoomTeamId = null;
let serverAvailable = false;
let wsConnection = null;
let wsReconnectTimer = null;
let pendingConfirmCallback = null;

// Per-card local typing debounce timers & dirty flags so remote updates never break Korean IME typing
const autoSaveTimers = new Map();
const dirtyTeamIds = new Set();

const DEFAULT_ROLES = [
  { id: 'speaker', title: '📢 발표자 (탐구 결과 발표)', cardClass: 'role-card-speaker' },
  { id: 'writer', title: '✍️ 기록자 (문장 정리 및 작성)', cardClass: 'role-card-writer' },
  { id: 'timer', title: '⏱️ 시간지기 (활동 시간 관리)', cardClass: 'role-card-timer' },
  { id: 'organizer', title: '🌟 정리대장 (자료 및 모둠 정리)', cardClass: 'role-card-organizer' }
];

document.addEventListener('DOMContentLoaded', async () => {
  applyUrlParameters();
  initLocalFallbackData();
  renderInputGrid();
  renderTeacherBoardGrid();
  updateStats();
  setupEventListeners();

  // Connect to shared server immediately so all users with the link see the same data
  await fetchLatestFromServer(false);
  connectWebSocket();
  startBackgroundSyncLoop();
});

/**
 * Check URL parameters (?mode=teacher or ?sheet=...)
 */
function applyUrlParameters() {
  try {
    const params = new URLSearchParams(window.location.search);
    const sheetParam = params.get('sheet');
    if (sheetParam && sheetParam.startsWith('https://script.google.com/')) {
      localStorage.setItem(STORAGE_KEY_SCRIPT_URL, sheetParam);
    }
    const modeParam = params.get('mode');
    if (modeParam === 'teacher') {
      setViewMode('teacher', true);
    }
  } catch (e) {
    console.warn('URL parameter check skipped:', e);
  }
}

/**
 * Initialize default 14 teams locally while awaiting server response
 */
function initLocalFallbackData() {
  const saved = localStorage.getItem(STORAGE_KEY_TEAMS);
  if (saved) {
    try {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length === DEFAULT_TEAMS_COUNT) {
        teamsData = parsed;
        return;
      }
    } catch {
      // Ignore corrupt local cache
    }
  }
  teamsData = generateDefaultTeams();
}

function generateDefaultTeams() {
  const list = [];
  for (let i = 1; i <= DEFAULT_TEAMS_COUNT; i++) {
    list.push({
      id: i,
      teamNumber: `${i}모둠`,
      countryName: '',
      sentence: '',
      updatedAt: null
    });
  }
  return list;
}

function saveToLocalStorage() {
  try {
    localStorage.setItem(STORAGE_KEY_TEAMS, JSON.stringify(teamsData));
  } catch {
    // Ignore storage quota issues in private browsing
  }
  updateStats();
}

/**
 * =========================================================================
 * Real-Time Shared Backend Synchronization (/api/teams + WebSocket /ws)
 * =========================================================================
 */

async function fetchLatestFromServer(showFeedback) {
  try {
    const res = await fetch('/api/teams', {
      method: 'GET',
      headers: { 'Accept': 'application/json' },
      cache: 'no-store'
    });

    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const payload = await res.json();
    serverAvailable = true;

    if (payload.appsScriptUrl && !localStorage.getItem(STORAGE_KEY_SCRIPT_URL)) {
      localStorage.setItem(STORAGE_KEY_SCRIPT_URL, payload.appsScriptUrl);
    }

    if (Array.isArray(payload.teams)) {
      applyRemoteTeamsList(payload.teams);
    }

    if (showFeedback) {
      showToast('🔄 모든 모둠의 최신 공유 내용을 동기화했습니다!');
    }
  } catch {
    // Fallback for static GitHub Pages if /api/teams does not exist: try Google Apps Script if configured
    const scriptUrl = localStorage.getItem(STORAGE_KEY_SCRIPT_URL);
    if (scriptUrl) {
      await silentFetchFromAppsScript(scriptUrl, showFeedback);
    } else if (showFeedback) {
      showToast('현재 로컬 모드로 동작 중입니다.');
    }
  }
}

function connectWebSocket() {
  if (wsReconnectTimer) {
    clearTimeout(wsReconnectTimer);
    wsReconnectTimer = null;
  }

  try {
    const protocol = window.location.protocol === 'https:' ? 'wss:' : 'ws:';
    const wsUrl = `${protocol}//${window.location.host}/ws`;
    const ws = new WebSocket(wsUrl);
    wsConnection = ws;

    ws.onopen = () => {
      serverAvailable = true;
    };

    ws.onmessage = (event) => {
      try {
        const data = JSON.parse(event.data);
        if (data.type === 'init' && Array.isArray(data.teams)) {
          if (data.appsScriptUrl) {
            localStorage.setItem(STORAGE_KEY_SCRIPT_URL, data.appsScriptUrl);
          }
          applyRemoteTeamsList(data.teams);
        } else if (data.type === 'team:updated' && data.team) {
          applyRemoteSingleTeam(data.team);
        } else if (data.type === 'teams:synced' && Array.isArray(data.teams)) {
          applyRemoteTeamsList(data.teams);
        } else if (data.type === 'config:updated' && typeof data.appsScriptUrl === 'string') {
          localStorage.setItem(STORAGE_KEY_SCRIPT_URL, data.appsScriptUrl);
        }
      } catch (err) {
        console.warn('WS message parse error:', err);
      }
    };

    ws.onclose = () => {
      wsConnection = null;
      wsReconnectTimer = setTimeout(connectWebSocket, 3000);
    };

    ws.onerror = () => {
      ws.close();
    };
  } catch {
    // Static host without WS support
  }
}

/**
 * Periodic background sync guarantees 100% consistency across accounts/tabs
 * even when school network proxies buffer or block WebSockets.
 */
function startBackgroundSyncLoop() {
  setInterval(() => {
    if (document.visibilityState === 'visible') {
      fetchLatestFromServer(false);
    }
  }, 2500);

  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible') {
      fetchLatestFromServer(false);
    }
  });

  window.addEventListener('focus', () => {
    fetchLatestFromServer(false);
  });
}

/**
 * Check if the current user is actively typing inside a given team's card
 */
function isUserActivelyEditingTeam(teamId) {
  if (dirtyTeamIds.has(teamId)) return true;
  const active = document.activeElement;
  if (!active) return false;
  return active.id === `country-${teamId}` || active.id === `sentence-${teamId}`;
}

/**
 * Apply remote update for a single team without disturbing other cards being typed in
 */
function applyRemoteSingleTeam(remoteTeam) {
  const id = Number(remoteTeam.id);
  const localTeam = teamsData.find(t => t.id === id);
  if (!localTeam) return;

  if (isUserActivelyEditingTeam(id)) {
    return;
  }

  const changed =
    localTeam.countryName !== (remoteTeam.countryName || '') ||
    localTeam.sentence !== (remoteTeam.sentence || '');

  localTeam.countryName = remoteTeam.countryName || '';
  localTeam.sentence = remoteTeam.sentence || '';
  localTeam.updatedAt = remoteTeam.updatedAt || null;

  if (changed) {
    patchSingleInputCardDom(localTeam);
    renderTeacherBoardGrid();
    saveToLocalStorage();

    if (activeZoomTeamId === id) {
      updateZoomModalContent(id);
    }
  }
}

/**
 * Apply full 14-team list from server while preserving any card currently being typed by this user
 */
function applyRemoteTeamsList(remoteTeams) {
  let anyBoardChange = false;

  for (let i = 0; i < DEFAULT_TEAMS_COUNT; i++) {
    const id = i + 1;
    const remote = remoteTeams.find(r => Number(r.id) === id) || remoteTeams[i];
    if (!remote) continue;

    if (isUserActivelyEditingTeam(id)) {
      continue;
    }

    const local = teamsData.find(t => t.id === id);
    if (!local) continue;

    const newCountry = remote.countryName || '';
    const newSentence = remote.sentence || '';

    if (local.countryName !== newCountry || local.sentence !== newSentence) {
      local.countryName = newCountry;
      local.sentence = newSentence;
      local.updatedAt = remote.updatedAt || null;
      patchSingleInputCardDom(local);
      anyBoardChange = true;

      if (activeZoomTeamId === id) {
        updateZoomModalContent(id);
      }
    }
  }

  if (anyBoardChange) {
    renderTeacherBoardGrid();
    saveToLocalStorage();
  } else {
    updateStats();
  }
}

/**
 * Push a single team's changes to the shared server (and Google Apps Script if configured)
 */
async function pushSingleTeamToServer(teamId, showSuccessToast) {
  const team = teamsData.find(t => t.id === teamId);
  if (!team) return;

  if (autoSaveTimers.has(teamId)) {
    clearTimeout(autoSaveTimers.get(teamId));
    autoSaveTimers.delete(teamId);
  }

  setCardStatusUI(teamId, 'saving');

  let syncedRemotely = false;

  try {
    const res = await fetch(`/api/teams/${teamId}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        countryName: team.countryName,
        sentence: team.sentence
      })
    });

    if (res.ok) {
      const data = await res.json();
      if (data.team && data.team.updatedAt) {
        team.updatedAt = data.team.updatedAt;
      }
      serverAvailable = true;
      syncedRemotely = true;
    }
  } catch {
    // Ignore if running on pure static host
  }

  // Also sync to Google Apps Script if teacher configured a URL
  const scriptUrl = localStorage.getItem(STORAGE_KEY_SCRIPT_URL);
  if (scriptUrl) {
    try {
      await fetch(scriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          project: '지구마불세계여행',
          savedAt: new Date().toISOString(),
          updatedTeam: team,
          teams: teamsData
        })
      });
      syncedRemotely = true;
    } catch (err) {
      console.warn('Apps Script auto-sync warning:', err);
    }
  }

  dirtyTeamIds.delete(teamId);
  saveToLocalStorage();
  renderTeacherBoardGrid();
  setCardStatusUI(teamId, 'saved');

  if (showSuccessToast) {
    if (syncedRemotely) {
      showToast(`✈️ ${team.teamNumber} 내용이 교사 종합 전광판에 실시간 공유되었습니다!`);
    } else {
      showToast(`💾 ${team.teamNumber} 내용이 저장되었습니다!`);
    }
  }
}

function scheduleAutoSave(teamId) {
  dirtyTeamIds.add(teamId);
  setCardStatusUI(teamId, 'unsaved');

  if (autoSaveTimers.has(teamId)) {
    clearTimeout(autoSaveTimers.get(teamId));
  }

  const timer = setTimeout(() => {
    autoSaveTimers.delete(teamId);
    pushSingleTeamToServer(teamId, false);
  }, 650);

  autoSaveTimers.set(teamId, timer);
}

/**
 * =========================================================================
 * Rendering & Targeted DOM Updates
 * =========================================================================
 */

function renderApp() {
  renderInputGrid();
  renderTeacherBoardGrid();
}

function renderInputGrid() {
  const container = document.getElementById('inputGrid');
  if (!container) return;

  container.innerHTML = teamsData.map((team, idx) => {
    const colorIndex = (idx % 6) + 1;
    const hasContent = team.sentence.trim().length > 0 || team.countryName.trim().length > 0;
    const isCompleted = team.sentence.trim().length > 0;
    const statusText = isCompleted ? '공유 완료 ✈️' : (hasContent ? '나라 입력됨 🌍' : '작성 대기 ✏️');
    const statusClass = isCompleted ? 'status-completed' : 'status-empty';
    const displayName = team.countryName ? team.countryName : '나라명 미입력';

    return `
      <div class="team-card" data-color="${colorIndex}" id="input-card-${team.id}">
        <div class="card-header">
          <div class="team-title-wrap">
            <span class="team-number-badge">${team.teamNumber}</span>
            <span class="team-country-badge" id="badge-country-${team.id}">${escapeHtml(displayName)}</span>
          </div>
          <span class="team-status-tag ${statusClass}" id="status-tag-${team.id}">${statusText}</span>
        </div>

        <div class="input-group">
          <label class="input-label" for="country-${team.id}">🌍 탐구 나라 이름</label>
          <input 
            type="text" 
            id="country-${team.id}" 
            class="input-country" 
            placeholder="예: 🇰🇷 대한민국, 🇪🇬 이집트 등" 
            value="${escapeHtml(team.countryName)}"
            oninput="handleCountryInput(${team.id}, this.value)"
            onblur="handleInputBlur(${team.id})"
          />
        </div>

        <div class="input-group input-group-grow">
          <label class="input-label" for="sentence-${team.id}">💡 일반화된 문장</label>
          <textarea 
            id="sentence-${team.id}" 
            class="input-sentence" 
            placeholder="다른 모둠의 기후, 지형, 생활모습을 관찰하고 탐구하여 알게 된 일반화된 문장을 적어보세요!"
            oninput="handleSentenceInput(${team.id}, this.value)"
            onblur="handleInputBlur(${team.id})"
            onkeydown="if((event.ctrlKey || event.metaKey) && event.key === 'Enter') manualSaveTeam(${team.id})"
          >${escapeHtml(team.sentence)}</textarea>
        </div>

        <div class="card-actions">
          <button type="button" class="btn-card-clear" onclick="confirmClearTeam(${team.id})" title="입력 내용 지우기">
            🗑️ 지우기
          </button>
          <button type="button" class="btn-card-save" id="btn-save-${team.id}" onclick="manualSaveTeam(${team.id})">
            💾 전광판에 저장
          </button>
        </div>
      </div>
    `;
  }).join('');
}

function patchSingleInputCardDom(team) {
  const countryInput = document.getElementById(`country-${team.id}`);
  const sentenceInput = document.getElementById(`sentence-${team.id}`);
  const badge = document.getElementById(`badge-country-${team.id}`);

  if (countryInput && document.activeElement !== countryInput) {
    countryInput.value = team.countryName;
  }
  if (sentenceInput && document.activeElement !== sentenceInput) {
    sentenceInput.value = team.sentence;
  }
  if (badge) {
    badge.textContent = team.countryName ? team.countryName : '나라명 미입력';
  }
  setCardStatusUI(team.id, 'saved');
}

function setCardStatusUI(teamId, state) {
  const team = teamsData.find(t => t.id === teamId);
  if (!team) return;

  const statusTag = document.getElementById(`status-tag-${teamId}`);
  const saveBtn = document.getElementById(`btn-save-${teamId}`);

  if (state === 'saving') {
    if (statusTag) {
      statusTag.textContent = '공유 중... ⏳';
      statusTag.className = 'team-status-tag status-saving';
    }
    if (saveBtn) {
      saveBtn.textContent = '⏳ 저장 중...';
      saveBtn.classList.remove('unsaved');
    }
    return;
  }

  if (state === 'unsaved') {
    if (statusTag) {
      statusTag.textContent = '작성 중... ✏️';
      statusTag.className = 'team-status-tag status-saving';
    }
    if (saveBtn) {
      saveBtn.textContent = '💾 전광판에 저장';
      saveBtn.classList.add('unsaved');
    }
    return;
  }

  // 'saved'
  const isCompleted = team.sentence.trim().length > 0;
  const hasCountry = team.countryName.trim().length > 0;
  if (statusTag) {
    if (isCompleted) {
      statusTag.textContent = '공유 완료 ✈️';
      statusTag.className = 'team-status-tag status-completed';
    } else if (hasCountry) {
      statusTag.textContent = '나라 입력됨 🌍';
      statusTag.className = 'team-status-tag status-empty';
    } else {
      statusTag.textContent = '작성 대기 ✏️';
      statusTag.className = 'team-status-tag status-empty';
    }
  }
  if (saveBtn) {
    saveBtn.textContent = '💾 전광판에 저장';
    saveBtn.classList.remove('unsaved');
  }
}

/**
 * Render Teacher Billboard Mode Cards
 */
function renderTeacherBoardGrid() {
  const container = document.getElementById('boardGrid');
  if (!container) return;

  container.innerHTML = teamsData.map((team, idx) => {
    const colorIndex = (idx % 6) + 1;
    const isCompleted = team.sentence.trim().length > 0;
    const countryLabel = team.countryName.trim() ? team.countryName : '(나라 미정)';
    const statusText = isCompleted ? '공유 완료' : '작성 대기';
    const statusClass = isCompleted ? 'status-completed' : 'status-empty';

    return `
      <div 
        class="team-card board-card" 
        data-color="${colorIndex}" 
        onclick="openZoomModal(${team.id})"
        tabindex="0"
        role="button"
        aria-label="${team.teamNumber} ${escapeHtml(countryLabel)} 내용 크게 보기"
        onkeydown="if(event.key === 'Enter') openZoomModal(${team.id})"
      >
        <div class="card-header">
          <div class="team-title-wrap">
            <span class="team-number-badge">${team.teamNumber}</span>
            <span class="team-country-badge">${escapeHtml(countryLabel)}</span>
          </div>
          <span class="team-status-tag ${statusClass}">${statusText}</span>
        </div>

        <div class="board-sentence-area">
          ${isCompleted 
            ? `<p class="board-sentence-text">“${escapeHtml(team.sentence)}”</p>`
            : `<p class="board-sentence-empty">아직 일반화된 문장이 작성되지 않았습니다.</p>`
          }
        </div>

        <div class="board-card-footer">
          <span class="click-hint">🔍 클릭하여 전체화면 크게 보기</span>
          <button 
            type="button" 
            class="board-delete-btn" 
            onclick="event.stopPropagation(); confirmClearTeam(${team.id})"
            title="문장 삭제"
          >
            지우기
          </button>
        </div>
      </div>
    `;
  }).join('');
}

/**
 * Input Event Handlers
 */
function handleCountryInput(id, val) {
  const team = teamsData.find(t => t.id === id);
  if (!team) return;

  team.countryName = val.trim();
  team.updatedAt = new Date().toISOString();

  const badge = document.getElementById(`badge-country-${id}`);
  if (badge) {
    badge.textContent = team.countryName ? team.countryName : '나라명 미입력';
  }

  saveToLocalStorage();
  renderTeacherBoardGrid();
  scheduleAutoSave(id);
}

function handleSentenceInput(id, val) {
  const team = teamsData.find(t => t.id === id);
  if (!team) return;

  team.sentence = val.trim();
  team.updatedAt = new Date().toISOString();

  saveToLocalStorage();
  renderTeacherBoardGrid();
  scheduleAutoSave(id);
}

function handleInputBlur(id) {
  if (dirtyTeamIds.has(id)) {
    pushSingleTeamToServer(id, false);
  }
}

async function manualSaveTeam(id) {
  const countryEl = document.getElementById(`country-${id}`);
  const sentenceEl = document.getElementById(`sentence-${id}`);
  const team = teamsData.find(t => t.id === id);
  if (!team) return;

  if (countryEl) team.countryName = countryEl.value.trim();
  if (sentenceEl) team.sentence = sentenceEl.value.trim();
  team.updatedAt = new Date().toISOString();

  await pushSingleTeamToServer(id, true);
}

/**
 * =========================================================================
 * Delete / Reset with Custom Confirmation Modal (iframe-safe)
 * =========================================================================
 */

function openConfirmModal(title, message, confirmBtnText, onConfirm) {
  const modal = document.getElementById('confirmModal');
  const titleEl = document.getElementById('confirmModalTitle');
  const msgEl = document.getElementById('confirmModalMessage');
  const btnEl = document.getElementById('btnConfirmProceed');

  if (titleEl) titleEl.textContent = title;
  if (msgEl) msgEl.textContent = message;
  if (btnEl) {
    btnEl.textContent = confirmBtnText || '확인';
    btnEl.onclick = async () => {
      closeConfirmModal();
      if (typeof onConfirm === 'function') {
        await onConfirm();
      }
    };
  }

  pendingConfirmCallback = onConfirm;
  if (modal) modal.classList.add('open');
}

function closeConfirmModal() {
  const modal = document.getElementById('confirmModal');
  if (modal) modal.classList.remove('open');
  pendingConfirmCallback = null;
}

function confirmClearTeam(id) {
  const team = teamsData.find(t => t.id === id);
  if (!team) return;

  openConfirmModal(
    `🗑️ ${team.teamNumber} 내용 지우기`,
    `${team.teamNumber}의 나라 이름과 일반화된 문장을 전광판에서 모두 지우시겠습니까?`,
    '네, 지우겠습니다',
    async () => {
      if (autoSaveTimers.has(id)) {
        clearTimeout(autoSaveTimers.get(id));
        autoSaveTimers.delete(id);
      }
      dirtyTeamIds.delete(id);

      team.countryName = '';
      team.sentence = '';
      team.updatedAt = null;

      try {
        await fetch(`/api/teams/${id}`, { method: 'DELETE' });
      } catch {
        // Fallback local clear
      }

      saveToLocalStorage();
      renderApp();
      showToast(`${team.teamNumber} 내용이 삭제되었습니다.`);

      if (activeZoomTeamId === id) {
        updateZoomModalContent(id);
      }
    }
  );
}

function resetAllTeams() {
  openConfirmModal(
    '⚠️ 전체 14개 모둠 초기화',
    '전체 14개 모둠의 모든 입력 내용을 깨끗하게 비우시겠습니까? (모든 사용자의 전광판에서 초기화됩니다)',
    '전체 초기화 실행',
    async () => {
      autoSaveTimers.forEach(timer => clearTimeout(timer));
      autoSaveTimers.clear();
      dirtyTeamIds.clear();

      teamsData = generateDefaultTeams();

      try {
        await fetch('/api/teams', {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ mode: 'reset-all' })
        });
      } catch {
        // Fallback local reset
      }

      saveToLocalStorage();
      renderApp();
      showToast('전체 14개 모둠이 깨끗하게 초기화되었습니다.');
    }
  );
}

/**
 * Update Top Stats Counter
 */
function updateStats() {
  const completedCount = teamsData.filter(t => t.sentence.trim().length > 0).length;
  const countEl = document.getElementById('completedCount');
  if (countEl) {
    countEl.textContent = `${completedCount} / ${DEFAULT_TEAMS_COUNT}`;
  }
}

/**
 * Switch between "Student Input Mode" and "Teacher Billboard Mode"
 */
function setViewMode(mode, silent) {
  // Flush any pending unsaved edits before switching to Teacher Billboard!
  if (mode === 'teacher' && dirtyTeamIds.size > 0) {
    Array.from(dirtyTeamIds).forEach(teamId => {
      pushSingleTeamToServer(teamId, false);
    });
  }

  currentViewMode = mode;
  const btnInput = document.getElementById('btnModeInput');
  const btnTeacher = document.getElementById('btnModeTeacher');
  const noticeText = document.getElementById('viewNoticeText');
  const noticeHint = document.getElementById('viewNoticeHint');

  if (mode === 'teacher') {
    document.body.classList.add('mode-teacher');
    if (btnInput) btnInput.classList.remove('active');
    if (btnTeacher) btnTeacher.classList.add('active');
    if (noticeText) {
      noticeText.innerHTML = '<span>📺 <strong>교사 종합 전광판</strong> — 각 모둠이 저장한 일반화 문장이 실시간으로 한눈에 모아집니다!</span>';
    }
    if (noticeHint) {
      noticeHint.textContent = '모둠 카드를 클릭하면 전체화면 발표 창으로 크게 확대됩니다.';
    }
    fetchLatestFromServer(false);
    renderTeacherBoardGrid();
    if (!silent) {
      showToast('📺 교사 종합 전광판 모드로 전환되었습니다!');
    }
  } else {
    document.body.classList.remove('mode-teacher');
    if (btnInput) btnInput.classList.add('active');
    if (btnTeacher) btnTeacher.classList.remove('active');
    if (noticeText) {
      noticeText.innerHTML = '<span>✨ <strong>14개 모둠 탐구 카드</strong> 중 우리 모둠 칸에 나라 이름과 일반화된 문장을 작성한 뒤 <strong>[💾 전광판에 저장]</strong>을 눌러 주세요!</span>';
    }
    if (noticeHint) {
      noticeHint.textContent = '저장 즉시 교사 종합 전광판에 실시간으로 반영됩니다.';
    }
    renderInputGrid();
  }
}

/**
 * Copy Shareable Classroom Link (converts ais-dev preview hostname to ais-pre shared URL automatically)
 */
async function copyShareLink() {
  let shareOrigin = window.location.origin;
  if (window.location.hostname.startsWith('ais-dev-')) {
    shareOrigin = `${window.location.protocol}//${window.location.hostname.replace(/^ais-dev-/, 'ais-pre-')}`;
  }

  let shareUrl = `${shareOrigin}${window.location.pathname}`;
  const scriptUrl = localStorage.getItem(STORAGE_KEY_SCRIPT_URL);
  if (!serverAvailable && scriptUrl) {
    shareUrl += `?sheet=${encodeURIComponent(scriptUrl)}`;
  }

  try {
    await navigator.clipboard.writeText(shareUrl);
    showToast('🔗 학생 공유용 링크가 클립보드에 복사되었습니다! (타 계정에서도 바로 접속 가능)');
  } catch {
    // Fallback copy method for iframes
    const tempInput = document.createElement('input');
    tempInput.value = shareUrl;
    document.body.appendChild(tempInput);
    tempInput.select();
    document.execCommand('copy');
    tempInput.remove();
    showToast('🔗 학생 공유용 링크가 복사되었습니다!');
  }
}

/**
 * =========================================================================
 * Fullscreen Zoom Presentation Modal
 * =========================================================================
 */

function openZoomModal(teamId) {
  activeZoomTeamId = teamId;
  updateZoomModalContent(teamId);
  const modal = document.getElementById('zoomModal');
  if (modal) {
    modal.classList.add('open');
    modal.setAttribute('aria-hidden', 'false');
  }
}

function updateZoomModalContent(teamId) {
  const team = teamsData.find(t => t.id === teamId);
  if (!team) return;

  const badgeEl = document.getElementById('zoomTeamBadge');
  const titleEl = document.getElementById('zoomCountryTitle');
  const bodyEl = document.getElementById('zoomBodyContainer');
  const btnPrev = document.getElementById('btnZoomPrev');
  const btnNext = document.getElementById('btnZoomNext');

  if (badgeEl) badgeEl.textContent = team.teamNumber;
  if (titleEl) {
    titleEl.textContent = team.countryName ? `🌍 ${team.countryName}` : '🌍 (탐구 국가 미입력)';
  }

  if (bodyEl) {
    if (team.sentence.trim().length > 0) {
      bodyEl.innerHTML = `
        <span class="quote-mark">“</span>
        <div class="zoom-sentence">${escapeHtml(team.sentence)}</div>
        <span class="quote-mark">”</span>
      `;
    } else {
      bodyEl.innerHTML = `
        <div class="zoom-empty-notice">
          아직 일반화된 문장이 작성되지 않았습니다.<br>
          <span class="zoom-empty-sub">[모둠별 입력 모드]에서 문장을 입력하고 저장하면 이곳에 실시간으로 표시됩니다!</span>
        </div>
      `;
    }
  }

  if (btnPrev) btnPrev.disabled = teamId <= 1;
  if (btnNext) btnNext.disabled = teamId >= DEFAULT_TEAMS_COUNT;
}

function closeZoomModal() {
  const modal = document.getElementById('zoomModal');
  if (modal) {
    modal.classList.remove('open');
    modal.setAttribute('aria-hidden', 'true');
  }
  activeZoomTeamId = null;
}

function navigateZoom(direction) {
  if (!activeZoomTeamId) return;
  const targetId = activeZoomTeamId + direction;
  if (targetId >= 1 && targetId <= DEFAULT_TEAMS_COUNT) {
    activeZoomTeamId = targetId;
    updateZoomModalContent(targetId);
  }
}

function deleteCurrentZoomTeam() {
  if (!activeZoomTeamId) return;
  confirmClearTeam(activeZoomTeamId);
}

/**
 * =========================================================================
 * Batch Country Names Modal
 * =========================================================================
 */

function openBatchNamesModal() {
  const container = document.getElementById('batchInputsContainer');
  if (!container) return;

  container.innerHTML = teamsData.map(team => `
    <div class="batch-row">
      <span class="batch-label">${team.teamNumber}</span>
      <input 
        type="text" 
        id="batch-country-${team.id}" 
        class="batch-input" 
        value="${escapeHtml(team.countryName)}" 
        placeholder="나라 이름 입력"
      />
    </div>
  `).join('');

  document.getElementById('batchModal').classList.add('open');
}

function closeBatchNamesModal() {
  document.getElementById('batchModal').classList.remove('open');
}

async function saveBatchNames() {
  const countriesMap = {};
  teamsData.forEach(team => {
    const input = document.getElementById(`batch-country-${team.id}`);
    if (input) {
      const val = input.value.trim();
      team.countryName = val;
      countriesMap[String(team.id)] = val;
    }
  });

  try {
    const res = await fetch('/api/teams', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        mode: 'batch-countries',
        countries: countriesMap
      })
    });
    if (res.ok) {
      const data = await res.json();
      if (Array.isArray(data.teams)) {
        teamsData = data.teams;
      }
    }
  } catch {
    // Fallback local
  }

  saveToLocalStorage();
  renderApp();
  closeBatchNamesModal();
  showToast('🌍 14개 모둠의 나라 이름이 모든 화면에 공유 적용되었습니다!');
}

/**
 * =========================================================================
 * Random Role Matching Modal
 * =========================================================================
 */

function openRoleModal() {
  document.getElementById('roleModal').classList.add('open');
}

function closeRoleModal() {
  document.getElementById('roleModal').classList.remove('open');
}

function runRandomRoleMatching() {
  const rawInput = document.getElementById('roleMembersInput').value;
  const members = rawInput
    .split(/[\n, ]+/)
    .map(name => name.trim())
    .filter(name => name.length > 0);

  if (members.length < 2) {
    showToast('모둠원 이름을 최소 2명 이상 입력해 주세요!');
    return;
  }

  const shuffled = [...members].sort(() => Math.random() - 0.5);
  const container = document.getElementById('roleResultGrid');
  if (!container) return;

  container.innerHTML = DEFAULT_ROLES.map((role, idx) => {
    const assignedPerson = shuffled[idx % shuffled.length];
    return `
      <div class="role-card ${role.cardClass}">
        <div class="role-title">${role.title}</div>
        <div class="role-person">${escapeHtml(assignedPerson)}</div>
      </div>
    `;
  }).join('');

  showToast('🎉 역할이 공정하고 즐겁게 배정되었습니다!');
}

/**
 * =========================================================================
 * Google Sheets Sync & JSON Backup Handlers
 * =========================================================================
 */

function openSheetModal(actionType) {
  const modal = document.getElementById('sheetModal');
  const title = document.getElementById('sheetModalTitle');
  const urlInput = document.getElementById('sheetUrlInput');
  const actionBtn = document.getElementById('btnExecuteSheetAction');

  const savedUrl = localStorage.getItem(STORAGE_KEY_SCRIPT_URL) || '';
  if (urlInput) urlInput.value = savedUrl;

  if (actionType === 'load') {
    title.textContent = '🔄 구글 스프레드시트에서 불러오기';
    actionBtn.textContent = '시트에서 지금 불러오기';
    actionBtn.onclick = () => executeSheetLoad();
  } else {
    title.textContent = '💾 구글 스프레드시트에 결과 저장';
    actionBtn.textContent = '시트에 지금 저장하기';
    actionBtn.onclick = () => executeSheetSave();
  }

  modal.classList.add('open');
}

function closeSheetModal() {
  document.getElementById('sheetModal').classList.remove('open');
}

async function saveSheetUrlToSharedConfig() {
  const url = document.getElementById('sheetUrlInput').value.trim();
  localStorage.setItem(STORAGE_KEY_SCRIPT_URL, url);
  try {
    await fetch('/api/config', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ appsScriptUrl: url })
    });
  } catch {
    // Ignore on static host
  }
  return url;
}

async function silentFetchFromAppsScript(scriptUrl, showFeedback) {
  try {
    const res = await fetch(scriptUrl);
    if (!res.ok) return;
    const data = await res.json();
    const list = Array.isArray(data) ? data : (Array.isArray(data?.teams) ? data.teams : null);
    if (list) {
      applyRemoteTeamsList(list);
      if (showFeedback) {
        showToast('🔄 구글 시트에서 최신 내용을 불러왔습니다!');
      }
    }
  } catch {
    // Silent fallback
  }
}

async function executeSheetLoad() {
  const scriptUrl = await saveSheetUrlToSharedConfig();

  if (!scriptUrl) {
    await fetchLatestFromServer(true);
    closeSheetModal();
    return;
  }

  const actionBtn = document.getElementById('btnExecuteSheetAction');
  const origText = actionBtn.textContent;
  actionBtn.disabled = true;
  actionBtn.textContent = '불러오는 중... ⏳';

  try {
    const res = await fetch(scriptUrl);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const raw = await res.json();
    const list = Array.isArray(raw) ? raw : (Array.isArray(raw?.teams) ? raw.teams : null);

    if (list && list.length > 0) {
      const normalized = teamsData.map((def, idx) => {
        const item = list.find(r => Number(r.id) === def.id) || list[idx] || {};
        return {
          id: def.id,
          teamNumber: def.teamNumber,
          countryName: item.countryName || '',
          sentence: item.sentence || '',
          updatedAt: item.updatedAt || new Date().toISOString()
        };
      });

      teamsData = normalized;
      try {
        await fetch('/api/teams', {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ mode: 'replace-all', teams: teamsData })
        });
      } catch {
        // Static fallback
      }

      saveToLocalStorage();
      renderApp();
      closeSheetModal();
      showToast('🎉 구글 스프레드시트 데이터를 불러와 전체 전광판에 반영했습니다!');
    } else {
      showToast('시트에 불러올 모둠 데이터가 없습니다.');
    }
  } catch (err) {
    console.error(err);
    showToast('시트 불러오기 실패: Apps Script 배포 URL과 권한(모든 사용자)을 확인해 주세요.');
  } finally {
    actionBtn.disabled = false;
    actionBtn.textContent = origText;
  }
}

async function executeSheetSave() {
  const scriptUrl = await saveSheetUrlToSharedConfig();

  if (!scriptUrl) {
    showToast('Apps Script URL을 입력하시거나 [결과 파일 다운로드]를 이용해 주세요.');
    return;
  }

  const actionBtn = document.getElementById('btnExecuteSheetAction');
  const origText = actionBtn.textContent;
  actionBtn.disabled = true;
  actionBtn.textContent = '저장하는 중... ⏳';

  try {
    const res = await fetch(scriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({
        project: '지구마불세계여행',
        savedAt: new Date().toISOString(),
        teams: teamsData
      })
    });

    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    closeSheetModal();
    showToast('🎉 구글 스프레드시트에 전체 모둠 결과가 저장되었습니다!');
  } catch (err) {
    console.error(err);
    showToast('시트 저장 실패: Apps Script 배포 설정을 확인해 주세요.');
  } finally {
    actionBtn.disabled = false;
    actionBtn.textContent = origText;
  }
}

function exportDataAsJson() {
  const dataStr = 'data:text/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(teamsData, null, 2));
  const downloadAnchor = document.createElement('a');
  downloadAnchor.setAttribute('href', dataStr);
  const now = new Date().toISOString().slice(0, 10);
  downloadAnchor.setAttribute('download', `지구마불세계여행_모둠탐구_${now}.json`);
  document.body.appendChild(downloadAnchor);
  downloadAnchor.click();
  downloadAnchor.remove();
  showToast('📥 결과 파일(JSON)이 다운로드되었습니다.');
}

function importDataFromJson() {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.json';
  input.onchange = e => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = async event => {
      try {
        const parsed = JSON.parse(event.target.result);
        if (Array.isArray(parsed)) {
          teamsData = teamsData.map((def, idx) => {
            const item = parsed.find(r => Number(r.id) === def.id) || parsed[idx] || {};
            return {
              id: def.id,
              teamNumber: def.teamNumber,
              countryName: item.countryName || '',
              sentence: item.sentence || '',
              updatedAt: item.updatedAt || new Date().toISOString()
            };
          });

          try {
            await fetch('/api/teams', {
              method: 'PUT',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ mode: 'replace-all', teams: teamsData })
            });
          } catch {
            // Static fallback
          }

          saveToLocalStorage();
          renderApp();
          closeSheetModal();
          showToast('📂 파일을 불러와 전체 전광판에 공유 반영했습니다!');
        } else {
          showToast('올바른 모둠 데이터 파일 형식이 아닙니다.');
        }
      } catch {
        showToast('파일을 읽는 도중 오류가 발생했습니다.');
      }
    };
    reader.readAsText(file);
  };
  input.click();
}

/**
 * Fill Demo / Sample Data and broadcast to everyone
 */
async function fillSampleData() {
  const sampleCountries = [
    { name: '🇰🇷 대한민국', sentence: '사계절이 뚜렷한 온대 기후 덕분에 계절별로 다양한 농작물과 의생활 문화가 발달하였다.' },
    { name: '🇪🇬 이집트', sentence: '건조한 사막 기후 지역에서는 나일강과 오아시스 주변에 인구가 밀집하여 독특한 문명을 형성하였다.' },
    { name: '🇧🇷 브라질', sentence: '풍부한 아마존 열대 우림과 따뜻한 기후 속에서 다양한 생태계와 활기찬 축제 문화가 공존한다.' },
    { name: '🇫🇷 프랑스', sentence: '온화한 서안 해양성 및 지중해성 기후로 포도 농사와 미식, 문화 예술이 크게 발달하였다.' },
    { name: '🇨🇦 캐나다', sentence: '넓은 영토에 걸친 냉대 기후와 침엽수림(타이가) 지형으로 목재 산업과 겨울 스포츠가 발달했다.' },
    { name: '🇦🇺 호주', sentence: '남반구에 위치하여 계절이 반대이며, 광활한 건조 지대와 독특한 고유 생물들이 살아간다.' },
    { name: '🇮🇸 아이슬란드', sentence: '빙하와 화산이 공존하는 독특한 지형으로 풍부한 지열 에너지를 일상생활에 적극 활용한다.' },
    { name: '🇲🇳 몽골', sentence: '건조한 스텝 초원 지대에서 계절에 따라 가축을 이동시키는 전통적인 유목 생활을 이어왔다.' }
  ];

  teamsData.forEach((team, idx) => {
    if (idx < sampleCountries.length) {
      team.countryName = sampleCountries[idx].name;
      team.sentence = sampleCountries[idx].sentence;
      team.updatedAt = new Date().toISOString();
    }
  });

  try {
    await fetch('/api/teams', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ mode: 'replace-all', teams: teamsData })
    });
  } catch {
    // Fallback local
  }

  saveToLocalStorage();
  renderApp();
  showToast('✨ 시연용 샘플 데이터가 전체 전광판에 반영되었습니다!');
}

/**
 * Keyboard shortcuts
 */
function setupEventListeners() {
  document.addEventListener('keydown', (e) => {
    const zoomModal = document.getElementById('zoomModal');
    const isZoomOpen = zoomModal && zoomModal.classList.contains('open');

    if (e.key === 'Escape') {
      if (isZoomOpen) {
        closeZoomModal();
      } else {
        closeConfirmModal();
        closeBatchNamesModal();
        closeRoleModal();
        closeSheetModal();
      }
    }

    if (isZoomOpen) {
      if (e.key === 'ArrowLeft') {
        navigateZoom(-1);
      } else if (e.key === 'ArrowRight') {
        navigateZoom(1);
      }
    }
  });
}

/**
 * Toast Notification Helper
 */
let toastTimeout = null;
function showToast(message) {
  const toast = document.getElementById('toastMsg');
  const toastText = document.getElementById('toastText');
  if (!toast || !toastText) return;

  toastText.textContent = message;
  toast.classList.add('show');

  if (toastTimeout) clearTimeout(toastTimeout);
  toastTimeout = setTimeout(() => {
    toast.classList.remove('show');
  }, 2800);
}

/**
 * Safe HTML string escape helper
 */
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
