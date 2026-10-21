/**
 * 지구마불세계여행 프로젝트 - 최종 차시 탐구 종합 전광판
 * 순수 Vanilla JavaScript 동작 로직
 */

// LocalStorage Keys
const STORAGE_KEY_TEAMS = 'earth_marble_teams_v1';
const STORAGE_KEY_SCRIPT_URL = 'earth_marble_apps_script_url';

// Default 14 Teams Setup
const DEFAULT_TEAMS_COUNT = 14;

// Global State
let teamsData = [];
let currentViewMode = 'input'; // 'input' | 'teacher'
let activeZoomTeamId = null;

// Initialize on DOM Ready
document.addEventListener('DOMContentLoaded', () => {
  initData();
  renderApp();
  setupEventListeners();
  updateStats();
});

/**
 * Initialize 14 Teams from LocalStorage or Defaults
 */
function initData() {
  const saved = localStorage.getItem(STORAGE_KEY_TEAMS);
  if (saved) {
    try {
      teamsData = JSON.parse(saved);
      // Ensure always exactly 14 teams
      if (!Array.isArray(teamsData) || teamsData.length !== DEFAULT_TEAMS_COUNT) {
        teamsData = generateDefaultTeams();
        saveToLocalStorage();
      }
    } catch (e) {
      console.error('Failed to parse saved data, resetting', e);
      teamsData = generateDefaultTeams();
      saveToLocalStorage();
    }
  } else {
    teamsData = generateDefaultTeams();
    saveToLocalStorage();
  }
}

/**
 * Generate 14 Initial Default Teams
 */
function generateDefaultTeams() {
  const list = [];
  for (let i = 1; i <= DEFAULT_TEAMS_COUNT; i++) {
    list.push({
      id: i,
      teamNumber: `${i}모둠`,
      countryName: '', // Children will enter directly
      sentence: '',    // Generalization sentence
      updatedAt: null
    });
  }
  return list;
}

/**
 * Save current state to browser LocalStorage
 */
function saveToLocalStorage() {
  localStorage.setItem(STORAGE_KEY_TEAMS, JSON.stringify(teamsData));
  updateStats();
}

/**
 * Render both Input Cards and Teacher Billboard Cards
 */
function renderApp() {
  renderInputGrid();
  renderTeacherBoardGrid();
}

/**
 * Render Input Mode Cards
 */
function renderInputGrid() {
  const container = document.getElementById('inputGrid');
  if (!container) return;

  container.innerHTML = teamsData.map((team, idx) => {
    const colorIndex = (idx % 6) + 1;
    const isCompleted = team.sentence.trim().length > 0;
    const statusText = isCompleted ? '작성 완료 ✈️' : '작성 중 ✏️';
    const statusClass = isCompleted ? 'status-completed' : 'status-empty';
    const displayName = team.countryName ? `${team.countryName}` : '나라명 미입력';

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
          />
        </div>

        <div class="input-group" style="flex: 1; display: flex; flex-direction: column;">
          <label class="input-label" for="sentence-${team.id}">💡 일반화된 문장</label>
          <textarea 
            id="sentence-${team.id}" 
            class="input-sentence" 
            placeholder="다른 모둠의 기후, 지형, 생활모습을 관찰하고 탐구하여 알게 된 일반화된 문장을 적어보세요!"
            oninput="handleSentenceInput(${team.id}, this.value)"
          >${escapeHtml(team.sentence)}</textarea>
        </div>

        <div class="card-actions">
          <button type="button" class="btn-card-clear" onclick="confirmClearTeam(${team.id})" title="입력 내용 지우기">
            🗑️ 지우기
          </button>
          <button type="button" class="btn-card-save" onclick="manualSaveTeam(${team.id})">
            💾 저장하기
          </button>
        </div>
      </div>
    `;
  }).join('');
}

/**
 * Render Teacher Billboard Mode Cards (Large readable cards for projector)
 */
function renderTeacherBoardGrid() {
  const container = document.getElementById('boardGrid');
  if (!container) return;

  container.innerHTML = teamsData.map((team, idx) => {
    const colorIndex = (idx % 6) + 1;
    const isCompleted = team.sentence.trim().length > 0;
    const countryLabel = team.countryName.trim() ? team.countryName : '(나라 미정)';
    const statusText = isCompleted ? '제출 완료' : '작성 대기';
    const statusClass = isCompleted ? 'status-completed' : 'status-empty';

    return `
      <div 
        class="team-card board-card" 
        data-color="${colorIndex}" 
        onclick="openZoomModal(${team.id})"
        tabindex="0"
        role="button"
        aria-label="${team.teamNumber} ${countryLabel} 내용 크게 보기"
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
          <span class="click-hint">🔍 클릭하여 화면 전체 확대</span>
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
 * Real-time Country Input Handler
 */
function handleCountryInput(id, val) {
  const team = teamsData.find(t => t.id === id);
  if (team) {
    team.countryName = val.trim();
    team.updatedAt = new Date().toISOString();
    
    // Update badge preview in input mode
    const badge = document.getElementById(`badge-country-${id}`);
    if (badge) {
      badge.textContent = team.countryName ? team.countryName : '나라명 미입력';
    }
    
    saveToLocalStorage();
    // Also re-render teacher board silently
    renderTeacherBoardGrid();
  }
}

/**
 * Real-time Sentence Input Handler
 */
function handleSentenceInput(id, val) {
  const team = teamsData.find(t => t.id === id);
  if (team) {
    team.sentence = val.trim();
    team.updatedAt = new Date().toISOString();
    
    // Update status tag in input card
    const statusTag = document.getElementById(`status-tag-${id}`);
    if (statusTag) {
      if (team.sentence.length > 0) {
        statusTag.textContent = '작성 완료 ✈️';
        statusTag.className = 'team-status-tag status-completed';
      } else {
        statusTag.textContent = '작성 중 ✏️';
        statusTag.className = 'team-status-tag status-empty';
      }
    }
    
    saveToLocalStorage();
    renderTeacherBoardGrid();
  }
}

/**
 * Manual Save Button click on single card
 */
function manualSaveTeam(id) {
  const team = teamsData.find(t => t.id === id);
  if (team) {
    saveToLocalStorage();
    showToast(`${team.teamNumber} 내용이 안전하게 저장되었습니다! ✨`);
  }
}

/**
 * Delete / Clear Sentence with Confirmation
 */
function confirmClearTeam(id) {
  const team = teamsData.find(t => t.id === id);
  if (!team) return;

  const confirmed = confirm(`${team.teamNumber}의 입력 내용(나라 이름 및 일반화된 문장)을 정말 지우시겠습니까?`);
  if (confirmed) {
    team.countryName = '';
    team.sentence = '';
    team.updatedAt = null;
    saveToLocalStorage();
    renderApp();
    showToast(`${team.teamNumber} 내용이 삭제되었습니다.`);
    
    // If zoom modal is currently open with this team, refresh zoom view
    if (activeZoomTeamId === id) {
      updateZoomModalContent(id);
    }
  }
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
function setViewMode(mode) {
  currentViewMode = mode;
  const btnInput = document.getElementById('btnModeInput');
  const btnTeacher = document.getElementById('btnModeTeacher');

  if (mode === 'teacher') {
    document.body.classList.add('mode-teacher');
    btnInput.classList.remove('active');
    btnTeacher.classList.add('active');
    renderTeacherBoardGrid();
    showToast('📺 교사 종합 전광판 모드로 전환되었습니다. (카드를 누르면 확대됩니다)');
  } else {
    document.body.classList.remove('mode-teacher');
    btnInput.classList.add('active');
    btnTeacher.classList.remove('active');
    renderInputGrid();
  }
}

/**
 * Fullscreen Zoom Presentation Modal Handlers
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
          <span style="font-size: 1.1rem; color: #7f8c8d; font-family: var(--font-body); margin-top: 8px; display: inline-block;">
            [모둠 입력 모드]에서 문장을 입력해보세요!
          </span>
        </div>
      `;
    }
  }

  // Prev / Next button states
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
 * Batch Country Names Modal
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

function saveBatchNames() {
  teamsData.forEach(team => {
    const input = document.getElementById(`batch-country-${team.id}`);
    if (input) {
      team.countryName = input.value.trim();
    }
  });

  saveToLocalStorage();
  renderApp();
  closeBatchNamesModal();
  showToast('모든 모둠의 나라 이름이 적용되었습니다! 🌍');
}

/**
 * Random Role Matching Modal (역할 랜덤 매칭)
 */
const DEFAULT_ROLES = [
  { id: 'speaker', title: '📢 발표자 (탐구 결과 발표)', color: '#ff7675' },
  { id: 'writer', title: '✍️ 기록자 (문장 정리 및 작성)', color: '#4fa8f7' },
  { id: 'timer', title: '⏱️ 시간지기 (활동 시간 관리)', color: '#feca57' },
  { id: 'organizer', title: '🌟 정리대장 (자료 및 모둠 정리)', color: '#48c9b0' }
];

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
    alert('모둠원 이름을 최소 2명 이상 입력해 주세요!');
    return;
  }

  // Shuffle members
  const shuffled = [...members].sort(() => Math.random() - 0.5);

  const container = document.getElementById('roleResultGrid');
  if (!container) return;

  container.innerHTML = DEFAULT_ROLES.map((role, idx) => {
    // Round-robin assignment if fewer or more than 4 members
    const assignedPerson = shuffled[idx % shuffled.length];
    return `
      <div class="role-card" style="border-color: ${role.color}">
        <div class="role-title" style="color: ${role.color}">${role.title}</div>
        <div class="role-person">${escapeHtml(assignedPerson)}</div>
      </div>
    `;
  }).join('');

  showToast('🎉 역할이 공정하고 즐겁게 배정되었습니다!');
}

/**
 * Google Apps Script / Sheet Sync Modal Handlers
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

function saveSheetUrl() {
  const url = document.getElementById('sheetUrlInput').value.trim();
  localStorage.setItem(STORAGE_KEY_SCRIPT_URL, url);
  return url;
}

/**
 * Execute Sheet Load via Apps Script Web App
 */
async function executeSheetLoad() {
  const scriptUrl = saveSheetUrl();
  
  if (!scriptUrl) {
    alert('아직 구글 Apps Script Web App URL이 설정되지 않았습니다.\n아래 [JSON 데이터로 즉시 내보내기/가져오기]를 사용하시거나 스크립트 URL을 입력해 주세요.');
    return;
  }

  const actionBtn = document.getElementById('btnExecuteSheetAction');
  const origText = actionBtn.textContent;
  actionBtn.disabled = true;
  actionBtn.textContent = '불러오는 중... ⏳';

  try {
    const res = await fetch(scriptUrl);
    if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
    const data = await res.json();
    
    if (Array.isArray(data) && data.length > 0) {
      // Merge with teams data
      teamsData = data.map((item, idx) => ({
        id: item.id || (idx + 1),
        teamNumber: item.teamNumber || `${idx + 1}모둠`,
        countryName: item.countryName || '',
        sentence: item.sentence || '',
        updatedAt: item.updatedAt || new Date().toISOString()
      }));
      
      saveToLocalStorage();
      renderApp();
      closeSheetModal();
      showToast('🎉 구글 스프레드시트로부터 모둠 결과를 성공적으로 불러왔습니다!');
    } else {
      throw new Error('시트 데이터 형식이 유효하지 않습니다.');
    }
  } catch (err) {
    console.error(err);
    alert(`시트 불러오기 실패: ${err.message}\n(배포된 Apps Script 주소 및 CORS 설정을 확인해 주세요.)`);
  } finally {
    actionBtn.disabled = false;
    actionBtn.textContent = origText;
  }
}

/**
 * Execute Sheet Save via Apps Script Web App
 */
async function executeSheetSave() {
  const scriptUrl = saveSheetUrl();

  if (!scriptUrl) {
    alert('아직 구글 Apps Script Web App URL이 설정되지 않았습니다.\n아래 [JSON 파일로 백업 저장]을 누르면 즉시 파일로 컴퓨터에 저장할 수 있습니다!');
    return;
  }

  const actionBtn = document.getElementById('btnExecuteSheetAction');
  const origText = actionBtn.textContent;
  actionBtn.disabled = true;
  actionBtn.textContent = '저장하는 중... ⏳';

  try {
    const res = await fetch(scriptUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        project: '지구마불세계여행',
        savedAt: new Date().toISOString(),
        teams: teamsData
      })
    });

    if (!res.ok) throw new Error(`HTTP status: ${res.status}`);
    
    closeSheetModal();
    showToast('🎉 구글 스프레드시트에 성공적으로 전송·저장되었습니다!');
  } catch (err) {
    console.error(err);
    alert(`시트 저장 실패: ${err.message}\n(네트워크 또는 Apps Script 배포 접근 권한을 확인해 주세요.)`);
  } finally {
    actionBtn.disabled = false;
    actionBtn.textContent = origText;
  }
}

/**
 * Quick JSON Backup (Local export & import without waiting for Apps script setup)
 */
function exportDataAsJson() {
  const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(teamsData, null, 2));
  const downloadAnchor = document.createElement('a');
  downloadAnchor.setAttribute("href", dataStr);
  const now = new Date().toISOString().slice(0, 10);
  downloadAnchor.setAttribute("download", `지구마불세계여행_모둠탐구_${now}.json`);
  document.body.appendChild(downloadAnchor);
  downloadAnchor.click();
  downloadAnchor.remove();
  showToast('💾 백업 파일(JSON)이 다운로드되었습니다.');
}

function importDataFromJson() {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.json';
  input.onchange = e => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = event => {
      try {
        const parsed = JSON.parse(event.target.result);
        if (Array.isArray(parsed)) {
          teamsData = parsed;
          saveToLocalStorage();
          renderApp();
          showToast('📂 파일을 성공적으로 불러와 적용했습니다!');
          closeSheetModal();
        } else {
          alert('올바른 모둠 데이터 파일이 아닙니다.');
        }
      } catch (err) {
        alert('파일을 읽는 도중 오류가 발생했습니다: ' + err.message);
      }
    };
    reader.readAsText(file);
  };
  input.click();
}

/**
 * Fill Demo / Sample Data for quick teacher testing
 */
function fillSampleData() {
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

  saveToLocalStorage();
  renderApp();
  showToast('✨ 수업 시연용 샘플 나라 및 일반화 문장이 채워졌습니다!');
}

/**
 * Reset all 14 teams
 */
function resetAllTeams() {
  const confirmed = confirm('전체 14개 모둠의 모든 입력 내용을 초기화하시겠습니까? (이 작업은 되돌릴 수 없습니다)');
  if (confirmed) {
    teamsData = generateDefaultTeams();
    saveToLocalStorage();
    renderApp();
    showToast('전체 모둠이 깨끗하게 초기화되었습니다.');
  }
}

/**
 * Keyboard shortcuts setup (Escape to close modals, Arrow keys to navigate zoom)
 */
function setupEventListeners() {
  document.addEventListener('keydown', (e) => {
    const zoomModal = document.getElementById('zoomModal');
    const isZoomOpen = zoomModal && zoomModal.classList.contains('open');

    if (e.key === 'Escape') {
      if (isZoomOpen) {
        closeZoomModal();
      } else {
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
