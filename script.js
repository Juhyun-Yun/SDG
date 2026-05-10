/* 
  SDG's 리틀 히어로 - Logic V6 (Heart Rating & Custom Tasks)
*/

const APP_STATE = {
  heroName: '',
  selectedMissions: [], // [ { themeId, taskId, taskText } ]
  startDate: '',
  endDate: '',
  myGoal: '',
  progress: 0,
  currentStep: 0,
  history: [],
  gasUrl: 'https://script.google.com/macros/s/AKfycbxDnb7Cb9h0ceBXLiojZPSshoqBBtdxnfJ5SCMOcv8laAkyQUBzTv4Hd5JhWGLBeU4B/exec'
};

const THEME_DETAILS = {
  1: { title: "빈곤 퇴치", icon: "💰", color: "#e5243b", tasks: ["안 쓰는 물건 기부하기", "용돈 아껴 쓰기", "어려운 이웃에게 응원 편지 쓰기"] },
  2: { title: "기아 종식", icon: "🍱", color: "#dda63a", tasks: ["음식 남기지 않고 다 먹기", "골고루 먹기", "우리 농산물 사랑하기"] },
  3: { title: "건강과 웰빙", icon: "🏃", color: "#4c9f38", tasks: ["매일 30분 운동하기", "손 씻기 잘하기", "일찍 자고 일찍 일어나기"] },
  4: { title: "양질의 교육", icon: "📚", color: "#c5192d", tasks: ["매일 책 한 권 읽기", "수업 시간에 집중하기", "친구와 지식 나누기"] },
  5: { title: "성평등", icon: "⚖️", color: "#ff3a21", tasks: ["편견 없는 말 사용하기", "집안일 스스로 돕기", "친구의 개성 존중하기"] },
  6: { title: "깨끗한 물과 위생", icon: "💧", color: "#26bde2", tasks: ["양치할 때 컵 사용하기", "물 아껴 쓰기", "샴푸 적게 쓰기"] },
  7: { title: "모두를 위한 에너지", icon: "💡", color: "#fcc30b", tasks: ["안 쓰는 전등 끄기", "TV 시청 시간 줄이기", "계단 이용하기"] },
  8: { title: "양질의 일자리", icon: "📈", color: "#a21942", tasks: ["꿈을 향해 노력하기", "정직한 생활 하기", "직업 체험 해보기"] },
  9: { title: "산업·혁신·인프라", icon: "🏗️", color: "#fd6925", tasks: ["새로운 발명 아이디어 내기", "대중교통 이용하기", "인터넷 예절 지키기"] },
  10: { title: "불평등 완화", icon: "🤝", color: "#dd1367", tasks: ["차별하는 말 하지 않기", "친구에게 먼저 다가가기", "도움 필요한 친구 돕기"] },
  11: { title: "지속가능한 도시", icon: "🏘️", color: "#fd9d24", tasks: ["동네 쓰레기 줍기", "자전거 이용하기", "우리 동네 식물 가꾸기"] },
  12: { title: "책임감 있는 소비", icon: "♻️", color: "#bf8b2e", tasks: ["분리배출 철저히 하기", "일회용품 사용 줄이기", "물건 아껴 쓰기"] },
  13: { title: "기후 위기 대응", icon: "🌡️", color: "#3f7e44", tasks: ["기온 변화에 관심 갖기", "냉방/난방 줄이기", "환경 일기 쓰기"] },
  14: { title: "해양생태계 보전", icon: "🌊", color: "#0a97d9", tasks: ["해변가 쓰레기 줍기", "미세 플라스틱 줄이기", "바다 동물 보호하기"] },
  15: { title: "육상생태계 보전", icon: "🌳", color: "#56c02b", tasks: ["나무와 꽃 가꾸기", "숲속 생물 보호하기", "종이 아껴 쓰기"] },
  16: { title: "평화·정의·제도", icon: "🕊️", color: "#00689d", tasks: ["친구와 싸우지 않기", "규칙 잘 지키기", "나쁜 말 사용하지 않기"] },
  17: { title: "지구촌 협력", icon: "🌍", color: "#19486a", tasks: ["세계 친구들에게 관심 갖기", "환경 캠페인 참여하기", "함께 목표 실천하기"] }
};

// Ratings [themeId-taskIndex] = score (1-5)
let DAILY_RATINGS = {};

let STEPS = {};

window.onload = () => {
  try {
    // Initialize STEPS inside onload to ensure elements exist
    STEPS = {
      0: document.getElementById('step-login'),
      1: document.getElementById('step-sdg-intro'),
      '1.5': document.getElementById('step-sdg-gallery'),
      2: document.getElementById('step-summary'),
      '2.5': document.getElementById('step-vision'),
      3: document.getElementById('step-dashboard'),
      4: document.getElementById('step-history'),
      5: document.getElementById('step-success'),
      6: document.getElementById('step-teacher')
    };

    // Check for saved state (but don't auto-redirect to Step 3, let them pick name)
    const codeGasUrl = APP_STATE.gasUrl;
    const savedState = localStorage.getItem('sdg_hero_v6');
    if (savedState) {
      try {
        const parsed = JSON.parse(savedState);
        Object.assign(APP_STATE, parsed);
        // Pre-select the name if it exists in APP_STATE (will be applied once list loads)
      } catch (e) {
        localStorage.removeItem('sdg_hero_v6');
      }
    }
    // gasUrl priority: teacher's custom URL (localStorage) > script.js default
    const customGasUrl = localStorage.getItem('sdg_gas_url');
    APP_STATE.gasUrl = customGasUrl || codeGasUrl;

    // Set default dates
    const today = new Date().toISOString().split('T')[0];
    const nextWeek = new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString().split('T')[0];
    if (document.getElementById('start-date')) document.getElementById('start-date').value = today;
    if (document.getElementById('end-date')) document.getElementById('end-date').value = nextWeek;

    // Start at Login
    goToStep(0);
    loadStudentList();
    renderThemesSelection();
    loadSidebarStats();
    setInterval(loadSidebarStats, 120000); // refresh every 2 min
  } catch (e) {
    console.error("Initialization error:", e);
    // If something crashes, try to at least show the login screen
    if (typeof goToStep === 'function') goToStep(0);
  }
};

let CURRENT_PICKER_THEME = null;

function openMissionPicker(themeId) {
  const modal = document.getElementById('mission-picker-modal');
  const title = document.getElementById('picker-title');
  const taskContainer = document.getElementById('picker-tasks');
  const theme = THEME_DETAILS[themeId];

  if (!theme) return;

  CURRENT_PICKER_THEME = themeId;
  modal.style.display = 'flex';

  const count = APP_STATE.selectedMissions.length;
  title.innerText = `${theme.icon} ${theme.title} (${count}/3)`;
  title.style.color = theme.color;

  // Inject task buttons
  taskContainer.innerHTML = theme.tasks.map((task, idx) => {
    const isSelected = APP_STATE.selectedMissions.some(m => m.taskText === task);
    const bgStyle = isSelected ? 'background: #dcfce7; border-color: #22c55e;' : `background: white; border: 2px solid ${theme.color}44;`;
    const textPrefix = isSelected ? '✅ ' : '';

    return `
      <button class="btn task-option-btn" style="${bgStyle} color: var(--text-main); text-align: left; padding: 12px 15px; font-size: 0.95rem; margin-bottom: 5px; width: 100%; box-shadow: 0 4px 0 ${isSelected ? '#22c55e22' : theme.color + '22'};" onclick="selectTask(${themeId}, '${task}', event)">
        ${textPrefix}${idx + 1}. ${task}
      </button>
    `;
  }).join('');

  // Clear custom input
  document.getElementById('custom-task-input').value = '';
}

function closeMissionPicker(event) {
  // If event is provided, only close if clicking the backdrop
  if (event && event.target !== document.getElementById('mission-picker-modal')) return;
  document.getElementById('mission-picker-modal').style.display = 'none';
}

function selectTask(themeId, taskText, event) {
  // Check limit (max 3)
  if (APP_STATE.selectedMissions.length >= 3) {
    showToast("이미 3개의 미션을 모두 선택하셨습니다! 😊");
    return;
  }

  // Check if already selected
  const isAlreadySelected = APP_STATE.selectedMissions.some(m => m.taskText === taskText);
  if (isAlreadySelected) {
    showToast("이미 선택한 미션입니다! 😊");
    return;
  }

  APP_STATE.selectedMissions.push({ themeId, taskText });
  saveState();

  const currentCount = APP_STATE.selectedMissions.length;
  showToast(`미션이 추가되었습니다! (${currentCount}/3) ✨`);

  // Show the next button on the MAIN screen
  const mainNextBtn = document.getElementById('btn-next-step');
  if (mainNextBtn) mainNextBtn.style.display = 'block';

  // Always close picker immediately on selection as requested
  setTimeout(() => {
    closeMissionPicker();
    // If it was the 3rd one, auto-advance to summary
    if (currentCount === 3) goToStep(2);
  }, 400);
}

function addCustomTask(themeId) {
  let input, targetTheme;

  // Case 1: Called from the Modal (no themeId passed)
  if (themeId === undefined || themeId === null || typeof themeId === 'object') {
    input = document.getElementById('custom-task-input');
    targetTheme = CURRENT_PICKER_THEME;
  }
  // Case 2: Called from somewhere else with a themeId
  else {
    input = document.getElementById(`add-task-input-${themeId}`);
    targetTheme = themeId;
  }

  const text = input ? input.value.trim() : '';
  if (!text || !targetTheme) return;

  if (APP_STATE.selectedMissions.length >= 3) {
    showToast("이미 3개의 미션을 모두 선택하셨습니다! 😊");
    return;
  }

  APP_STATE.selectedMissions.push({ themeId: targetTheme, taskText: text });
  saveState();

  const currentCount = APP_STATE.selectedMissions.length;
  showToast(`나만의 미션이 추가되었습니다! (${currentCount}/3) ✨`);

  // Show the next button on the MAIN screen
  const mainNextBtn = document.getElementById('btn-next-step');
  if (mainNextBtn) mainNextBtn.style.display = 'block';

  // Always close picker
  closeMissionPicker();

  // If it's the 3rd one, go to summary
  if (currentCount === 3) {
    setTimeout(() => goToStep(2), 500);
  }

  if (input) input.value = '';
  if (themeId && typeof themeId !== 'object') renderThemesSelection();
}

function updateMissionSummary() {
  const list = document.getElementById('summary-list');
  const nextBtn = document.getElementById('btn-next-step');
  const countSpan = document.getElementById('mission-count');

  const count = APP_STATE.selectedMissions.length;
  if (countSpan) countSpan.innerText = `${count} / 3`;

  if (count === 0) {
    list.innerHTML = '<li style="color: #94a3b8; text-align: center; padding: 20px;">아직 선택한 미션이 없습니다.<br>위의 아이콘을 눌러 미션을 골라보세요!</li>';
    nextBtn.style.display = 'none';
    return;
  }

  list.innerHTML = APP_STATE.selectedMissions.map((m, idx) => `
    <li style="margin-bottom: 12px; display: flex; justify-content: space-between; align-items: center; background: #fff; padding: 15px 20px; border-radius: 16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); border: 1px solid #f1f5f9;">
      <div style="display: flex; flex-direction: column;">
        <span style="font-size: 0.75rem; color: var(--text-muted); margin-bottom: 2px;">${THEME_DETAILS[m.themeId] ? THEME_DETAILS[m.themeId].icon + ' ' + THEME_DETAILS[m.themeId].title : '✨ 미션'}</span>
        <span style="font-weight: 800; color: var(--text-main); font-size: 1rem;">${m.taskText}</span>
      </div>
      <button onclick="removeMission(${idx})" style="background: #fee2e2; border: none; color: #ef4444; font-size: 0.8rem; padding: 6px 12px; border-radius: 10px; cursor: pointer; font-weight: 800; transition: all 0.2s;">삭제</button>
    </li>
  `).join('');

  nextBtn.style.display = 'block';
  saveState();
}

function removeMission(index) {
  APP_STATE.selectedMissions.splice(index, 1);
  updateMissionSummary();
}

function handleLogin() {
  const nameSelect = document.getElementById('hero-name');
  const name = nameSelect.value;

  if (!name) {
    showToast("이름을 선택해주세요! 👋");
    return;
  }

  // Load student-specific config if exists
  const configKey = 'sdg_hero_config_' + name;
  const savedConfig = localStorage.getItem(configKey);

  if (savedConfig) {
    try {
      const parsed = JSON.parse(savedConfig);
      // Explicitly reset APP_STATE and then apply parsed data
      APP_STATE.selectedMissions = parsed.selectedMissions || [];
      APP_STATE.startDate = parsed.startDate || '';
      APP_STATE.endDate = parsed.endDate || '';
      APP_STATE.myGoal = parsed.myGoal || '';
      APP_STATE.progress = parsed.progress || 0;
    } catch (e) {
      console.error("Config parse error for", name, e);
    }
  } else {
    // Fresh start for a new student or one with no saved config
    APP_STATE.selectedMissions = [];
    APP_STATE.startDate = '';
    APP_STATE.endDate = '';
    APP_STATE.myGoal = '';
    APP_STATE.progress = 0;
  }

  APP_STATE.heroName = name;
  saveState();

  // Transition effect
  const btn = document.querySelector('#step-login .btn');
  const loader = document.getElementById('login-loader');
  if (btn) btn.disabled = true;
  if (loader) loader.style.display = 'block';

  setTimeout(() => {
    // Reset loader/button state
    if (btn) btn.disabled = false;
    if (loader) loader.style.display = 'none';

    // Date Logic for Redirection
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let isChallengeActive = false;
    if (APP_STATE.endDate && APP_STATE.selectedMissions.length > 0) {
      const end = new Date(APP_STATE.endDate);
      end.setHours(0, 0, 0, 0);
      if (today <= end) isChallengeActive = true;
    }

    if (isChallengeActive) {
      goToStep(3); // Go straight to Dashboard (Growth Records)
      showToast(`${name} 히어로, 어서오세요! 오늘의 성장을 기록해볼까요? ✨`);
    } else {
      goToStep(1); // Move to Intro/Selection for new or finished challenges
    }
  }, 800);
}

async function loadStudentList() {
  const select = document.getElementById('hero-name');
  if (!select) return;

  if (APP_STATE.gasUrl.includes('PASTE_YOUR_GAS_WEB_APP_URL_HERE') || !APP_STATE.gasUrl.startsWith('https')) {
    select.innerHTML = '<option value="">서버 URL 설정이 필요합니다.</option>';
    return;
  }

  try {
    select.innerHTML = '<option value="">명단 불러오는 중... 📡</option>';
    console.log("Fetching from:", APP_STATE.gasUrl);

    const response = await fetch(APP_STATE.gasUrl + '?action=getStudents');
    if (!response.ok) throw new Error('서버 응답 오류 (HTTP ' + response.status + ')');

    const data = await response.json();
    if (data.students && data.students.length > 0) {
      select.innerHTML = '<option value="">이름을 선택해주세요.</option>' +
        data.students.map(name => `<option value="${name}">${name}</option>`).join('');

      // Ensure no name is pre-selected
      select.value = "";

      console.log("Students loaded successfully:", data.students.length);
    } else {
      select.innerHTML = '<option value="">학생 명단이 비어 있습니다.</option>';
    }
  } catch (err) {
    console.error("Student load error:", err);
    select.innerHTML = '<option value="">서버 연결 실패 (URL 확인 필요)</option>';
    alert("명단 로딩 중 문제가 발생했습니다.\n- 구글 앱스 스크립트 배포 확인\n- 인터넷 연결 확인\n- 오류 내용: " + err.message);
  }
}

function goToStep(step) {
  if (step === 6) {
    const pw = prompt("선생님 암호를 입력하세요");
    if (pw !== '1234') {
      if (pw !== null) alert("암호가 틀렸습니다.");
      return;
    }
  }

  // Hide all
  Object.values(STEPS).forEach(el => { if (el) el.style.display = 'none'; });

  if (STEPS[step]) {
    // Preserve flex for landing screens
    if (step === 0) {
      STEPS[step].style.display = 'flex';
    } else {
      STEPS[step].style.display = 'block';
    }
  }

  if (step === 2) {
    const goalArea = document.getElementById('my-goal');
    if (goalArea) goalArea.value = APP_STATE.myGoal || '';

    const chipsContainer = document.getElementById('selection-chips');
    if (chipsContainer) {
      chipsContainer.innerHTML = APP_STATE.selectedMissions.map(m => {
        const t = THEME_DETAILS[m.themeId];
        return `
          <div style="background:white; border:1px solid ${t.color}66; padding:15px; border-radius:15px; width:100%; display:flex; align-items:center; gap:12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); margin-bottom: 5px;">
            <div style="background:${t.color}15; padding:10px; border-radius:12px; font-size:1.5rem;">${t.icon}</div>
            <div style="flex:1; min-width:0;">
              <span style="font-size:0.75rem; color:var(--text-muted); display:block; font-weight:700; margin-bottom:4px;">${t.title}</span>
              <span style="font-size:1rem; font-weight:800; color:var(--text-main); line-height:1.2; word-break: keep-all;">${m.taskText}</span>
            </div>
          </div>
        `;
      }).join('');
    }
  }
  if (step === 2.5) {
    // Sync current values if any
    if (APP_STATE.startDate) document.getElementById('start-date').value = APP_STATE.startDate;
    if (APP_STATE.endDate) document.getElementById('end-date').value = APP_STATE.endDate;
    updateDurationDisplay();
  }
  if (step === 3) {
    updateDashboard();
  }
  if (step === 6) {
    switchTeacherTab('stats');
  }

  if (STEPS[step]) STEPS[step].style.display = 'block';
  APP_STATE.currentStep = step;
  saveState();
}

function switchTeacherTab(tabName) {
  document.querySelectorAll('.teacher-tab').forEach(t => {
    t.classList.toggle('active', t.dataset.tab === tabName);
  });
  document.querySelectorAll('.teacher-tab-content').forEach(c => {
    c.style.display = (c.id === 'teacher-tab-' + tabName) ? 'block' : 'none';
  });
  if (tabName === 'stats') loadTeacherStats();
  if (tabName === 'settings') populateSettingsInputs();
}

function populateSettingsInputs() {
  const gasInput = document.getElementById('gas-url-input');
  const sheetInput = document.getElementById('spreadsheet-url-input');
  const currentEl = document.getElementById('gas-url-current');
  if (gasInput) gasInput.value = APP_STATE.gasUrl || '';
  if (sheetInput) sheetInput.value = localStorage.getItem('sdg_spreadsheet_url') || '';
  if (currentEl) {
    const isCustom = !!localStorage.getItem('sdg_gas_url');
    const tag = isCustom ? '(설정값)' : '(기본값)';
    currentEl.innerText = (APP_STATE.gasUrl || '(없음)') + ' ' + tag;
  }
}

function resetGasUrlToDefault() {
  if (!confirm('저장된 본부 URL 설정을 지우고 기본값으로 되돌립니다. 진행할까요?')) return;
  localStorage.removeItem('sdg_gas_url');
  alert('✅ 기본값으로 복원되었습니다. 페이지를 새로고침합니다.');
  location.reload();
}

function saveGasUrl() {
  const input = document.getElementById('gas-url-input');
  const url = input.value.trim();
  if (!url) {
    alert('URL을 입력해주세요.');
    return;
  }
  if (!url.startsWith('https://script.google.com/')) {
    alert('올바른 Apps Script URL이 아닙니다.\nhttps://script.google.com/ 으로 시작해야 해요.');
    return;
  }
  if (!url.endsWith('/exec')) {
    if (!confirm('이 URL은 일반적인 Web App 형식(/exec로 끝남)이 아닙니다.\n그래도 저장할까요?')) return;
  }
  localStorage.setItem('sdg_gas_url', url);
  APP_STATE.gasUrl = url;
  alert('✅ 본부 URL이 저장되었습니다.\n페이지를 새로고침하면 새 학급 데이터가 표시됩니다.');
}

function saveSpreadsheetUrl() {
  const input = document.getElementById('spreadsheet-url-input');
  const url = input.value.trim();
  if (!url) {
    alert('URL을 입력해주세요.');
    return;
  }
  if (!url.startsWith('https://docs.google.com/spreadsheets/')) {
    alert('올바른 Google Sheets URL이 아닙니다.\nhttps://docs.google.com/spreadsheets/ 으로 시작해야 해요.');
    return;
  }
  localStorage.setItem('sdg_spreadsheet_url', url);
  alert('✅ 스프레드시트 주소가 저장되었습니다.');
}

async function loadSidebarStats() {
  if (!APP_STATE.gasUrl || !APP_STATE.gasUrl.startsWith('https')) return;
  try {
    const res = await fetch(APP_STATE.gasUrl + '?action=getClassStats');
    if (!res.ok) return;
    const text = await res.text();
    let data;
    try { data = JSON.parse(text); } catch (e) { return; }
    if (!data.stats) return;
    const s = data.stats;
    const setText = (id, v) => { const el = document.getElementById(id); if (el) el.innerText = v; };
    setText('sb-students', s.totalStudents);
    setText('sb-records', s.totalRecords);
    setText('sb-tasks', s.totalTasks);
    setText('sb-avg', s.avgProgress + '%');
    let msg;
    if (s.totalRecords >= 30) msg = '🌟 우리 반은 진짜 히어로!';
    else if (s.totalRecords >= 10) msg = '🌱 멋진 실천이 쌓이고 있어요!';
    else if (s.totalRecords > 0) msg = '✨ 시작이 반! 함께 해요!';
    else msg = '함께 만드는 멋진 변화 🌱';
    setText('sb-msg', msg);
  } catch (err) { /* silent */ }
}

async function loadTeacherStats() {
  const loadingEl = document.getElementById('class-stats-loading');
  const contentEl = document.getElementById('class-stats-content');
  if (!loadingEl || !contentEl) return;

  loadingEl.style.display = 'block';
  loadingEl.innerHTML = '📡 데이터를 불러오는 중...';
  contentEl.style.display = 'none';

  if (APP_STATE.gasUrl === 'PASTE_YOUR_GAS_WEB_APP_URL_HERE' || !APP_STATE.gasUrl.startsWith('https')) {
    loadingEl.innerHTML = '⚠️ 서버 URL이 설정되지 않았습니다.';
    return;
  }

  try {
    const res = await fetch(APP_STATE.gasUrl + '?action=getClassStats');
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const text = await res.text();
    let data;
    try {
      data = JSON.parse(text);
    } catch (parseErr) {
      console.error('JSON parse failed. Raw response:', text.substring(0, 200));
      loadingEl.innerHTML = '⚠️ 서버가 통계 기능을 인식하지 못했습니다.<br><br><b>해결 방법</b><br>1. Google Apps Script 편집기 열기<br>2. <b>배포 → 배포 관리 → ✏️ 편집</b><br>3. 버전을 <b>"새 버전"</b>으로 선택 후 배포<br>4. 같은 URL이 유지되면 그대로 사용';
      return;
    }
    if (!data.stats || data.stats.totalRecords === 0) {
      loadingEl.innerHTML = '아직 우리 반의 기록이 없어요. 친구들이 실천을 시작하면 통계가 나타납니다! ✨';
      return;
    }
    renderClassStats(data.stats);
    loadingEl.style.display = 'none';
    contentEl.style.display = 'block';
  } catch (err) {
    console.error('Class stats error:', err);
    const safeUrl = (APP_STATE.gasUrl || '').replace(/[<>&"]/g, c => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;'}[c]));
    loadingEl.innerHTML = '❌ 데이터를 불러올 수 없습니다.<br>설정 탭에서 본부 URL을 확인해주세요.<br><span style="font-size:0.65rem; color:#94a3b8; word-break:break-all; display:block; margin-top:6px;">사용 중 URL: ' + safeUrl + '<br>오류: ' + (err.message || err) + '</span>';
  }
}

function renderClassStats(stats) {
  document.getElementById('class-total-students').innerText = stats.totalStudents;
  document.getElementById('class-total-records').innerText = stats.totalRecords;
  document.getElementById('class-total-tasks').innerText = stats.totalTasks;
  document.getElementById('class-avg-progress').innerText = stats.avgProgress + '%';
  renderDailyChart(stats.dailyStats || []);
}

function renderDailyChart(dailyStats) {
  const chartEl = document.getElementById('daily-chart');
  const emptyEl = document.getElementById('daily-chart-empty');
  if (!chartEl) return;

  const recent = dailyStats.slice(-10);
  if (recent.length === 0) {
    chartEl.style.display = 'none';
    if (emptyEl) emptyEl.style.display = 'block';
    return;
  }
  chartEl.style.display = 'block';
  if (emptyEl) emptyEl.style.display = 'none';

  const w = Math.max(260, chartEl.clientWidth || 280);
  const h = 120;
  const padL = 26, padR = 6, padT = 8, padB = 22;
  const innerW = w - padL - padR;
  const innerH = h - padT - padB;
  const COLOR = '#6ee7b7';

  // Y-axis is fixed 0~100% (avg progress) regardless of data range
  const yMax = 100;
  const yTicks = 4;

  // Use avgProgress if present (new GAS), fall back to deriving from totalTasks if needed
  const hasAvg = recent.some(d => typeof d.avgProgress === 'number');
  if (!hasAvg) {
    console.warn('[chart] avgProgress 필드 없음 — GAS 재배포가 필요합니다.');
  }
  const points = recent.map((d, i) => {
    const val = (typeof d.avgProgress === 'number') ? d.avgProgress : 0;
    const x = padL + (recent.length === 1 ? innerW / 2 : (i * innerW / (recent.length - 1)));
    const y = padT + innerH - (val / yMax) * innerH;
    return { x: x, y: y, val: val, label: d.label };
  });

  // If only 1 point, draw a short horizontal line so it's visible
  if (points.length === 1) {
    const p = points[0];
    points.unshift({ x: padL + 8, y: p.y, val: p.val, label: '' });
    points.push({ x: padL + innerW - 8, y: p.y, val: p.val, label: '' });
  }

  const pathD = points.map((p, i) => (i === 0 ? `M ${p.x} ${p.y}` : `L ${p.x} ${p.y}`)).join(' ');
  const areaD = points.length > 1
    ? pathD + ` L ${points[points.length - 1].x} ${padT + innerH} L ${points[0].x} ${padT + innerH} Z`
    : '';

  let yAxisSvg = '';
  for (let i = 0; i <= yTicks; i++) {
    const val = Math.round(yMax * (yTicks - i) / yTicks);
    const y = padT + (i * innerH / yTicks);
    yAxisSvg += `<line x1="${padL}" y1="${y}" x2="${w - padR}" y2="${y}" stroke="#f1f5f9" stroke-width="0.4"/>`;
    yAxisSvg += `<text x="${padL - 3}" y="${y + 2}" text-anchor="end" font-size="7" fill="#94a3b8" font-weight="400">${val}%</text>`;
  }

  const showEvery = recent.length > 6 ? 2 : 1;
  const pointsSvg = points.map((p, i) => {
    const showLabel = i % showEvery === 0 || i === points.length - 1;
    return `
      <circle cx="${p.x}" cy="${p.y}" r="2" fill="white" stroke="${COLOR}" stroke-width="1"/>
      ${showLabel ? `<text x="${p.x}" y="${p.y - 5}" text-anchor="middle" font-size="7" font-weight="500" fill="#64748b">${p.val}%</text>` : ''}
    `;
  }).join('');

  const xLabelsSvg = points.map((p, i) => {
    const showLabel = i % showEvery === 0 || i === points.length - 1;
    if (!showLabel) return '';
    return `
      <text x="${p.x}" y="${padT + innerH + 10}" text-anchor="middle" font-size="7.5" fill="#94a3b8" font-weight="500">${p.label}</text>
    `;
  }).join('');

  chartEl.innerHTML = `
    <svg width="100%" height="${h}" viewBox="0 0 ${w} ${h}" style="display:block; overflow:visible;">
      ${yAxisSvg}
      <line x1="${padL}" y1="${padT}" x2="${padL}" y2="${padT + innerH}" stroke="#e2e8f0" stroke-width="0.6"/>
      <line x1="${padL}" y1="${padT + innerH}" x2="${w - padR}" y2="${padT + innerH}" stroke="#e2e8f0" stroke-width="0.6"/>
      ${areaD ? `<path d="${areaD}" fill="${COLOR}" fill-opacity="0.12"/>` : ''}
      ${points.length > 1 ? `<path d="${pathD}" fill="none" stroke="${COLOR}" stroke-width="1" stroke-linecap="round" stroke-linejoin="round"/>` : ''}
      ${pointsSvg}
      ${xLabelsSvg}
    </svg>
  `;
}

function confirmMissions() {
  const startDate = document.getElementById('start-date').value;
  const endDate = document.getElementById('end-date').value;
  const myGoal = document.getElementById('my-goal').value;

  if (!startDate || !endDate) {
    alert("챌린지 시작일과 종료일을 설정해주세요! 📅");
    return;
  }

  const start = new Date(startDate);
  const end = new Date(endDate);
  if (end < start) {
    alert("종료일은 시작일보다 빠를 수 없어요! 😅");
    return;
  }

  APP_STATE.startDate = startDate;
  APP_STATE.endDate = endDate;
  APP_STATE.myGoal = myGoal;

  saveState();
  goToStep(3);
  showToast("미션 본부에 챌린지 계획이 전송되었습니다! 🚀");
}

function renderThemesSelection() {
  const container = document.getElementById('themes-list');
  if (!container) return;
  container.innerHTML = Object.entries(THEME_DETAILS).map(([id, detail]) => `
    <div class="theme-card" style="border-left:8px solid ${detail.color};">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
        <span style="font-size:1.3rem;">${detail.icon} <b>${detail.title}</b></span>
      </div>
      <div class="task-selection-list">
        ${detail.tasks.map((task, idx) => {
    const isSelected = APP_STATE.selectedMissions.some(m => m.themeId === id && m.taskId === 'p' + idx);
    return `
            <label style="display:flex; align-items:center; gap:10px; font-size:0.9rem; margin-bottom:8px; cursor:pointer;">
              <input type="checkbox" ${isSelected ? 'checked' : ''} onchange="toggleTask('${id}', 'p${idx}', '${task}', this.checked)" style="width:18px; height:18px;">
              <span>${task}</span>
            </label>
          `;
  }).join('')}
        <div id="custom-tasks-${id}" style="margin-top:10px;">
          ${APP_STATE.selectedMissions.filter(m => m.themeId === id && m.taskId && m.taskId.startsWith('c')).map(m => `
            <label style="display:flex; align-items:center; gap:10px; font-size:0.9rem; margin-bottom:8px; cursor:pointer; color:${detail.color}; font-weight:700;">
              <input type="checkbox" checked onchange="toggleTask('${id}', '${m.taskId}', '${m.taskText}', false)" style="width:18px; height:18px;">
              <span>(나만의 실천) ${m.taskText}</span>
            </label>
          `).join('')}
        </div>
      </div>
      <div class="custom-task-box">
        <div style="display:flex; gap:8px;">
          <input type="text" id="add-task-input-${id}" placeholder="직접 쓰고 싶은 실천 내용이 있나요?" style="padding:10px; font-size:0.85rem; border-radius:10px; flex:1;">
          <button onclick="addCustomTask('${id}')" style="background:${detail.color}; color:white; border:none; padding:10px 15px; border-radius:10px; cursor:pointer; font-weight:bold; white-space:nowrap;">추가</button>
        </div>
      </div>
    </div>
  `).join('');
}

function toggleTask(themeId, taskId, taskText, isChecked) {
  if (isChecked) {
    if (APP_STATE.selectedMissions.length >= 3) {
      renderThemesSelection();
      return;
    }
    const alreadyInTheme = APP_STATE.selectedMissions.some(m => m.themeId === themeId && m.taskId === taskId);
    if (alreadyInTheme) {
      renderThemesSelection();
      return;
    }
    APP_STATE.selectedMissions.push({ themeId, taskId, taskText });
  } else {
    APP_STATE.selectedMissions = APP_STATE.selectedMissions.filter(m => !(m.themeId === themeId && m.taskId === taskId));
  }
  saveState();
  renderThemesSelection();
}



function updateDashboard() {
  document.getElementById('display-name').innerText = APP_STATE.heroName + " 히어로";

  // Format period
  if (APP_STATE.startDate && APP_STATE.endDate) {
    const s = APP_STATE.startDate.replace(/-/g, '.');
    const e = APP_STATE.endDate.replace(/-/g, '.');
    const periodElem = document.getElementById('display-period');
    if (periodElem) periodElem.innerText = `${s} ~ ${e}`;
  }

  // Calculate Days
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const start = new Date(APP_STATE.startDate);
  start.setHours(0, 0, 0, 0);
  const end = new Date(APP_STATE.endDate);
  end.setHours(0, 0, 0, 0);

  const totalTime = end.getTime() - start.getTime();
  const totalDays = Math.ceil(totalTime / (1000 * 60 * 60 * 24)) + 1;

  const elapsedTime = today.getTime() - start.getTime();
  const currentDay = Math.ceil(elapsedTime / (1000 * 60 * 60 * 24)) + 1;

  const dayLabel = document.getElementById('current-day-label');
  if (dayLabel) {
    if (currentDay < 1) {
      dayLabel.innerText = "챌린지 준비 중";
    } else if (currentDay > totalDays) {
      dayLabel.innerText = "챌린지 완료! 🏆";
    } else {
      dayLabel.innerText = `Day ${currentDay} / ${totalDays}`;
    }
  }

  const container = document.getElementById('active-missions');
  if (!container) return;
  container.classList.add('active-missions-area');

  container.innerHTML = APP_STATE.selectedMissions.map((mission, idx) => {
    const detail = THEME_DETAILS[mission.themeId];
    return `
      <div class="mission-mini-card" style="border-left: 6px solid ${detail.color};">
        <div style="flex: 1; min-width: 0;">
          <span style="font-size:0.65rem; color:var(--text-muted); display: block;">${detail.icon} ${detail.title}</span>
          <p style="font-weight:800; font-size:0.95rem; color:var(--text-main); line-height:1.2; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${mission.taskText}</p>
        </div>
        <div class="heart-rating-mini" data-midx="${idx}">
          <span onclick="rateTask(${idx}, 1)">❤</span>
          <span onclick="rateTask(${idx}, 2)">❤</span>
          <span onclick="rateTask(${idx}, 3)">❤</span>
          <span onclick="rateTask(${idx}, 4)">❤</span>
          <span onclick="rateTask(${idx}, 5)">❤</span>
        </div>
      </div>
    `;
  }).join('');
}

function rateTask(mIdx, rating) {
  const container = document.querySelector(`.heart-rating-mini[data-midx="${mIdx}"]`);
  const hearts = container.querySelectorAll('span');
  hearts.forEach((h, i) => {
    h.style.color = (i < rating) ? '#ef4444' : '#e2e8f0';
    h.style.transform = (i < rating) ? 'scale(1.2)' : 'scale(1)';
  });
  DAILY_RATINGS[mIdx] = rating;
  updateProgress();
  debounceAutoSave();
}

function updateProgress() {
  const allRatings = Object.values(DAILY_RATINGS);
  if (allRatings.length === 0) return;
  const totalPotential = (document.querySelectorAll('.heart-rating-mini').length) * 5;
  const totalScore = allRatings.reduce((a, b) => a + b, 0);
  const percentage = Math.round((totalScore / totalPotential) * 100);
  APP_STATE.progress = percentage;
  document.getElementById('total-progress').style.width = percentage + '%';
  document.getElementById('progress-text').innerText = percentage + '%';
}

let saveTimeout;
function debounceAutoSave() {
  clearTimeout(saveTimeout);
  document.getElementById('save-status').innerText = '✍️ 기록을 정리 중이에요...';
  saveTimeout = setTimeout(autoSave, 2000);
}

async function autoSave() {
  updateProgress();
  const status = document.getElementById('save-status');
  status.innerHTML = '☁️ <span style="animation: pulse 1s infinite;">기록을 본부로 보내는 중...</span>';

  const logDataTasks = APP_STATE.selectedMissions.map((m, idx) => {
    const rating = DAILY_RATINGS[idx] || 0;
    const detail = THEME_DETAILS[m.themeId];
    return {
      theme: detail.title, icon: detail.icon, task: m.taskText, score: rating
    };
  });

  // Format as a single string for Spreadsheet display (e.g., "기후 위기 대응, 안 쓰는 전기 끄기, 2")
  const tasksDisplayString = logDataTasks.map(t => `${t.theme}, ${t.task}, ${t.score}`).join('\n');

  const reflection = document.getElementById('daily-reflection').value.trim();

  const logData = {
    heroName: APP_STATE.heroName, progress: APP_STATE.progress,
    tasks: logDataTasks, reflection: reflection, myGoal: APP_STATE.myGoal,
    timestamp: new Date().toLocaleString('ko-KR')
  };

  try {
    if (APP_STATE.gasUrl !== 'PASTE_YOUR_GAS_WEB_APP_URL_HERE') {
      // Send the formatted string to GAS instead of objects
      const postData = { ...logData, tasks: tasksDisplayString };
      await fetch(APP_STATE.gasUrl, { method: 'POST', mode: 'no-cors', body: JSON.stringify(postData) });
    }
    const hKey = 'sdg_history_' + APP_STATE.heroName;
    const currentHist = JSON.parse(localStorage.getItem(hKey) || '[]');
    const todayKey = new Date().toLocaleDateString('ko-KR');
    const latest = currentHist[0];
    const latestDateKey = latest && latest.timestamp ? latest.timestamp.startsWith(todayKey) : false;
    if (latest && latestDateKey) {
      currentHist[0] = logData;
    } else {
      currentHist.unshift(logData);
      if (currentHist.length > 50) currentHist.pop();
    }
    localStorage.setItem(hKey, JSON.stringify(currentHist));
    status.innerText = '✅ 본부에 전송되었습니다!';
    setTimeout(() => { status.innerText = ''; }, 3000);
    // Refresh sidebar stats after a brief delay (give GAS time to process)
    setTimeout(loadSidebarStats, 3000);
  } catch (err) { status.innerText = '❌ 연결 실패'; }
}

function resetChallenge() {
  if (confirm("주제나 날짜를 처음부터 다시 정할까요?")) {
    localStorage.removeItem('sdg_hero_v6');
    location.reload();
  }
}
function saveState() {
  localStorage.setItem('sdg_hero_v6', JSON.stringify(APP_STATE));
  if (APP_STATE.heroName) {
    localStorage.setItem('sdg_hero_config_' + APP_STATE.heroName, JSON.stringify(APP_STATE));
  }
}
function resetView() { location.reload(); }
async function toggleHistory() {
  // Flush any pending debounced save so the latest reflection is captured
  if (saveTimeout) {
    clearTimeout(saveTimeout);
    saveTimeout = null;
    const reflection = document.getElementById('daily-reflection');
    const hasRatings = Object.keys(DAILY_RATINGS).length > 0;
    const hasReflection = reflection && reflection.value.trim().length > 0;
    if (hasRatings || hasReflection) {
      await autoSave();
    }
  }
  goToStep(4);
  renderHistory();
}
function renderHistory() {
  const hKey = 'sdg_history_' + APP_STATE.heroName;
  const history = JSON.parse(localStorage.getItem(hKey) || '[]');
  const container = document.getElementById('history-container');

  const goalElem = document.getElementById('history-display-goal');
  const visionBox = document.getElementById('history-vision-box');
  if (APP_STATE.myGoal) {
    if (goalElem) goalElem.innerText = APP_STATE.myGoal;
    if (visionBox) visionBox.style.display = 'flex';
  } else {
    if (visionBox) visionBox.style.display = 'none';
  }

  if (history.length === 0) {
    container.innerHTML = '<div style="text-align:center; padding:40px; color:var(--text-muted);">아직 기록이 없어요. 오늘의 행동을 먼저 기록해보세요!</div>';
    document.getElementById('history-chart-container').innerHTML = '';
    return;
  }

  container.innerHTML = history.map(item => {
    const tasks = Array.isArray(item.tasks) ? item.tasks : [];
    return `
    <div class="card" style="padding: 0; overflow: hidden; margin-bottom: 25px; border: 1px solid #e2e8f0;">
      <div style="background: #f8fafc; padding: 12px 20px; border-bottom: 1px solid #e2e8f0; display: flex; justify-content: space-between; align-items: center;">
        <span style="font-size: 0.8rem; color: #64748b; font-weight: 800;">🕒 ${item.timestamp}</span>
        <span style="background: var(--primary); color: white; padding: 3px 10px; border-radius: 8px; font-size: 0.75rem; font-weight: 800;">성장 ${item.progress}%</span>
      </div>
      
      <table style="width: 100%; border-collapse: collapse; font-size: 0.9rem;">
        <thead>
          <tr style="background: #fff; border-bottom: 1px solid #f1f5f9;">
            <th style="padding: 12px; text-align: center; color: #94a3b8; font-size: 0.75rem; width: 35%;">주제</th>
            <th style="padding: 12px; text-align: left; color: #94a3b8; font-size: 0.75rem;">실천 내용</th>
            <th style="padding: 12px; text-align: center; color: #94a3b8; font-size: 0.75rem; width: 20%;">점수</th>
          </tr>
        </thead>
        <tbody>
          ${tasks.map(t => `
            <tr style="border-bottom: 1px solid #f8fafc;">
              <td style="padding: 12px; font-weight: 700; color: var(--primary); font-size: 0.8rem; text-align: center; word-break: keep-all; line-height: 1.3;">${t.icon} ${t.theme}</td>
              <td style="padding: 12px; color: var(--text-main); font-weight: 700; line-height: 1.3;">${t.task}</td>
              <td style="padding: 12px; text-align: center; font-size: 0.9rem; font-weight: 800; color: #ef4444;">${t.score}점</td>
            </tr>
          `).join('')}
        </tbody>
        ${item.reflection ? `
          <tfoot>
            <tr>
              <td colspan="3" style="padding: 15px 20px; background: #fffbeb; color: #92400e; font-size: 0.85rem; border-top: 1px solid #fef3c7;">
                <div style="display: flex; gap: 10px; align-items: baseline;">
                  <span style="font-weight: 900; color: #b45309; font-size: 0.75rem; white-space: nowrap;">💬 히어로 한 마디:</span>
                  <span style="font-weight: 500; color: #78350f; line-height: 1.5;">${item.reflection}</span>
                </div>
              </td>
            </tr>
          </tfoot>
        ` : ''}
      </table>
    </div>
  `;
  }).join('');

  renderHistoryChart(history);
}

function renderHistoryChart(history) {
  const chartContainer = document.getElementById('history-chart-container');
  if (!chartContainer) return;

  const recent = [...history].reverse().slice(-7);
  if (recent.length === 0) return;

  const w = chartContainer.clientWidth || 300;
  const h = 250;
  const padding = 35;

  // Points calculation
  const points = recent.map((item, idx) => {
    const x = padding + (idx * ((w - padding * 2) / Math.max(1, recent.length - 1)));
    const y = h - padding - (item.progress * ((h - padding * 2) / 100));

    // Parse date from timestamp (format is usually "2026. 4. 26. 오후 9:13:21")
    let dateStr = "";
    try {
      const match = item.timestamp.match(/(\d+)\.\s*(\d+)\.\s*(\d+)/);
      if (match) dateStr = `${match[2]}/${match[3]}`;
      else dateStr = item.timestamp.split('.')[1].trim() + '/' + item.timestamp.split('.')[2].trim();
    } catch (e) { dateStr = (idx + 1).toString(); }

    return { x, y, val: item.progress, date: dateStr };
  });

  const pathD = points.map((p, i) => (i === 0 ? `M ${p.x} ${p.y}` : `L ${p.x} ${p.y}`)).join(' ');
  const areaD = `${pathD} L ${points[points.length - 1].x} ${h - padding} L ${points[0].x} ${h - padding} Z`;

  chartContainer.innerHTML = `
    <svg width="100%" height="${h}" style="overflow: visible;">
      <defs>
        <linearGradient id="areaGrad" x1="0%" y1="0%" x2="0%" y2="100%">
          <stop offset="0%" style="stop-color:var(--primary);stop-opacity:0.2" />
          <stop offset="100%" style="stop-color:var(--primary);stop-opacity:0" />
        </linearGradient>
      </defs>

      <!-- Grid lines & Y-Axis Labels -->
      ${[0, 25, 50, 75, 100].map(val => {
    const y = h - padding - (val * ((h - padding * 2) / 100));
    return `
          <line x1="${padding}" y1="${y}" x2="${w - padding}" y2="${y}" stroke="#e2e8f0" stroke-width="1.2" stroke-dasharray="4" />
          <text x="${padding - 10}" y="${y + 4}" text-anchor="end" font-size="9" fill="#64748b" font-weight="900">${val}%</text>
        `;
  }).join('')}

      <!-- Main Axes -->
      <line x1="${padding}" y1="${h - padding}" x2="${w - padding}" y2="${h - padding}" stroke="#94a3b8" stroke-width="2" />
      <line x1="${padding}" y1="${padding}" x2="${padding}" y2="${h - padding}" stroke="#94a3b8" stroke-width="2" />
      
      <!-- Area Fill -->
      <path d="${areaD}" fill="url(#areaGrad)" />

      <!-- Line Path -->
      <path d="${pathD}" fill="none" stroke="var(--primary)" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round" />
      
      <!-- Data Points -->
      ${points.map(p => `
        <circle cx="${p.x}" cy="${p.y}" r="5" fill="white" stroke="var(--primary)" stroke-width="3" />
        <text x="${p.x}" y="${p.y - 12}" text-anchor="middle" font-size="10" font-weight="900" fill="var(--primary)">${p.val}%</text>
      `).join('')}
      
      <!-- X Axis Labels (Dates) -->
      ${points.map((p) => `
        <text x="${p.x}" y="${h - 10}" text-anchor="middle" font-size="9" fill="#94a3b8" font-weight="700">${p.date}</text>
      `).join('')}
    </svg>
  `;
}
function teacherReset() {
  if (confirm("🚨 경고: 이 기기의 학생 데이터를 초기화합니까?\n(서버 기록은 보존되나, 이 기기에서의 현재 진행 상황은 초기화됩니다.)")) {
    const pw = prompt("선생님 암호를 입력하세요");
    if (pw === '1234') {
      localStorage.removeItem('sdg_hero_v6');
      alert("데이터가 깨끗하게 지워졌습니다. 처음으로 돌아갑니다.");
      location.reload();
    } else {
      alert("암호가 틀렸습니다.");
    }
  }
}

function openSpreadsheet() {
  const url = localStorage.getItem('sdg_spreadsheet_url');
  if (!url) {
    alert('먼저 위쪽 "스프레드시트 주소" 칸에 URL을 입력하고 저장해주세요.');
    const sheetInput = document.getElementById('spreadsheet-url-input');
    if (sheetInput) sheetInput.focus();
    return;
  }
  window.open(url, '_blank');
}

async function generateSheetDashboard(event) {
  if (APP_STATE.gasUrl === 'PASTE_YOUR_GAS_WEB_APP_URL_HERE' || !APP_STATE.gasUrl.startsWith('https')) {
    alert('먼저 본부 URL을 설정해주세요.');
    return;
  }
  if (!confirm('📊 스프레드시트에 "📊 대시보드" 시트를 생성/갱신합니다.\n기존 대시보드가 있으면 새 데이터로 덮어씁니다.\n진행할까요?')) return;

  const btn = event && event.target ? event.target : null;
  const originalText = btn ? btn.innerText : '';
  if (btn) { btn.disabled = true; btn.innerText = '⏳ 생성 요청 전송 중...'; }

  try {
    await fetch(APP_STATE.gasUrl, {
      method: 'POST',
      mode: 'no-cors',
      body: JSON.stringify({ action: 'generateDashboard' })
    });
    setTimeout(function() {
      if (btn) { btn.disabled = false; btn.innerText = originalText; }
      alert('✅ 대시보드 생성 요청을 보냈습니다.\n10~30초 후 스프레드시트의 "📊 대시보드" 시트를 확인해주세요.\n(차트가 들어가서 시간이 좀 걸려요)');
    }, 1500);
  } catch (err) {
    if (btn) { btn.disabled = false; btn.innerText = originalText; }
    alert('❌ 오류: ' + err.message);
  }
}

async function clearSpreadsheetData() {
  if (APP_STATE.gasUrl === 'PASTE_YOUR_GAS_WEB_APP_URL_HERE') {
    alert("먼저 script.js 상단에 GAS 웹 앱 URL을 붙여넣어 주세요!");
    return;
  }

  if (confirm("🚨 경고: 스프레드시트의 모든 기록 데이터가 영구적으로 삭제됩니다.\n정말로 초기화하시겠습니까?")) {
    const pw = prompt("관리자 암호를 입력하세요");
    if (pw !== '1234') { alert("암호가 틀렸습니다."); return; }

    try {
      await fetch(APP_STATE.gasUrl, {
        method: 'POST',
        mode: 'no-cors', // Use no-cors to avoid CORS issues with GAS POST
        body: JSON.stringify({ action: 'clearData' })
      });
      alert("기록 삭제 요청이 본부에 전송되었습니다. 잠시 후 시트를 확인하세요.");
    } catch (err) {
      alert("초기화 요청 중 오류가 발생했습니다. (잠시 후 다시 시도)");
      console.error(err);
    }
  }
}
function showToast(msg) {
  const t = document.getElementById('toast');
  if (!t) return;
  t.innerText = msg;
  t.style.display = 'block';
  setTimeout(() => { t.style.display = 'none'; }, 3000);
}

function updateDurationDisplay() {
  const s = document.getElementById('start-date').value;
  const e = document.getElementById('end-date').value;
  if (!s || !e) return;

  const start = new Date(s);
  const end = new Date(e);
  const diff = Math.ceil((end - start) / (1000 * 60 * 60 * 24)) + 1;
  const displayVal = diff > 0 ? diff : 0;

  const totalDaysSpan = document.getElementById('setup-total-days');
  if (totalDaysSpan) totalDaysSpan.innerText = displayVal;
}