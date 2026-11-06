// 동적 데이터 배열 (오직 엑셀에서만 데이터를 가져옵니다)
const PROJECTS_DATA = [];
const GRIEVANCE_DATA = [];

const catMap = {
  '기술지원': { id: 1, sub: '기술혁신과 R&D 역량을 키우는 남양주' },
  '자금지원': { id: 2, sub: '기업의 성장을 뒷받침하는 남양주' },
  '판로개척': { id: 3, sub: '국내외 시장으로 뻗어나가는 남양주' },
  '기업환경개선': { id: 4, sub: '기업하기 좋은 도시 남양주' }
};

const nyjColors = ['#002978', '#009fe3', '#e6007e', '#ef7d00', '#00a760', '#b39b70'];
let donutChartInst = null;
let barChartInst = null;

// 보안 접근 플래그
let isGrievanceUnlocked = false;
// 'sos8220'의 SHA-256 해시값 (프론트엔드 코드에 원본 비밀번호 노출 방지)
const TARGET_HASH = "43ca49cf349f794f591fdc429a0801a628dfc7bbf4f5cbcc03c09578136f33d1"; 

function toggleSidebar() {
  const sidebar = document.querySelector('aside');
  const overlay = document.getElementById('overlay');
  if (window.innerWidth <= 768) {
    sidebar.classList.toggle('open');
    overlay.classList.toggle('show');
  } else {
    sidebar.classList.toggle('desktop-collapsed');
  }
}

window.addEventListener('resize', () => {
  if (window.innerWidth > 768) {
    document.querySelector('aside').classList.remove('open');
    document.getElementById('overlay').classList.remove('show');
  }
});

async function fetchExcelData() {
  try {
    const res = await fetch('./data.xlsx?v=' + new Date().getTime());
    if (!res.ok) throw new Error();
    const arrayBuffer = await res.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });

    if (workbook.SheetNames.includes('사업현황')) {
      const bizSheet = workbook.Sheets['사업현황'];
      const bizRows = XLSX.utils.sheet_to_json(bizSheet, { header: 1 });
      if (bizRows.length > 1) {
        PROJECTS_DATA.length = 0;
        bizRows.forEach((cols, idx) => {
          if (idx === 0 || !cols[0]) return;
          const catName = cols[0].toString().trim();
          const catInfo = catMap[catName] || { id: 99, sub: '' };
          const totalBudget = parseInt(cols[3]) || 0;
          
          let parsedPlans = [];
          if (cols[14]) {
              try { parsedPlans = JSON.parse(cols[14]); } 
              catch (e) { parsedPlans = [{ title: '세부 지원내용', budget: (totalBudget > 0 ? (totalBudget / 1000).toLocaleString() + '천원' : '비예산'), desc: cols[10] || '-' }]; }
          } else {
              parsedPlans = [{ title: '세부 지원내용', budget: (totalBudget > 0 ? (totalBudget / 1000).toLocaleString() + '천원' : '비예산'), desc: cols[10] || '-' }];
          }
          
          PROJECTS_DATA.push({
            id: 'biz_' + idx,
            catId: catInfo.id,
            catName: catName,
            catSub: catInfo.sub,
            title: cols[1] || '-',
            shortTitle: cols[2] || '-',
            badge: totalBudget > 0 ? (totalBudget / 1000).toLocaleString() + '천원' : '비예산',
            finance: {
              total: totalBudget,
              city: parseInt(cols[4]) || 0,
              prov: parseInt(cols[5]) || 0,
              nat: parseInt(cols[6]) || 0,
              fund: parseInt(cols[7]) || 0,
              agency: cols[8] || '-',
              matching: cols[9] || '-'
            },
            desc: cols[10] || '-',
            plans: parsedPlans,
            target: cols[11] || '-',
            applyInfo: cols[12] || '-',
            contact: cols[13] || '-',
            docs: {
                official: cols[15] ? cols[15].toString().trim() : '',
                plan: cols[16] ? cols[16].toString().trim() : '',
                status: cols[17] ? cols[17].toString().trim() : ''
            }
          });
        });
      }
    }

    if (workbook.SheetNames.includes('애로사항')) {
      const worksheet = workbook.Sheets['애로사항'];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      GRIEVANCE_DATA.length = 0;
      rows.forEach((cols, idx) => {
        if (idx === 0 || !cols[0]) return;
        GRIEVANCE_DATA.push({
          corpName: cols[1] || cols[0] || '-',
          location: cols[2] || '-',
          region: cols[3] || '기타',
          field: cols[4] || '-',
          grievanceContent: cols[5] || '-',
          consultContent: cols[6] || '-',
          processContent: cols[7] || '-',
          rawStatus: (cols[8] || '검토').toString().trim()
        });
      });
    }
  } catch (e) {
    console.error("데이터 로드 실패:", e);
  }

  updateDashboardStats();
  renderSidebarMenu();
  renderDashboardCards();
  initCharts();
  updateRegionSelect();
  renderGrievanceTables(GRIEVANCE_DATA);
}

function updateDashboardStats() {
  let total = 0, cityFund = 0, prov = 0;
  let catBudgets = { 1: 0, 2: 0, 3: 0, 4: 0 };

  PROJECTS_DATA.forEach(p => {
    total += p.finance.total;
    cityFund += (p.finance.city + p.finance.fund);
    prov += p.finance.prov;
    if (catBudgets[p.catId] !== undefined) catBudgets[p.catId] += p.finance.total;
  });

  const toUk = (v) => (v / 100000).toFixed(2) + '억원';
  document.getElementById('banner-total-budget').innerText = toUk(total);
  document.getElementById('banner-total-count').innerText = `2026년 남양주시 4대 분야 ${PROJECTS_DATA.length}개 대표 시책 예산액 총 ${total.toLocaleString()}천원 편성 현황`;
  document.getElementById('banner-city-budget').innerText = toUk(cityFund);
  document.getElementById('banner-prov-budget').innerText = toUk(prov);

  document.getElementById('sidebar-total-budget').innerText = toUk(total);
  document.getElementById('sidebar-cat1-budget').innerText = toUk(catBudgets[1]);
  document.getElementById('sidebar-cat2-budget').innerText = toUk(catBudgets[2]);
  document.getElementById('sidebar-cat3-budget').innerText = toUk(catBudgets[3]);
  document.getElementById('sidebar-cat4-budget').innerText = toUk(catBudgets[4]);
}

function renderSidebarMenu() {
  const makeNavItem = (p, badgeBgColor) => `
    <li class="nav-item sub-item" style="padding: 7px 12px 7px 24px; font-size: 0.82rem; position: relative; display: flex; justify-content: space-between; align-items: center;" id="nav-${p.id}" onclick="openProjectDetail('${p.id}')">
      <div style="overflow:hidden; display:flex; align-items:center;">
        <span style="position: absolute; left: 8px; top: 7px; color: #64748b;">└</span>
        <span style="overflow:hidden; text-overflow:ellipsis; white-space:nowrap; padding-left:4px;" title="${p.title}">${p.shortTitle}</span>
      </div>
      <span style="font-size: 0.65rem; background: ${badgeBgColor}; color: #ffffff; padding: 2px 6px; border-radius: 10px; margin-left: 6px; white-space: nowrap; font-weight: 700;">${p.badge}</span>
    </li>
  `;

  document.getElementById('dynamic-nav-cat1').innerHTML = PROJECTS_DATA.filter(p => p.catId === 1).map(p => makeNavItem(p, '#2563eb')).join('');
  document.getElementById('dynamic-nav-cat2').innerHTML = PROJECTS_DATA.filter(p => p.catId === 2).map(p => makeNavItem(p, '#7c3aed')).join('');
  document.getElementById('dynamic-nav-cat3').innerHTML = PROJECTS_DATA.filter(p => p.catId === 3).map(p => makeNavItem(p, '#059669')).join('');
  document.getElementById('dynamic-nav-cat4').innerHTML = PROJECTS_DATA.filter(p => p.catId === 4).map(p => makeNavItem(p, '#d97706')).join('');
}

function renderDashboardCards() {
  const grid = document.getElementById('searchableCardsGrid');
  grid.innerHTML = PROJECTS_DATA.map(p => {
    let badgeColor = p.catId === 2 ? 'badge-purple' : p.catId === 3 ? 'badge-green' : p.catId === 4 ? 'badge-amber' : 'badge-blue';
    return `
      <div class="content-box searchable-card" data-category="${p.catName}">
        <div class="box-header">
          <span class="box-title" style="color:var(--primary);">${p.title}</span>
          <span class="box-badge ${badgeColor}">${p.badge}</span>
        </div>
        <div class="box-desc">• ${p.desc}</div>
        <div style="font-size:0.78rem; color:#64748b; margin-bottom:10px;">
          <b>수행:</b> ${p.finance.agency} | <b>분야:</b> ${p.catName}
        </div>
        <button class="tab-btn" style="background:#f0f7ff; color:var(--primary); font-weight:700; border-radius:6px; padding:6px; width:100%; border:1px solid var(--primary-light);" onclick="openProjectDetail('${p.id}')">
          상세내용 및 핵심문서 관리 &rarr;
        </button>
      </div>
    `;
  }).join('');
  document.getElementById('searchResultCount').innerText = `총 ${PROJECTS_DATA.length}개 사업 표시 중`;
}

// SHA-256 해시 함수 (비밀번호 암호화 비교용)
async function hashPassword(message) {
  const msgBuffer = new TextEncoder('utf-8').encode(message);
  const hashBuffer = await crypto.subtle.digest('SHA-256', msgBuffer);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

// 비밀번호 검증 로직
async function verifyPassword() {
  const inputEl = document.getElementById("auth-password");
  const inputHash = await hashPassword(inputEl.value);
  
  if (inputHash === TARGET_HASH) {
    isGrievanceUnlocked = true;
    inputEl.value = ""; // 입력창 초기화
    
    // 잠금 아이콘 풀림 상태로 변경
    const navBadge = document.querySelector("#nav-grievance .nav-badge");
    if(navBadge) {
        navBadge.innerText = "현황조회";
        navBadge.style.background = "#059669"; // 초록색으로 변경
    }
    
    switchView('grievance'); // 실제 화면으로 이동
  } else {
    alert("비밀번호가 일치하지 않습니다.");
    inputEl.value = "";
    inputEl.focus();
  }
}

function switchView(viewId) {
  // 기업애로 관리 탭 클릭 시, 잠금 해제 상태가 아니면 비밀번호 화면으로 강제 이동
  if (viewId === 'grievance' && !isGrievanceUnlocked) {
    viewId = 'password';
  }

  document.querySelectorAll('.view-panel').forEach(el => el.classList.add('hidden'));
  document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
  const target = document.getElementById(`view-${viewId}`);
  if (target) target.classList.remove('hidden');

  const meta = {
    'dashboard': { title: '2026년 남양주시 기업 지원시책 종합 대시보드', desc: '기술지원·자금지원·판로개척·기업환경개선 4대 분야 총 26개 대표 기업지원사업 및 예산액 현황 분석' },
    'password':  { title: '관리자 권한 인증', desc: '민감한 기업애로사항 정보를 보호하기 위한 보안 페이지입니다.' },
    'grievance': { title: '💬 기업애로사항 접수·관리', desc: '관내 기업체 현장 애로사항 접수 목록 및 상태·지역별 통계를 조회합니다.' },
    'manual':    { title: '📖 업무매뉴얼', desc: '맞춤형 기업애로 해결사업(e호조) 및 기업환경개선사업(보탬e) 업무 지침 가이드' }
  };

  if (meta[viewId]) {
    document.getElementById('view-title').innerText = meta[viewId].title;
    document.getElementById('view-desc').innerText = meta[viewId].desc;
  }
  
  if (viewId === 'dashboard') {
    document.querySelector('.nav-list .nav-item:first-of-type').classList.add('active');
  } else if (viewId === 'password' || viewId === 'grievance') {
    document.getElementById('nav-grievance').classList.add('active');
  }
}

function openProjectDetail(bizId) {
  const proj = PROJECTS_DATA.find(p => p.id === bizId);
  if (!proj) return;

  switchView('project-detail');
  document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
  const navEl = document.getElementById(`nav-${bizId}`);
  if (navEl) navEl.classList.add('active');

  document.getElementById('view-title').innerText = proj.title;
  document.getElementById('view-desc').innerText = `${proj.catName} - ${proj.catSub} | ${proj.desc}`;
  document.getElementById('detail-proj-desc').innerText = proj.desc;
  document.getElementById('detail-proj-agency').innerText = proj.finance.agency;

  const fmt = v => v === 0 ? '-' : v.toLocaleString() + '천원';
  document.getElementById('fin-total').innerText = fmt(proj.finance.total);
  document.getElementById('fin-city').innerHTML = proj.finance.fund > 0 ? `${fmt(proj.finance.city)} <span style="font-size:0.75rem; color:#d97706;">(기금: ${fmt(proj.finance.fund)})</span>` : fmt(proj.finance.city);
  document.getElementById('fin-prov').innerText = fmt(proj.finance.prov);
  document.getElementById('fin-nat').innerText = fmt(proj.finance.nat);
  document.getElementById('fin-matching').innerText = proj.finance.matching;
  document.getElementById('fin-exec').innerText = proj.finance.agency;

  document.getElementById('detail-tab-plan').innerHTML = proj.plans.map(pl => `
    <div class="content-box">
      <div class="box-header">
        <span class="box-title" style="color:var(--primary);">${pl.title}</span>
        <span class="box-badge badge-green">${pl.budget}</span>
      </div>
      <div class="box-desc" style="white-space: pre-wrap;">• ${pl.desc}</div>
    </div>
  `).join('');

  document.getElementById('detail-history-custom').innerHTML = `
    <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:20px; line-height:1.7;">
      <h4 style="font-size:1rem; color:var(--primary); margin-bottom:12px; font-weight:700;">📌 신청방법 및 문의처 세부안내</h4>
      <p style="font-size:0.86rem; color:#334155; margin-bottom:8px;"><b>지원 대상:</b> ${proj.target}</p>
      <p style="font-size:0.86rem; color:#334155; margin-bottom:8px;"><b>신청 방법 및 기간:</b> ${proj.applyInfo}</p>
      <p style="font-size:0.86rem; color:#334155; margin-bottom:8px;"><b>문의처 및 담당부서:</b> ${proj.contact}</p>
    </div>
  `;

  renderAttachmentRows(proj);
  switchDetailTab('plan');
}

function switchDetailTab(type) {
  const isPlan = (type === 'plan');
  document.getElementById('tab-btn-plan').classList.toggle('active', isPlan);
  document.getElementById('tab-btn-history').classList.toggle('active', !isPlan);
  document.getElementById('detail-tab-plan').classList.toggle('hidden', !isPlan);
  document.getElementById('detail-tab-history').classList.toggle('hidden', isPlan);
}

const FILE_CATS = [
  { key: 'official', name: '사업관련 공문', badge: 'badge-blue' },
  { key: 'plan', name: '추진계획 및 사업계획서', badge: 'badge-green' },
  { key: 'status', name: '추진현황 / 접수내역 및 정산서', badge: 'badge-amber' }
];

function renderAttachmentRows(proj) {
  document.getElementById('detail-attachments-container').innerHTML = FILE_CATS.map(cat => {
    const filename = proj.docs[cat.key];
    const isRegistered = Boolean(filename && filename.length > 0);
    const safeFilename = encodeURIComponent(filename);

    return `
      <div class="file-manage-row">
        <div class="file-info-col">
          <div style="display:flex; align-items:center; gap:10px;">
            <span class="box-badge ${cat.badge}" style="padding:4px 10px; border-radius:6px;">${cat.name}</span>
            <span style="font-weight:700; font-size:0.92rem; color:var(--text);">${cat.name} 관리</span>
          </div>
          <div style="font-size:0.83rem; margin-top:2px;">
            ${isRegistered 
              ? `<span style="display:inline-flex; align-items:center; gap:4px; color:#059669; font-weight:700; background:#dcfce7; padding:2px 8px; border-radius:6px; font-size:0.78rem;">✓ 엑셀 연동됨</span> <span style="color:#334155; font-weight:600; margin-left:4px;">${filename}</span>` 
              : `<span style="color:#94a3b8; font-style:italic;">미등록 상태 (data.xlsx 엑셀에 파일명을 입력하세요)</span>`}
          </div>
        </div>
        <div class="file-actions-col">
          <button class="btn-download-file ${isRegistered ? '' : 'disabled'}" onclick="window.open('./${safeFilename}', '_blank')">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
            서버 다운로드
          </button>
        </div>
      </div>
    `;
  }).join('');
}

function filterDashboard() {
  const q = document.getElementById('dashboardSearch').value.toLowerCase().trim();
  const cat = document.getElementById('categoryFilter').value;
  let count = 0;
  document.querySelectorAll('.searchable-card').forEach((card, idx) => {
    const item = PROJECTS_DATA[idx];
    const match = (!q || card.innerText.toLowerCase().includes(q)) && (cat === 'all' || item.catName === cat);
    card.style.display = match ? 'flex' : 'none';
    if (match) count++;
  });
  document.getElementById('searchResultCount').innerText = `총 ${count}개 사업 표시 중`;
}

function updateRegionSelect() {
  const select = document.getElementById('regionFilter');
  const regions = [...new Set(GRIEVANCE_DATA.map(item => item.region || '기타'))].sort();
  select.innerHTML = '<option value="all">전체 지역 보기</option>' + regions.map(r => `<option value="${r}">${r}</option>`).join('');
}

function renderGrievanceTables(dataList) {
  const tbody = document.getElementById('grievance-tbody');
  const summaryBox = document.getElementById('grievance-summary-content');
  const regionBox = document.getElementById('region-summary-content');
  const selectedRegion = document.getElementById('regionFilter').value;
  tbody.innerHTML = '';

  if (!dataList || dataList.length === 0) {
    tbody.innerHTML = `<tr><td colspan="9" style="text-align:center; padding:30px; color:#94a3b8;">데이터가 없습니다.</td></tr>`;
    return;
  }

  let statusCounts = { '대기': 0, '진행 중': 0, '검토': 0, '장기 검토': 0, '불가': 0, '완료': 0 };
  let regionCounts = {};

  dataList.forEach(item => {
    let stKey = item.rawStatus.includes('장기') ? '장기 검토' : item.rawStatus;
    if (statusCounts[stKey] !== undefined) statusCounts[stKey]++;
    let regKey = item.region || '기타';
    regionCounts[regKey] = (regionCounts[regKey] || 0) + 1;
  });

  regionBox.innerHTML = Object.entries(regionCounts).map(([reg, count]) => `
    <div class="content-box" style="text-align:center; padding:12px;"><div style="font-size:0.8rem; font-weight:700; color:var(--text-muted);">${reg} 지역</div><div style="font-size:1.3rem; color:var(--primary-light); font-weight:800;">${count}건</div></div>
  `).join('');

  summaryBox.innerHTML = Object.entries(statusCounts).map(([status, count]) => `
    <div class="content-box" style="text-align:center; padding:12px;"><div style="font-size:0.8rem; font-weight:700; color:var(--text-muted);">${status}</div><div style="font-size:1.3rem; color:var(--primary); font-weight:800;">${count}건</div></div>
  `).join('');

  let filteredData = dataList.filter(item => selectedRegion === 'all' || item.region === selectedRegion);

  filteredData.forEach((item, idx) => {
    let badgeStyle = 'background: #fef9c3; color: #a16207;';
    if (item.rawStatus.includes('진행 중')) badgeStyle = 'background: #dcfce7; color: #15803d;';
    else if (item.rawStatus.includes('불가')) badgeStyle = 'background: #fee2e2; color: #dc2626;';
    else if (item.rawStatus.includes('완료')) badgeStyle = 'background: #dbeafe; color: #1d4ed8;';

    tbody.innerHTML += `
      <tr>
        <td style="text-align:center; font-weight:700; color:var(--primary);">${idx + 1}</td>
        <td><b>${item.corpName}</b></td><td>${item.location}</td>
        <td><span class="box-badge badge-blue">${item.region}</span></td>
        <td>${item.field}</td><td>${item.grievanceContent}</td>
        <td>${item.consultContent}</td><td>${item.processContent}</td>
        <td style="text-align:center;"><span class="box-badge" style="${badgeStyle}">${item.rawStatus}</span></td>
      </tr>
    `;
  });
  document.getElementById('grievance-count').innerText = `총 접수: ${dataList.length}건`;
}

function filterGrievanceTable() { renderGrievanceTables(GRIEVANCE_DATA); }

function handleExcelUpload(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const workbook = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      let sheetName = workbook.SheetNames.includes('애로사항') ? '애로사항' : workbook.SheetNames[0];
      GRIEVANCE_DATA.length = 0;
      XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 }).forEach((cols, idx) => {
        if (idx === 0 || !cols[0]) return;
        GRIEVANCE_DATA.push({
          corpName: cols[1] || cols[0] || '-',
          location: cols[2] || '-',
          region: cols[3] || '기타',
          field: cols[4] || '-',
          grievanceContent: cols[5] || '-',
          consultContent: cols[6] || '-',
          processContent: cols[7] || '-',
          rawStatus: (cols[8] || '검토').toString().trim()
        });
      });
      updateRegionSelect();
      renderGrievanceTables(GRIEVANCE_DATA);
      alert('엑셀 파일 갱신 완료');
    } catch (err) { alert('오류 발생'); }
  };
  reader.readAsArrayBuffer(file);
  event.target.value = '';
}

function clearGrievanceData() {
  fetchExcelData();
  alert('깃허브 데이터로 다시 연동되었습니다.');
}

function showManual(type) {
  const box = document.getElementById('manual-detail-box');
  box.classList.remove('hidden');
  const isEhojo = (type === 'ehojo');
  document.getElementById('manual-detail-title').innerText = isEhojo ? '💰 e호조 회계지출 프로세스 매뉴얼' : '📦 보탬e 보조금 교부정산 매뉴얼';
  document.getElementById('manual-detail-desc').innerHTML = isEhojo 
    ? '세부사업계획 등록 ➔ 예산과목 배정 ➔ 지출품의 ➔ 지출원인행위 확정 ➔ 지급명령 및 계좌이체' 
    : '보조사업 공모 및 신청 접수 ➔ 심의위원회 선정평가 ➔ 보조금 교부결정 통보 ➔ 집행 모니터링 ➔ 실적보고서 검사 및 정산반납 확정';
}

function initCharts() {
  if (donutChartInst) donutChartInst.destroy();
  if (barChartInst) barChartInst.destroy();

  let catBudgets = { 1: 0, 2: 0, 3: 0, 4: 0 };
  let catCity = { 1: 0, 2: 0, 3: 0, 4: 0 };
  let catProv = { 1: 0, 2: 0, 3: 0, 4: 0 };
  let catNatFund = { 1: 0, 2: 0, 3: 0, 4: 0 };

  PROJECTS_DATA.forEach(p => {
    catBudgets[p.catId] += p.finance.total;
    catCity[p.catId] += p.finance.city;
    catProv[p.catId] += p.finance.prov;
    catNatFund[p.catId] += (p.finance.nat + p.finance.fund);
  });

  const totalOverall = Object.values(catBudgets).reduce((a, b) => a + b, 0);
  const donutDesc = document.getElementById('chart-donut-desc');
  if (donutDesc) donutDesc.innerText = `총 ${(totalOverall / 1000).toLocaleString()}천원 기준`;

  const donutCtx = document.getElementById('categoryDonutChart').getContext('2d');
  donutChartInst = new Chart(donutCtx, {
    type: 'doughnut',
    data: {
      labels: ['1. 기술지원', '2. 자금지원', '3. 판로개척', '4. 기업환경개선'],
      datasets: [{
        data: [catBudgets[1], catBudgets[2], catBudgets[3], catBudgets[4]],
        backgroundColor: [nyjColors[1], nyjColors[2], nyjColors[4], nyjColors[3]],
        borderWidth: 2,
        borderColor: '#ffffff',
        hoverOffset: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 10 } } },
        tooltip: {
          callbacks: {
            label: function(ctx) {
              const val = ctx.raw;
              const pct = ((val / totalOverall) * 100).toFixed(1);
              return ` ${ctx.label}: ${val.toLocaleString()}천원 (${pct}%)`;
            }
          }
        }
      },
      cutout: '66%'
    }
  });

  const barCtx = document.getElementById('categoryBarChart').getContext('2d');
  barChartInst = new Chart(barCtx, {
    type: 'bar',
    data: {
      labels: ['기술지원', '자금지원', '판로개척', '기업환경개선'],
      datasets: [
        { label: '시비', data: [catCity[1] / 1000, catCity[2] / 1000, catCity[3] / 1000, catCity[4] / 1000], backgroundColor: nyjColors[0], borderRadius: 4 },
        { label: '도비', data: [catProv[1] / 1000, catProv[2] / 1000, catProv[3] / 1000, catProv[4] / 1000], backgroundColor: nyjColors[4], borderRadius: 4 },
        { label: '국비/기금', data: [catNatFund[1] / 1000, catNatFund[2] / 1000, catNatFund[3] / 1000, catNatFund[4] / 1000], backgroundColor: nyjColors[3], borderRadius: 4 }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { stacked: true },
        y: { stacked: true, ticks: { callback: v => v + '백만' } }
      },
      plugins: {
        legend: { position: 'top', labels: { boxWidth: 10, font: { size: 11 } } }
      }
    }
  });
}

window.addEventListener('DOMContentLoaded', fetchExcelData);
