// 동적 데이터 배열 (오직 엑셀에서만 데이터를 가져옵니다)
const PROJECTS_DATA = [];
const GRIEVANCE_DATA = [];

const catMap = {
  '기술지원': { id: 1, sub: '기술혁신과 R&D 역량을 키우는 남양주' },
  '자금지원': { id: 2, sub: '기업의 성장을 뒷받침하는 남양주' },
  '판로개척': { id: 3, sub: '국내외 시장으로 뻗어나가는 남양주' },
  '기업환경개선': { id: 4, sub: '기업하기 좋은 도시 남양주' }
};

let donutChartInst = null;
let barChartInst = null;

// 보안 접근 플래그
let isGrievanceUnlocked = false;
// 'sos8220'의 SHA-256 해시값 
const TARGET_HASH = "43ca49cf349f794f591fdc429a0801a628dfc7bbf4f5cbcc03c09578136f33d1"; 

function toggleSidebar() {
  const sidebar = document.getElementById('sidebar');
  const overlay = document.getElementById('overlay');
  if (window.innerWidth <= 768) {
    sidebar.classList.toggle('-translate-x-full');
    overlay.classList.toggle('hidden');
  } else {
    // Desktop collapse logic if needed, currently hidden on desktop
  }
}

window.addEventListener('resize', () => {
  if (window.innerWidth > 768) {
    document.getElementById('sidebar').classList.remove('-translate-x-full');
    document.getElementById('overlay').classList.add('hidden');
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

function animateNumber(elementId, targetVal, isBillion) {
    const el = document.getElementById(elementId);
    if(!el) return;
    let current = 0;
    const increment = targetVal / 40; 
    const timer = setInterval(() => {
        current += increment;
        if (current >= targetVal) {
            clearInterval(timer);
            current = targetVal;
        }
        if(isBillion) {
            el.innerHTML = current.toFixed(2) + '<span class="text-sm font-bold ml-1">억</span>';
        } else {
            el.innerHTML = current.toFixed(2);
        }
    }, 20);
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

  document.getElementById('banner-total-count').innerText = `2026년 남양주시 4대 분야 ${PROJECTS_DATA.length}개 대표 시책 예산액 현황`;
  
  animateNumber('banner-total-budget', total / 100000, true);
  animateNumber('banner-city-budget', cityFund / 100000, true);
  
  document.getElementById('sidebar-total-budget').innerText = (total / 100000).toFixed(2) + '억';
  document.getElementById('sidebar-cat1-budget').innerText = (catBudgets[1] / 100000).toFixed(2) + '억';
  document.getElementById('sidebar-cat2-budget').innerText = (catBudgets[2] / 100000).toFixed(2) + '억';
  document.getElementById('sidebar-cat3-budget').innerText = (catBudgets[3] / 100000).toFixed(2) + '억';
  document.getElementById('sidebar-cat4-budget').innerText = (catBudgets[4] / 100000).toFixed(2) + '억';
}

function renderSidebarMenu() {
  const makeNavItem = (p, badgeClass) => `
    <div class="nav-item-hover px-2 py-2 rounded-lg flex justify-between items-center cursor-pointer transition-all border-l-2 border-transparent hover:border-gray-400 gap-2" id="nav-${p.id}" onclick="openProjectDetail('${p.id}')">
      <span class="text-xs text-gray-800 font-bold truncate flex-1" title="${p.title}">· ${p.shortTitle}</span>
      <span class="text-[0.65rem] font-bold px-1.5 py-0.5 rounded-md whitespace-nowrap shadow-sm ${badgeClass}">${p.badge}</span>
    </div>
  `;

  document.getElementById('dynamic-nav-cat1').innerHTML = PROJECTS_DATA.filter(p => p.catId === 1).map(p => makeNavItem(p, 'bg-blue-100 text-blue-700')).join('');
  document.getElementById('dynamic-nav-cat2').innerHTML = PROJECTS_DATA.filter(p => p.catId === 2).map(p => makeNavItem(p, 'bg-purple-100 text-purple-700')).join('');
  document.getElementById('dynamic-nav-cat3').innerHTML = PROJECTS_DATA.filter(p => p.catId === 3).map(p => makeNavItem(p, 'bg-green-100 text-green-700')).join('');
  document.getElementById('dynamic-nav-cat4').innerHTML = PROJECTS_DATA.filter(p => p.catId === 4).map(p => makeNavItem(p, 'bg-orange-100 text-orange-700')).join('');
}

function renderDashboardCards() {
  const grid = document.getElementById('searchableCardsGrid');
  grid.innerHTML = PROJECTS_DATA.map(p => {
    let borderColor = p.catId === 2 ? 'border-t-purple-500' : p.catId === 3 ? 'border-t-green-500' : p.catId === 4 ? 'border-t-orange-500' : 'border-t-blue-600';
    let badgeBg = p.catId === 2 ? 'bg-purple-100 text-purple-700' : p.catId === 3 ? 'bg-green-100 text-green-700' : p.catId === 4 ? 'bg-orange-100 text-orange-700' : 'bg-blue-100 text-blue-700';

    return `
      <div class="glass-card rounded-2xl p-6 flex flex-col justify-between group border-t-4 ${borderColor} searchable-card" data-category="${p.catName}">
        <div>
          <div class="flex justify-between items-start mb-4 gap-2">
            <h3 class="text-base md:text-lg font-bold text-gray-900 leading-tight group-hover:text-nyjBlue transition-colors">${p.title}</h3>
            <span class="text-xs font-bold px-2 py-1 rounded-md ${badgeBg} whitespace-nowrap shadow-sm">${p.badge}</span>
          </div>
          <p class="text-sm text-gray-800 font-medium mb-6 leading-relaxed line-clamp-3">${p.desc}</p>
        </div>
        <div>
          <div class="text-xs text-gray-700 font-medium mb-4 bg-white/70 p-2 rounded-lg border border-white">
            <span class="font-bold text-gray-900">수행:</span> ${p.finance.agency}
          </div>
          <button class="w-full py-2.5 bg-gradient-to-r from-blue-50 to-indigo-50 hover:from-blue-100 hover:to-indigo-100 text-nyjBlue font-bold rounded-xl text-sm transition-all border border-blue-100 shadow-sm" onclick="openProjectDetail('${p.id}')">
            상세내용 및 문서 관리 &rarr;
          </button>
        </div>
      </div>
    `;
  }).join('');
  document.getElementById('searchResultCount').innerText = `총 ${PROJECTS_DATA.length}개`;
}

// SHA-256 해시 함수
async function hashPassword(message) {
  const msgBuffer = new TextEncoder('utf-8').encode(message);
  const hashBuffer = await crypto.subtle.digest('SHA-256', msgBuffer);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

// 비밀번호 검증
async function verifyPassword() {
  const inputEl = document.getElementById("auth-password");
  const inputHash = await hashPassword(inputEl.value);
  
  if (inputHash === TARGET_HASH) {
    isGrievanceUnlocked = true;
    inputEl.value = ""; 
    
    const badge = document.getElementById("grievance-badge");
    if(badge) {
        badge.innerHTML = `<svg class="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 11V7a4 4 0 118 0m-4 8v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2z"></path></svg> 조회`;
        badge.className = "text-[0.7rem] bg-green-100 text-green-700 px-2 py-0.5 rounded-full border border-green-200 shadow-sm font-bold flex items-center gap-1";
    }
    document.getElementById("nav-grievance").querySelector('svg').classList.replace('text-red-400', 'text-green-500');
    
    switchView('grievance'); 
  } else {
    alert("비밀번호가 일치하지 않습니다.");
    inputEl.value = "";
    inputEl.focus();
  }
}

function switchView(viewId) {
  if (viewId === 'grievance' && !isGrievanceUnlocked) viewId = 'password';

  document.querySelectorAll('.view-panel').forEach(el => {
      el.classList.add('hidden');
      el.classList.remove('active');
  });
  
  document.querySelectorAll('.nav-item-active').forEach(el => {
      if(el.id === 'nav-dashboard') {
          el.classList.remove('nav-item-active');
          el.classList.add('text-gray-800');
      }
  });

  const target = document.getElementById(`view-${viewId}`);
  if (target) {
      target.classList.remove('hidden');
      target.classList.add('active');
  }

  const meta = {
    'dashboard': { title: '2026 기업 지원시책 대시보드', desc: '기술지원·자금지원·판로개척·기업환경개선 종합 현황 분석' },
    'password':  { title: '관리자 권한 인증', desc: '민감한 정보를 보호하기 위한 보안 페이지입니다.' },
    'grievance': { title: '기업애로사항 접수·관리', desc: '관내 기업체 현장 애로사항 목록 및 통계를 조회합니다.' },
    'project-detail': { title: '사업 상세정보', desc: '해당 시책의 세부 추진계획 및 관련 문서를 관리합니다.' }
  };

  if (meta[viewId]) {
    document.getElementById('view-title').innerText = meta[viewId].title;
    document.getElementById('view-desc').innerText = meta[viewId].desc;
  }
  
  if (viewId === 'dashboard') {
    const dNav = document.getElementById('nav-dashboard');
    dNav.classList.add('nav-item-active');
    dNav.classList.remove('text-gray-800');
  }
}

function openProjectDetail(bizId) {
  const proj = PROJECTS_DATA.find(p => p.id === bizId);
  if (!proj) return;

  switchView('project-detail');

  document.getElementById('view-title').innerText = proj.title;
  document.getElementById('view-desc').innerText = `${proj.catName} - ${proj.catSub} | ${proj.desc}`;
  document.getElementById('detail-proj-desc').innerText = proj.desc;
  document.getElementById('detail-proj-agency').innerText = proj.finance.agency;

  const fmt = v => v === 0 ? '-' : v.toLocaleString() + '천원';
  document.getElementById('fin-total').innerText = fmt(proj.finance.total);
  document.getElementById('fin-city').innerHTML = proj.finance.fund > 0 ? `${fmt(proj.finance.city)} <br><span class="text-[0.65rem] text-orange-600 font-bold">(기금: ${fmt(proj.finance.fund)})</span>` : fmt(proj.finance.city);
  document.getElementById('fin-prov').innerText = fmt(proj.finance.prov);
  document.getElementById('fin-nat').innerText = fmt(proj.finance.nat);
  document.getElementById('fin-matching').innerText = proj.finance.matching;
  document.getElementById('fin-exec').innerText = proj.finance.agency;

  document.getElementById('detail-tab-plan').innerHTML = proj.plans.map(pl => `
    <div class="bg-white/90 p-5 rounded-2xl border border-gray-200 shadow-sm">
      <div class="flex justify-between items-start mb-3 gap-2">
        <span class="font-bold text-nyjBlue text-sm">${pl.title}</span>
        <span class="text-[0.65rem] font-bold bg-green-100 text-green-700 px-2 py-1 rounded-md whitespace-nowrap">${pl.budget}</span>
      </div>
      <div class="text-xs text-gray-800 font-medium leading-relaxed whitespace-pre-wrap">${pl.desc}</div>
    </div>
  `).join('');

  document.getElementById('detail-history-custom').innerHTML = `
    <h4 class="text-sm font-bold text-nyjBlue mb-3 border-l-4 border-nyjSkyBlue pl-2">신청방법 및 문의처 세부안내</h4>
    <div class="space-y-2 text-xs text-gray-800 font-medium">
        <p><b class="text-gray-900">지원 대상:</b> ${proj.target}</p>
        <p><b class="text-gray-900">신청 방법:</b> ${proj.applyInfo}</p>
        <p><b class="text-gray-900">문의 부서:</b> ${proj.contact}</p>
    </div>
  `;

  renderAttachmentRows(proj);
  switchDetailTab('plan');
}

function switchDetailTab(type) {
  const isPlan = (type === 'plan');
  
  const btnPlan = document.getElementById('tab-btn-plan');
  const btnHist = document.getElementById('tab-btn-history');
  
  if(isPlan) {
      btnPlan.className = "flex-1 py-4 text-sm font-bold text-nyjBlue border-b-2 border-nyjBlue bg-white transition-colors";
      btnHist.className = "flex-1 py-4 text-sm font-bold text-gray-600 border-b-2 border-transparent hover:bg-gray-50 transition-colors";
      document.getElementById('detail-tab-plan').classList.remove('hidden');
      document.getElementById('detail-tab-history').classList.add('hidden');
  } else {
      btnHist.className = "flex-1 py-4 text-sm font-bold text-nyjBlue border-b-2 border-nyjBlue bg-white transition-colors";
      btnPlan.className = "flex-1 py-4 text-sm font-bold text-gray-600 border-b-2 border-transparent hover:bg-gray-50 transition-colors";
      document.getElementById('detail-tab-history').classList.remove('hidden');
      document.getElementById('detail-tab-plan').classList.add('hidden');
  }
}

const FILE_CATS = [
  { key: 'official', name: '사업관련 공문', badge: 'bg-blue-100 text-blue-700 border-blue-200' },
  { key: 'plan', name: '추진계획/계획서', badge: 'bg-green-100 text-green-700 border-green-200' },
  { key: 'status', name: '추진현황/정산서', badge: 'bg-orange-100 text-orange-700 border-orange-200' }
];

function renderAttachmentRows(proj) {
  document.getElementById('detail-attachments-container').innerHTML = FILE_CATS.map(cat => {
    const filename = proj.docs[cat.key];
    const isRegistered = Boolean(filename && filename.length > 0);
    const safeFilename = encodeURIComponent(filename);

    return `
      <div class="flex flex-col md:flex-row justify-between items-start md:items-center p-4 bg-white/70 rounded-xl border border-gray-100 shadow-sm gap-4">
        <div>
          <div class="flex items-center gap-2 mb-1">
            <span class="text-[0.7rem] font-bold border px-2 py-0.5 rounded-md ${cat.badge}">${cat.name}</span>
            <span class="text-sm font-bold text-gray-900">문서 연동</span>
          </div>
          <div class="text-xs font-medium text-gray-800">
            ${isRegistered 
              ? `<span class="text-green-700 font-bold mr-1">✓ 등록됨:</span> ${filename}` 
              : `<span class="text-gray-500 italic">미등록 상태 (data.xlsx에 파일명 입력 필요)</span>`}
          </div>
        </div>
        <button class="w-full md:w-auto px-4 py-2 text-xs font-bold rounded-lg transition-colors flex items-center justify-center gap-1 ${isRegistered ? 'bg-nyjBlue text-white hover:bg-blue-800' : 'bg-gray-300 text-gray-500 cursor-not-allowed'}" ${isRegistered ? `onclick="window.open('./${safeFilename}', '_blank')"` : 'disabled'}>
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
          다운로드
        </button>
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
  document.getElementById('searchResultCount').innerText = `총 ${count}개`;
}

function updateRegionSelect() {
  const select = document.getElementById('regionFilter');
  const regions = [...new Set(GRIEVANCE_DATA.map(item => item.region || '기타'))].sort();
  select.innerHTML = '<option value="all">전체 지역</option>' + regions.map(r => `<option value="${r}">${r}</option>`).join('');
}

function renderGrievanceTables(dataList) {
  const tbody = document.getElementById('grievance-tbody');
  const summaryBox = document.getElementById('grievance-summary-content');
  const regionBox = document.getElementById('region-summary-content');
  const selectedRegion = document.getElementById('regionFilter').value;
  tbody.innerHTML = '';

  if (!dataList || dataList.length === 0) {
    tbody.innerHTML = `<tr><td colspan="9" class="text-center p-8 text-gray-500 font-bold">데이터가 없습니다.</td></tr>`;
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
    <div class="bg-white/80 border border-gray-100 p-3 rounded-xl text-center shadow-sm">
      <div class="text-[0.7rem] font-bold text-gray-700 mb-1">${reg}</div>
      <div class="text-lg font-black text-nyjGreen">${count}건</div>
    </div>
  `).join('');

  summaryBox.innerHTML = Object.entries(statusCounts).map(([status, count]) => `
    <div class="bg-white/80 border border-gray-100 p-3 rounded-xl text-center shadow-sm">
      <div class="text-[0.7rem] font-bold text-gray-700 mb-1">${status}</div>
      <div class="text-lg font-black text-nyjBlue">${count}건</div>
    </div>
  `).join('');

  let filteredData = dataList.filter(item => selectedRegion === 'all' || item.region === selectedRegion);

  filteredData.forEach((item, idx) => {
    let badgeStyle = 'bg-yellow-100 text-yellow-800 border-yellow-200';
    if (item.rawStatus.includes('진행 중')) badgeStyle = 'bg-green-100 text-green-800 border-green-200';
    else if (item.rawStatus.includes('불가')) badgeStyle = 'bg-red-100 text-red-800 border-red-200';
    else if (item.rawStatus.includes('완료')) badgeStyle = 'bg-blue-100 text-blue-800 border-blue-200';

    tbody.innerHTML += `
      <tr class="hover:bg-gray-100/70 transition-colors">
        <td class="p-3 text-center font-bold text-gray-600">${idx + 1}</td>
        <td class="p-3 font-bold text-gray-900">${item.corpName}</td>
        <td class="p-3 font-medium text-gray-800">${item.location}</td>
        <td class="p-3"><span class="text-[0.65rem] px-2 py-1 bg-gray-200 rounded-md font-bold text-gray-800 whitespace-nowrap">${item.region}</span></td>
        <td class="p-3 font-medium text-gray-800 whitespace-nowrap">${item.field}</td>
        <td class="p-3 font-medium text-gray-800 leading-relaxed">${item.grievanceContent}</td>
        <td class="p-3 font-medium text-gray-800 leading-relaxed border-l border-gray-200"><span class="block mb-1 font-bold text-nyjSkyBlue">상담:</span>${item.consultContent}<br><span class="block mt-2 font-bold text-nyjBlue">처리:</span>${item.processContent}</td>
        <td class="p-3 text-center"><span class="text-[0.7rem] font-bold px-2 py-1 rounded-md border whitespace-nowrap ${badgeStyle}">${item.rawStatus}</span></td>
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
  document.getElementById('chart-donut-desc').innerText = `총 ${(totalOverall / 100000).toFixed(2)}억원`;

  Chart.defaults.font.family = "'Pretendard', sans-serif";
  Chart.defaults.color = "#475569"; // 진한 텍스트로 변경

  const donutCtx = document.getElementById('categoryDonutChart').getContext('2d');
  donutChartInst = new Chart(donutCtx, {
    type: 'doughnut',
    data: {
      labels: ['1. 기술지원', '2. 자금지원', '3. 판로개척', '4. 기업환경개선'],
      datasets: [{
        data: [catBudgets[1], catBudgets[2], catBudgets[3], catBudgets[4]],
        backgroundColor: ['#002978', '#009fe3', '#00a760', '#ef7d00'],
        borderWidth: 3,
        borderColor: '#ffffff',
        hoverOffset: 8
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { boxWidth: 12, padding: 15, font: { size: 11, weight: 'bold' } } },
        tooltip: {
          backgroundColor: 'rgba(255,255,255,0.95)',
          titleColor: '#0f172a',
          bodyColor: '#0f172a',
          borderColor: '#cbd5e1',
          borderWidth: 1,
          padding: 12,
          callbacks: {
            label: function(ctx) {
              const val = ctx.raw;
              const pct = ((val / totalOverall) * 100).toFixed(1);
              return ` ${ctx.label}: ${(val/100000).toFixed(2)}억원 (${pct}%)`;
            }
          }
        }
      },
      cutout: '68%'
    }
  });

  const barCtx = document.getElementById('categoryBarChart').getContext('2d');
  barChartInst = new Chart(barCtx, {
    type: 'bar',
    data: {
      labels: ['기술지원', '자금지원', '판로개척', '기업환경개선'],
      datasets: [
        { label: '시비', data: [catCity[1] / 100000, catCity[2] / 100000, catCity[3] / 100000, catCity[4] / 100000], backgroundColor: '#002978', borderRadius: 4 },
        { label: '도비', data: [catProv[1] / 100000, catProv[2] / 100000, catProv[3] / 100000, catProv[4] / 100000], backgroundColor: '#00a760', borderRadius: 4 },
        { label: '국비/기금', data: [catNatFund[1] / 100000, catNatFund[2] / 100000, catNatFund[3] / 100000, catNatFund[4] / 100000], backgroundColor: '#ef7d00', borderRadius: 4 }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { stacked: true, grid: { display: false } },
        y: { stacked: true, grid: { color: 'rgba(0,0,0,0.1)' }, ticks: { callback: v => v + '억' } }
      },
      plugins: {
        legend: { position: 'top', labels: { boxWidth: 12, font: { size: 11, weight: 'bold' } } }
      }
    }
  });
}

window.addEventListener('DOMContentLoaded', fetchExcelData);
