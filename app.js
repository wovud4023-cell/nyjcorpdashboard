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

// 'sos8220'의 SHA-256 해시값 
const TARGET_HASH = "43ca49cf349f794f591fdc429a0801a628dfc7bbf4f5cbcc03c09578136f33d1"; 

function toggleSidebar() {
  const sidebar = document.getElementById('sidebar');
  const overlay = document.getElementById('overlay');
  if (window.innerWidth <= 768) {
    sidebar.classList.toggle('-translate-x-full');
    overlay.classList.toggle('hidden');
  } else {
    // Desktop: collapse to the left by applying negative margin
    sidebar.classList.toggle('-ml-80');
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
            el.innerHTML = current.toFixed(2) + '<span class="text-lg md:text-xl font-extrabold ml-1">억</span>';
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

  const countEl = document.getElementById('banner-total-count');
  if(countEl) countEl.innerText = `2026년 남양주시 4대 분야 ${PROJECTS_DATA.length}개 대표 시책 예산액 현황`;
  
  animateNumber('banner-total-budget', total / 100000, true);
  animateNumber('banner-city-budget', cityFund / 100000, true);
  
  const setTxt = (id, val) => { const e = document.getElementById(id); if(e) e.innerText = val; };
  setTxt('sidebar-total-budget', (total / 100000).toFixed(2) + '억');
  setTxt('sidebar-cat1-budget', (catBudgets[1] / 100000).toFixed(2) + '억');
  setTxt('sidebar-cat2-budget', (catBudgets[2] / 100000).toFixed(2) + '억');
  setTxt('sidebar-cat3-budget', (catBudgets[3] / 100000).toFixed(2) + '억');
  setTxt('sidebar-cat4-budget', (catBudgets[4] / 100000).toFixed(2) + '억');
}

function renderSidebarMenu() {
  const makeNavItem = (p, badgeClass) => `
    <div class="nav-item-hover px-2 md:px-3 py-2.5 md:py-3 rounded-xl flex justify-between items-center cursor-pointer transition-all border-l-4 border-transparent hover:border-gray-400 gap-2 md:gap-3 bg-white/40 mb-1.5 md:mb-2 shadow-sm" id="nav-${p.id}" onclick="openProjectDetail('${p.id}')">
      <span class="text-xs md:text-sm text-gray-900 font-extrabold truncate flex-1" title="${p.title}">· ${p.shortTitle}</span>
      <span class="text-[0.65rem] md:text-xs font-bold px-1.5 py-0.5 md:px-2 md:py-1 rounded-lg whitespace-nowrap shadow-sm ${badgeClass}">${p.badge}</span>
    </div>
  `;

  const setHTML = (id, html) => { const e = document.getElementById(id); if(e) e.innerHTML = html; };
  setHTML('dynamic-nav-cat1', PROJECTS_DATA.filter(p => p.catId === 1).map(p => makeNavItem(p, 'bg-blue-100 text-blue-800 border border-blue-200')).join(''));
  setHTML('dynamic-nav-cat2', PROJECTS_DATA.filter(p => p.catId === 2).map(p => makeNavItem(p, 'bg-purple-100 text-purple-800 border border-purple-200')).join(''));
  setHTML('dynamic-nav-cat3', PROJECTS_DATA.filter(p => p.catId === 3).map(p => makeNavItem(p, 'bg-green-100 text-green-800 border border-green-200')).join(''));
  setHTML('dynamic-nav-cat4', PROJECTS_DATA.filter(p => p.catId === 4).map(p => makeNavItem(p, 'bg-orange-100 text-orange-800 border border-orange-200')).join(''));
}

function renderDashboardCards() {
  const grid = document.getElementById('searchableCardsGrid');
  if(!grid) return;
  grid.innerHTML = PROJECTS_DATA.map(p => {
    let borderColor = p.catId === 2 ? 'border-t-purple-500' : p.catId === 3 ? 'border-t-green-500' : p.catId === 4 ? 'border-t-orange-500' : 'border-t-blue-600';
    let badgeBg = p.catId === 2 ? 'bg-purple-100 text-purple-800 border border-purple-200' : p.catId === 3 ? 'bg-green-100 text-green-800 border border-green-200' : p.catId === 4 ? 'bg-orange-100 text-orange-800 border border-orange-200' : 'bg-blue-100 text-blue-800 border border-blue-200';

    return `
      <div class="glass-card rounded-2xl md:rounded-3xl p-5 md:p-7 flex flex-col justify-between group border-t-[6px] md:border-t-8 ${borderColor} searchable-card shadow-md hover:shadow-xl" data-category="${p.catName}">
        <div>
          <div class="flex justify-between items-start mb-4 md:mb-5 gap-3">
            <h3 class="text-base md:text-xl font-extrabold text-gray-900 leading-tight group-hover:text-nyjBlue transition-colors">${p.title}</h3>
            <span class="text-xs md:text-sm font-bold px-2 py-1 md:px-3 md:py-1.5 rounded-lg ${badgeBg} whitespace-nowrap shadow-sm">${p.badge}</span>
          </div>
          <p class="text-sm md:text-base text-gray-800 font-bold mb-6 md:mb-8 leading-relaxed line-clamp-3">${p.desc}</p>
        </div>
        <div>
          <div class="text-xs md:text-sm text-gray-800 font-bold mb-4 md:mb-5 bg-white/80 p-2 md:p-3 rounded-xl border border-gray-200 shadow-sm">
            <span class="font-extrabold text-gray-900">수행기관:</span> ${p.finance.agency}
          </div>
          <button class="w-full py-2.5 md:py-3.5 bg-gradient-to-r from-blue-50 to-indigo-50 hover:from-blue-100 hover:to-indigo-100 text-nyjBlue font-extrabold rounded-xl md:rounded-2xl text-sm md:text-base transition-all border border-blue-200 shadow-md focus:ring-4 focus:ring-blue-100" onclick="openProjectDetail('${p.id}')">
            상세내용 및 문서 관리 &rarr;
          </button>
        </div>
      </div>
    `;
  }).join('');
  const cnt = document.getElementById('searchResultCount');
  if(cnt) cnt.innerText = `총 ${PROJECTS_DATA.length}개 사업 검색됨`;
}

// SHA-256 해시 함수
async function hashPassword(message) {
  const msgBuffer = new TextEncoder('utf-8').encode(message);
  const hashBuffer = await crypto.subtle.digest('SHA-256', msgBuffer);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

// 전체 대시보드 비밀번호 검증
async function verifyPassword() {
  const inputEl = document.getElementById("auth-password");
  const inputHash = await hashPassword(inputEl.value);
  
  if (inputHash === TARGET_HASH) {
    // 앱 전체 잠금 해제
    document.getElementById('login-screen').classList.add('hidden');
    
    const appContainer = document.getElementById('app-container');
    appContainer.classList.remove('hidden');
    appContainer.classList.add('flex');
    
    // 데이터 로딩 시작
    fetchExcelData();
  } else {
    alert("인증번호가 일치하지 않습니다.");
    inputEl.value = "";
    inputEl.focus();
  }
}

function switchView(viewId) {
  document.querySelectorAll('.view-panel').forEach(el => {
      el.classList.add('hidden');
      el.classList.remove('active');
  });
  
  document.querySelectorAll('.nav-item-active').forEach(el => {
      if(el.id === 'nav-dashboard') {
          el.classList.remove('nav-item-active');
          el.classList.add('text-gray-900');
      }
  });

  const target = document.getElementById(`view-${viewId}`);
  if (target) {
      target.classList.remove('hidden');
      target.classList.add('active');
  }

  const meta = {
    'dashboard': { title: '2026 기업 지원시책 대시보드', desc: '기술지원·자금지원·판로개척·기업환경개선 종합 현황 분석' },
    'grievance': { title: '기업애로사항 접수·관리', desc: '관내 기업체 현장 애로사항 목록 및 통계를 조회합니다.' },
    'project-detail': { title: '사업 상세정보', desc: '해당 시책의 세부 추진계획 및 관련 문서를 관리합니다.' }
  };

  if (meta[viewId]) {
    const titleEl = document.getElementById('view-title');
    const descEl = document.getElementById('view-desc');
    if(titleEl) titleEl.innerText = meta[viewId].title;
    if(descEl) descEl.innerText = meta[viewId].desc;
  }
  
  if (viewId === 'dashboard') {
    const dNav = document.getElementById('nav-dashboard');
    if(dNav) {
        dNav.classList.add('nav-item-active');
        dNav.classList.remove('text-gray-900');
    }
  }
}

function openProjectDetail(bizId) {
  const proj = PROJECTS_DATA.find(p => p.id === bizId);
  if (!proj) return;

  switchView('project-detail');

  const setTxt = (id, val) => { const e = document.getElementById(id); if(e) e.innerText = val; };
  setTxt('view-title', proj.title);
  setTxt('view-desc', `${proj.catName} - ${proj.catSub} | ${proj.desc}`);
  setTxt('detail-proj-desc', proj.desc);
  setTxt('detail-proj-agency', proj.finance.agency);

  const fmt = v => v === 0 ? '-' : v.toLocaleString() + '천원';
  setTxt('fin-total', fmt(proj.finance.total));
  
  const finCity = document.getElementById('fin-city');
  if(finCity) finCity.innerHTML = proj.finance.fund > 0 ? `${fmt(proj.finance.city)} <br><span class="text-xs md:text-sm text-orange-600 font-extrabold mt-1 inline-block">(기금: ${fmt(proj.finance.fund)})</span>` : fmt(proj.finance.city);
  
  setTxt('fin-prov', fmt(proj.finance.prov));
  setTxt('fin-nat', fmt(proj.finance.nat));
  setTxt('fin-matching', proj.finance.matching);
  setTxt('fin-exec', proj.finance.agency);

  const planEl = document.getElementById('detail-tab-plan');
  if(planEl) {
      planEl.innerHTML = proj.plans.map(pl => `
        <div class="bg-white p-5 md:p-6 rounded-2xl md:rounded-3xl border border-gray-200 shadow-md">
          <div class="flex justify-between items-start mb-3 md:mb-4 gap-3">
            <span class="font-extrabold text-nyjBlue text-sm md:text-base">${pl.title}</span>
            <span class="text-xs md:text-sm font-bold bg-green-100 text-green-800 border border-green-200 px-2.5 py-1 md:px-3 md:py-1.5 rounded-lg whitespace-nowrap">${pl.budget}</span>
          </div>
          <div class="text-xs md:text-base text-gray-800 font-bold leading-relaxed whitespace-pre-wrap">${pl.desc}</div>
        </div>
      `).join('');
  }

  const histEl = document.getElementById('detail-history-custom');
  if(histEl) {
      histEl.innerHTML = `
        <h4 class="text-sm md:text-base font-extrabold text-nyjBlue mb-3 md:mb-4 border-l-[4px] md:border-l-[5px] border-nyjSkyBlue pl-2 md:pl-3">신청방법 및 문의처 세부안내</h4>
        <div class="space-y-2 md:space-y-3 text-sm md:text-base text-gray-800 font-bold bg-gray-50 p-4 md:p-5 rounded-xl md:rounded-2xl border border-gray-100">
            <p><b class="text-gray-900">지원 대상:</b> ${proj.target}</p>
            <p><b class="text-gray-900">신청 방법:</b> ${proj.applyInfo}</p>
            <p><b class="text-gray-900">문의 부서:</b> ${proj.contact}</p>
        </div>
      `;
  }

  renderAttachmentRows(proj);
  switchDetailTab('plan');
}

function switchDetailTab(type) {
  const isPlan = (type === 'plan');
  
  const btnPlan = document.getElementById('tab-btn-plan');
  const btnHist = document.getElementById('tab-btn-history');
  
  if(!btnPlan || !btnHist) return;

  if(isPlan) {
      btnPlan.className = "flex-1 py-4 md:py-5 text-sm md:text-lg font-extrabold text-nyjBlue border-b-4 border-nyjBlue bg-white transition-colors";
      btnHist.className = "flex-1 py-4 md:py-5 text-sm md:text-lg font-extrabold text-gray-500 border-b-4 border-transparent hover:bg-gray-50 transition-colors";
      document.getElementById('detail-tab-plan').classList.remove('hidden');
      document.getElementById('detail-tab-history').classList.add('hidden');
  } else {
      btnHist.className = "flex-1 py-4 md:py-5 text-sm md:text-lg font-extrabold text-nyjBlue border-b-4 border-nyjBlue bg-white transition-colors";
      btnPlan.className = "flex-1 py-4 md:py-5 text-sm md:text-lg font-extrabold text-gray-500 border-b-4 border-transparent hover:bg-gray-50 transition-colors";
      document.getElementById('detail-tab-history').classList.remove('hidden');
      document.getElementById('detail-tab-plan').classList.add('hidden');
  }
}

const FILE_CATS = [
  { key: 'official', name: '사업관련 공문', badge: 'bg-blue-100 text-blue-800 border-blue-200' },
  { key: 'plan', name: '추진계획/계획서', badge: 'bg-green-100 text-green-800 border-green-200' },
  { key: 'status', name: '추진현황/정산서', badge: 'bg-orange-100 text-orange-800 border-orange-200' }
];

function renderAttachmentRows(proj) {
  const cont = document.getElementById('detail-attachments-container');
  if(!cont) return;
  cont.innerHTML = FILE_CATS.map(cat => {
    const filename = proj.docs[cat.key];
    const isRegistered = Boolean(filename && filename.length > 0);
    const safeFilename = encodeURIComponent(filename);

    return `
      <div class="flex flex-col md:flex-row justify-between items-start md:items-center p-4 md:p-5 bg-white rounded-xl md:rounded-2xl border border-gray-200 shadow-sm gap-4 md:gap-5 hover:shadow-md transition-shadow">
        <div>
          <div class="flex items-center gap-2 md:gap-3 mb-1 md:mb-2">
            <span class="text-[0.65rem] md:text-xs font-extrabold border px-2 py-1 md:px-3 rounded-lg shadow-sm ${cat.badge}">${cat.name}</span>
            <span class="text-sm md:text-base font-extrabold text-gray-900">문서 연동 상태</span>
          </div>
          <div class="text-xs md:text-sm font-bold text-gray-800 mt-1 md:mt-0">
            ${isRegistered 
              ? `<span class="text-green-700 font-black mr-2 bg-green-50 px-2 py-1 rounded-md border border-green-100">✓ 등록됨</span> <span class="text-gray-700">${filename}</span>` 
              : `<span class="text-gray-500 italic bg-gray-50 px-2 py-1 rounded-md">미등록 (data.xlsx에 파일명 입력 필요)</span>`}
          </div>
        </div>
        <button class="w-full md:w-auto px-5 py-3 md:px-6 md:py-3.5 text-xs md:text-sm font-extrabold rounded-xl transition-colors flex items-center justify-center gap-2 shadow-sm ${isRegistered ? 'bg-nyjBlue text-white hover:bg-blue-800 focus:ring-4 focus:ring-blue-200' : 'bg-gray-200 text-gray-400 cursor-not-allowed border border-gray-300'}" ${isRegistered ? `onclick="window.open('./${safeFilename}', '_blank')"` : 'disabled'}>
          <svg class="w-4 h-4 md:w-5 md:h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"></path></svg>
          ${isRegistered ? '서버에서 다운로드' : '다운로드 불가'}
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
  const cntEl = document.getElementById('searchResultCount');
  if(cntEl) cntEl.innerText = `검색 결과: ${count}건`;
}

function updateRegionSelect() {
  const select = document.getElementById('regionFilter');
  if(!select) return;
  const regions = [...new Set(GRIEVANCE_DATA.map(item => item.region || '기타'))].sort();
  select.innerHTML = '<option value="all">📍 전체 지역 검색</option>' + regions.map(r => `<option value="${r}">${r}</option>`).join('');
}

function renderGrievanceTables(dataList) {
  const tbody = document.getElementById('grievance-tbody');
  const summaryBox = document.getElementById('grievance-summary-content');
  const regionBox = document.getElementById('region-summary-content');
  const select = document.getElementById('regionFilter');
  const selectedRegion = select ? select.value : 'all';
  
  if(!tbody || !summaryBox || !regionBox) return;
  
  tbody.innerHTML = '';

  if (!dataList || dataList.length === 0) {
    tbody.innerHTML = `<tr><td colspan="9" class="text-center p-6 md:p-10 text-gray-500 font-extrabold text-base md:text-lg bg-white">데이터가 없습니다.</td></tr>`;
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
    <div class="bg-white border-2 border-gray-100 p-3 md:p-4 rounded-xl md:rounded-2xl text-center shadow-sm hover:border-nyjGreen transition-colors">
      <div class="text-xs md:text-sm font-extrabold text-gray-600 mb-1 md:mb-2">${reg}</div>
      <div class="text-xl md:text-2xl font-black text-nyjGreen">${count}건</div>
    </div>
  `).join('');

  summaryBox.innerHTML = Object.entries(statusCounts).map(([status, count]) => `
    <div class="bg-white border-2 border-gray-100 p-3 md:p-4 rounded-xl md:rounded-2xl text-center shadow-sm hover:border-nyjBlue transition-colors">
      <div class="text-xs md:text-sm font-extrabold text-gray-600 mb-1 md:mb-2">${status}</div>
      <div class="text-xl md:text-2xl font-black text-nyjBlue">${count}건</div>
    </div>
  `).join('');

  let filteredData = dataList.filter(item => selectedRegion === 'all' || item.region === selectedRegion);

  filteredData.forEach((item, idx) => {
    let badgeStyle = 'bg-yellow-100 text-yellow-800 border-yellow-300';
    if (item.rawStatus.includes('진행 중')) badgeStyle = 'bg-green-100 text-green-800 border-green-300';
    else if (item.rawStatus.includes('불가')) badgeStyle = 'bg-red-100 text-red-800 border-red-300';
    else if (item.rawStatus.includes('완료')) badgeStyle = 'bg-blue-100 text-blue-800 border-blue-300';

    tbody.innerHTML += `
      <tr class="hover:bg-blue-50/50 transition-colors bg-white group border-b-2 border-gray-100">
        <td class="p-3 md:p-4 text-center font-extrabold text-gray-500 whitespace-nowrap">${idx + 1}</td>
        <td class="p-3 md:p-4 font-black text-gray-900 whitespace-nowrap">${item.corpName}</td>
        <td class="p-3 md:p-4 font-bold text-gray-700 whitespace-nowrap">${item.location}</td>
        <td class="p-3 md:p-4"><span class="text-[0.65rem] md:text-xs px-2 md:px-2.5 py-1 md:py-1.5 bg-gray-200 border border-gray-300 rounded-lg font-extrabold text-gray-800 whitespace-nowrap shadow-sm">${item.region}</span></td>
        <td class="p-3 md:p-4 font-bold text-gray-700 whitespace-nowrap">${item.field}</td>
        <td class="p-3 md:p-4 font-bold text-gray-800 leading-loose min-w-[200px]">${item.grievanceContent}</td>
        <td class="p-3 md:p-4 font-bold text-gray-800 leading-loose border-l-2 border-gray-100 min-w-[250px]">
            <div class="bg-blue-50/50 p-2 rounded-lg mb-2 border border-blue-100"><span class="font-black text-nyjSkyBlue mr-2">상담:</span>${item.consultContent}</div>
            <div class="bg-indigo-50/50 p-2 rounded-lg border border-indigo-100"><span class="font-black text-nyjBlue mr-2">처리:</span>${item.processContent}</div>
        </td>
        <td class="p-3 md:p-4 text-center"><span class="text-xs md:text-sm font-extrabold px-2.5 md:px-3 py-1 md:py-1.5 rounded-xl border-2 whitespace-nowrap shadow-sm ${badgeStyle}">${item.rawStatus}</span></td>
      </tr>
    `;
  });
  const gCount = document.getElementById('grievance-count');
  if(gCount) gCount.innerText = `현재 표시: ${filteredData.length}건`;
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
      alert('업로드된 엑셀 파일로 현황판이 성공적으로 갱신되었습니다.');
    } catch (err) { alert('엑셀 처리 중 오류가 발생했습니다.'); }
  };
  reader.readAsArrayBuffer(file);
  event.target.value = '';
}

function clearGrievanceData() {
  fetchExcelData();
  alert('깃허브(서버) 원본 데이터로 복원되었습니다.');
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
  const chartDesc = document.getElementById('chart-donut-desc');
  if(chartDesc) chartDesc.innerText = `총 ${(totalOverall / 100000).toFixed(2)}억원`;

  Chart.defaults.font.family = "'Pretendard', sans-serif";
  Chart.defaults.color = "#334155"; // 더욱 진한 차트 텍스트

  const donutCtx = document.getElementById('categoryDonutChart');
  if(donutCtx) {
      donutChartInst = new Chart(donutCtx.getContext('2d'), {
        type: 'doughnut',
        data: {
          labels: ['1. 기술지원', '2. 자금지원', '3. 판로개척', '4. 기업환경개선'],
          datasets: [{
            data: [catBudgets[1], catBudgets[2], catBudgets[3], catBudgets[4]],
            backgroundColor: ['#002978', '#009fe3', '#00a760', '#ef7d00'],
            borderWidth: 4,
            borderColor: '#ffffff',
            hoverOffset: 12
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { position: 'bottom', labels: { boxWidth: 14, padding: 18, font: { size: window.innerWidth > 768 ? 14 : 12, weight: '900' }, color: '#1e293b' } },
            tooltip: {
              backgroundColor: 'rgba(255,255,255,0.95)',
              titleColor: '#0f172a',
              bodyColor: '#0f172a',
              borderColor: '#cbd5e1',
              borderWidth: 2,
              padding: 16,
              titleFont: { size: 16, weight: 'bold' },
              bodyFont: { size: 15, weight: 'bold' },
              callbacks: {
                label: function(ctx) {
                  const val = ctx.raw;
                  const pct = ((val / totalOverall) * 100).toFixed(1);
                  return ` ${ctx.label}: ${(val/100000).toFixed(2)}억원 (${pct}%)`;
                }
              }
            }
          },
          cutout: '65%'
        }
      });
  }

  const barCtx = document.getElementById('categoryBarChart');
  if(barCtx) {
      barChartInst = new Chart(barCtx.getContext('2d'), {
        type: 'bar',
        data: {
          labels: ['기술지원', '자금지원', '판로개척', '기업환경개선'],
          datasets: [
            { label: '시비', data: [catCity[1] / 100000, catCity[2] / 100000, catCity[3] / 100000, catCity[4] / 100000], backgroundColor: '#002978', borderRadius: 6 },
            { label: '도비', data: [catProv[1] / 100000, catProv[2] / 100000, catProv[3] / 100000, catProv[4] / 100000], backgroundColor: '#00a760', borderRadius: 6 },
            { label: '국비/기금', data: [catNatFund[1] / 100000, catNatFund[2] / 100000, catNatFund[3] / 100000, catNatFund[4] / 100000], backgroundColor: '#ef7d00', borderRadius: 6 }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          scales: {
            x: { stacked: true, grid: { display: false }, ticks: { font: { size: window.innerWidth > 768 ? 14 : 11, weight: 'bold' }, color: '#1e293b' } },
            y: { stacked: true, grid: { color: 'rgba(0,0,0,0.1)' }, ticks: { font: { size: window.innerWidth > 768 ? 14 : 11, weight: 'bold' }, color: '#1e293b', callback: v => v + '억' } }
          },
          plugins: {
            legend: { position: 'top', labels: { boxWidth: 14, font: { size: window.innerWidth > 768 ? 14 : 12, weight: '900' }, color: '#1e293b' } },
            tooltip: {
                backgroundColor: 'rgba(255,255,255,0.95)',
                titleColor: '#0f172a',
                bodyColor: '#0f172a',
                borderColor: '#cbd5e1',
                borderWidth: 2,
                padding: 16,
                titleFont: { size: 16, weight: 'bold' },
                bodyFont: { size: 15, weight: 'bold' }
            }
          }
        }
      });
  }
}
