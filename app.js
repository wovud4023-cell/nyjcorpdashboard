// 동적 데이터 배열
const PROJECTS_DATA = [];
const GRIEVANCE_DATA = [];

// 카테고리 매핑용 (남양주 컬러 반영)
const catMap = {
    '기술지원': { id: 1, sub: '기술혁신과 R&D 역량을 키우는 남양주' },
    '자금지원': { id: 2, sub: '기업의 성장을 뒷받침하는 남양주' },
    '판로개척': { id: 3, sub: '국내외 시장으로 뻗어나가는 남양주' },
    '기업환경개선': { id: 4, sub: '기업하기 좋은 도시 남양주' }
};

// 남양주 전용 보조색상 팔레트 (Chart.js 용)
const nyjColors = ['#002978', '#009fe3', '#e6007e', '#ef7d00', '#00a760', '#b39b70'];

let donutChartInst = null;
let barChartInst = null;

// 모바일 사이드바 토글
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

// GitHub 엑셀 데이터 연동 (핵심 로직)
async function fetchExcelData() {
  try {
    const res = await fetch('./data.xlsx?v=' + new Date().getTime());
    if (!res.ok) {
        console.warn('data.xlsx 파일을 찾을 수 없습니다.');
        return;
    }
    
    const arrayBuffer = await res.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    
    // [1] 사업현황 시트 파싱
    if (workbook.SheetNames.includes('사업현황')) {
      const bizSheet = workbook.Sheets['사업현황'];
      const bizRows = XLSX.utils.sheet_to_json(bizSheet, { header: 1 });
      PROJECTS_DATA.length = 0;
      
      bizRows.forEach((cols, idx) => {
        if (idx === 0 || !cols[0]) return; // 헤더 스킵
        const catName = cols[0].toString().trim();
        const catInfo = catMap[catName] || { id: 99, sub: '' };
        const totalBudget = parseInt(cols[3]) || 0;
        
        PROJECTS_DATA.push({
          id: 'biz_' + idx,
          catId: catInfo.id,
          catName: catName,
          catSub: catInfo.sub,
          title: cols[1] || '-',
          shortTitle: cols[2] || '-',
          badge: totalBudget > 0 ? (totalBudget / 1000).toLocaleString() + '백만원' : '비예산',
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
          target: cols[11] || '-',
          applyInfo: cols[12] || '-',
          contact: cols[13] || '-'
        });
      });
      
      updateDashboardStats();
      renderSidebarMenu();
      renderDashboardCards();
      initCharts();
    }

    // [2] 애로사항 시트 파싱
    if (workbook.SheetNames.includes('애로사항')) {
       const worksheet = workbook.Sheets['애로사항'];
       const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
       GRIEVANCE_DATA.length = 0;
       
       rows.forEach((cols, idx) => {
         if (idx === 0 || !cols || cols.length === 0 || cols.every(cell => !cell)) return;
         GRIEVANCE_DATA.push({
           corpName: cols[1] || cols[0] || '-', 
           location: cols[2] || '-', 
           region: cols[3] || '기타',
           field: cols[4] || '-',
           grievanceContent: cols[5] || '-', 
           consultContent: cols[6] || '-', 
           processContent: cols[7] || '-',
           rawStatus: (cols[8] || cols[cols.length - 1] || '검토').toString().trim()
         });
       });
       
       updateRegionSelect();
       renderGrievanceTables(GRIEVANCE_DATA);
    }
  } catch (err) {
    console.error('엑셀 연동 오류:', err);
  }
}

function initApp() { fetchExcelData(); }

// 통계 요약 갱신
function updateDashboardStats() {
    let total = 0, city = 0, prov = 0;
    let catBudgets = {1:0, 2:0, 3:0, 4:0};

    PROJECTS_DATA.forEach(p => {
        total += p.finance.total;
        city += p.finance.city;
        prov += p.finance.prov;
        if(catBudgets[p.catId] !== undefined) catBudgets[p.catId] += p.finance.total;
    });

    const toUk = (v) => (v / 100000).toFixed(2) + '억원';
    
    document.getElementById('banner-total-budget').innerText = toUk(total);
    document.getElementById('banner-total-count').innerText = `총 지원예산 (${PROJECTS_DATA.length}개 사업)`;
    document.getElementById('banner-city-budget').innerText = toUk(city);
    document.getElementById('banner-prov-budget').innerText = toUk(prov);
    
    document.getElementById('sidebar-total-budget').innerText = toUk(total);
    document.getElementById('sidebar-cat1-budget').innerText = toUk(catBudgets[1]);
    document.getElementById('sidebar-cat2-budget').innerText = toUk(catBudgets[2]);
    document.getElementById('sidebar-cat3-budget').innerText = toUk(catBudgets[3]);
    document.getElementById('sidebar-cat4-budget').innerText = toUk(catBudgets[4]);
}

// 사이드바 메뉴 생성
function renderSidebarMenu() {
  const makeNavItem = (p) => `
    <li class="nav-item sub-item" id="nav-${p.id}" onclick="openProjectDetail('${p.id}')">
      <span style="overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${p.title}">${p.shortTitle}</span>
    </li>
  `;
  document.getElementById('dynamic-nav-cat1').innerHTML = PROJECTS_DATA.filter(p => p.catId === 1).map(makeNavItem).join('');
  document.getElementById('dynamic-nav-cat2').innerHTML = PROJECTS_DATA.filter(p => p.catId === 2).map(makeNavItem).join('');
  document.getElementById('dynamic-nav-cat3').innerHTML = PROJECTS_DATA.filter(p => p.catId === 3).map(makeNavItem).join('');
  document.getElementById('dynamic-nav-cat4').innerHTML = PROJECTS_DATA.filter(p => p.catId === 4).map(makeNavItem).join('');
}

// 메인 카드 생성
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
        <button class="tab-btn" style="background:var(--primary-bg); color:var(--primary); font-weight:700; border-radius:6px; padding:6px; width:100%; border:1px solid #bfdbfe;" onclick="openProjectDetail('${p.id}')">상세내용 보기</button>
      </div>
    `;
  }).join('');
  document.getElementById('searchResultCount').innerText = `총 ${PROJECTS_DATA.length}개 사업 표시 중`;
}

// 뷰 전환
function switchView(viewId) {
  document.querySelectorAll('.view-panel').forEach(el => el.classList.add('hidden'));
  document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
  const target = document.getElementById(`view-${viewId}`);
  if (target) target.classList.remove('hidden');
}

// 개별 상세 페이지 열기
function openProjectDetail(bizId) {
  const proj = PROJECTS_DATA.find(p => p.id === bizId);
  if (!proj) return;

  switchView('project-detail');
  document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
  const navEl = document.getElementById(`nav-${bizId}`);
  if (navEl) navEl.classList.add('active');

  document.getElementById('view-title').innerText = proj.title;
  document.getElementById('detail-proj-desc').innerText = proj.desc;
  document.getElementById('detail-proj-agency').innerText = proj.finance.agency;

  const fmt = v => v === 0 ? '-' : v.toLocaleString() + '천원';
  document.getElementById('fin-total').innerText = fmt(proj.finance.total);
  document.getElementById('fin-city').innerHTML = proj.finance.fund > 0 ? `${fmt(proj.finance.city)} <span style="font-size:0.75rem; color:#d97706;">(기금: ${fmt(proj.finance.fund)})</span>` : fmt(proj.finance.city);
  document.getElementById('fin-prov').innerText = fmt(proj.finance.prov);
  document.getElementById('fin-nat').innerText = fmt(proj.finance.nat);
  document.getElementById('fin-matching').innerText = proj.finance.matching;
  document.getElementById('fin-exec').innerText = proj.finance.agency;

  document.getElementById('detail-tab-plan').innerHTML = `
    <div class="content-box">
      <div class="box-header"><span class="box-title">세부 사업내용</span></div>
      <div class="box-desc">• ${proj.desc}</div>
    </div>
  `;
  document.getElementById('detail-history-custom').innerHTML = `
    <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:20px; line-height:1.7;">
      <h4 style="font-size:1rem; color:var(--primary); margin-bottom:12px; font-weight:700;">📌 신청방법 및 문의처 세부안내</h4>
      <p style="font-size:0.86rem; color:#334155; margin-bottom:8px;"><b>지원 대상:</b> ${proj.target}</p>
      <p style="font-size:0.86rem; color:#334155; margin-bottom:8px;"><b>신청 방법:</b> ${proj.applyInfo}</p>
      <p style="font-size:0.86rem; color:#334155; margin-bottom:8px;"><b>문의처:</b> ${proj.contact}</p>
    </div>
  `;
}

function switchDetailTab(type) {
  const isPlan = (type === 'plan');
  document.getElementById('tab-btn-plan').classList.toggle('active', isPlan);
  document.getElementById('tab-btn-history').classList.toggle('active', !isPlan);
  document.getElementById('detail-tab-plan').classList.toggle('hidden', !isPlan);
  document.getElementById('detail-tab-history').classList.toggle('hidden', isPlan);
}

// 검색 필터링
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

// 애로사항 렌더링
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

  let statusCounts = {'대기':0, '진행 중':0, '검토':0, '장기 검토':0, '불가':0, '완료':0};
  let regionCounts = {};

  dataList.forEach(item => {
    let stKey = item.rawStatus.includes('장기') ? '장기 검토' : item.rawStatus;
    if (statusCounts[stKey] !== undefined) statusCounts[stKey]++;
    
    let regKey = item.region || '기타';
    regionCounts[regKey] = (regionCounts[regKey] || 0) + 1;
  });

  regionBox.innerHTML = Object.entries(regionCounts).map(([reg, count]) => `
    <div class="content-box" style="text-align:center; padding:12px;"><div style="font-size:0.8rem; font-weight:700;">${reg} 지역</div><div style="font-size:1.3rem; color:#0284c7; font-weight:800;">${count}건</div></div>
  `).join('');

  summaryBox.innerHTML = Object.entries(statusCounts).map(([status, count]) => `
    <div class="content-box" style="text-align:center; padding:12px;"><div style="font-size:0.8rem; font-weight:700;">${status}</div><div style="font-size:1.3rem; color:var(--primary); font-weight:800;">${count}건</div></div>
  `).join('');

  let filteredData = dataList.filter(item => selectedRegion === 'all' || item.region === selectedRegion);

  filteredData.forEach((item, idx) => {
    let badgeStyle = 'background: #fef9c3; color: #a16207;'; // 기본 노란색
    if (item.rawStatus.includes('진행 중')) badgeStyle = 'background: #dcfce7; color: #15803d;';
    else if (item.rawStatus.includes('불가')) badgeStyle = 'background: #fee2e2; color: #dc2626;';
    else if (item.rawStatus.includes('완료')) badgeStyle = 'background: #dbeafe; color: #1d4ed8;';

    tbody.innerHTML += `
      <tr>
        <td style="text-align:center; font-weight:700;">${idx + 1}</td>
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

// 로컬 엑셀 업로드 테스트
function handleExcelUpload(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const workbook = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      if(workbook.SheetNames.includes('애로사항')) {
          GRIEVANCE_DATA.length = 0;
          XLSX.utils.sheet_to_json(workbook.Sheets['애로사항'], { header: 1 }).forEach((cols, idx) => {
            if (idx === 0 || !cols[0]) return;
            GRIEVANCE_DATA.push({
               corpName: cols[1]||'-', location: cols[2]||'-', region: cols[3]||'기타', field: cols[4]||'-', grievanceContent: cols[5]||'-', consultContent: cols[6]||'-', processContent: cols[7]||'-', rawStatus: (cols[8]||'검토').toString().trim()
            });
          });
          updateRegionSelect();
          renderGrievanceTables(GRIEVANCE_DATA);
          alert('로컬 파일 갱신 완료');
      }
    } catch (err) { alert('오류 발생'); }
  };
  reader.readAsArrayBuffer(file);
  event.target.value = '';
}

// Chart.js 초기화
function initCharts() {
  if(donutChartInst) donutChartInst.destroy();
  if(barChartInst) barChartInst.destroy();

  let catBudgets = {1:0, 2:0, 3:0, 4:0};
  let catCity = {1:0, 2:0, 3:0, 4:0};
  let catProv = {1:0, 2:0, 3:0, 4:0};
  let catNatFund = {1:0, 2:0, 3:0, 4:0};

  PROJECTS_DATA.forEach(p => {
      catBudgets[p.catId] += p.finance.total;
      catCity[p.catId] += p.finance.city;
      catProv[p.catId] += p.finance.prov;
      catNatFund[p.catId] += (p.finance.nat + p.finance.fund);
  });

  const totalOverall = Object.values(catBudgets).reduce((a,b)=>a+b, 0);
  document.getElementById('chart-donut-desc').innerText = `총 ${(totalOverall/1000).toLocaleString()}백만원 기준`;

  const donutCtx = document.getElementById('categoryDonutChart').getContext('2d');
  donutChartInst = new Chart(donutCtx, {
    type: 'doughnut',
    data: {
      labels: [`1. 기술지원`, `2. 자금지원`, `3. 판로개척`, `4. 기업환경개선`],
      datasets: [{
        data: [catBudgets[1], catBudgets[2], catBudgets[3], catBudgets[4]],
        backgroundColor: [nyjColors[1], nyjColors[2], nyjColors[4], nyjColors[3]],
        borderWidth: 2, borderColor: '#ffffff', hoverOffset: 6
      }]
    },
    options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } }, cutout: '66%' }
  });

  const barCtx = document.getElementById('categoryBarChart').getContext('2d');
  barChartInst = new Chart(barCtx, {
    type: 'bar',
    data: {
      labels: ['기술지원', '자금지원', '판로개척', '기업환경개선'],
      datasets: [
        { label: '시비', data: [catCity[1]/1000, catCity[2]/1000, catCity[3]/1000, catCity[4]/1000], backgroundColor: nyjColors[0] },
        { label: '도비', data: [catProv[1]/1000, catProv[2]/1000, catProv[3]/1000, catProv[4]/1000], backgroundColor: nyjColors[4] },
        { label: '국비/기금', data: [catNatFund[1]/1000, catNatFund[2]/1000, catNatFund[3]/1000, catNatFund[4]/1000], backgroundColor: nyjColors[3] }
      ]
    },
    options: { responsive: true, maintainAspectRatio: false, scales: { x: { stacked: true }, y: { stacked: true } } }
  });
}

window.addEventListener('DOMContentLoaded', initApp);
