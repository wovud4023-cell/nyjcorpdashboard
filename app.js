
async function fetchExcelFromGitHub() {
  try {
    // data.xlsx 파일 자동 로드 (캐시 무효화 파라미터 추가)
    const res = await fetch('./data.xlsx?v=' + new Date().getTime());
    if (!res.ok) throw new Error('업로드된 data.xlsx 파일을 찾을 수 없습니다. 기본 데이터를 사용합니다.');
    
    const arrayBuffer = await res.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    
    let isDataUpdated = false;

    // [1] '사업현황' 시트가 있으면 대시보드 데이터(지원팀, SOS팀) 업데이트
    if (workbook.SheetNames.includes('사업현황')) {
      const bizSheet = workbook.Sheets['사업현황'];
      const bizRows = XLSX.utils.sheet_to_json(bizSheet, { header: 1 });
      
      let newSupport = [];
      let newProjects = [];
      
      bizRows.forEach((cols, idx) => {
        if (idx === 0) return; // 헤더 스킵
        if (!cols[0]) return; // 팀구분 없으면 스킵
        
        const team = cols[0].toString().trim(); // '지원팀' or 'SOS팀'
        const obj = {
          id: 'biz_' + idx,
          team: team,
          title: cols[1] || '-',
          shortTitle: cols[2] || '-',
          badge: '총 ' + (Math.round(parseInt(cols[3] || 0) / 1000)).toLocaleString() + '백만',
          finance: { 
            total: parseInt(cols[3] || 0), 
            city: parseInt(cols[4] || 0), 
            prov: parseInt(cols[5] || 0), 
            nat: parseInt(cols[6] || 0), 
            matching: cols[8] || '-', 
            agency: cols[7] || '-' 
          },
          desc: cols[9] || '-',
          plans: [{ title: '엑셀 데이터 연동됨', budget: '확인완료', desc: 'GitHub에 업로드된 엑셀(data.xlsx)에서 로드된 자료입니다.' }]
        };
        
        if (team === '지원팀') newSupport.push(obj);
        else newProjects.push(obj);
      });
      
      if (newSupport.length > 0) {
         SUPPORT_TEAM_DATA.length = 0; 
         SUPPORT_TEAM_DATA.push(...newSupport);
      }
      if (newProjects.length > 0) {
         PROJECTS_DATA.length = 0; 
         PROJECTS_DATA.push(...newProjects);
      }
      isDataUpdated = true;
    }

    // [2] '애로사항' 시트가 있으면 민원 현황 업데이트
    const grievanceSheetName = workbook.SheetNames.includes('애로사항') ? '애로사항' : (workbook.SheetNames.includes('사업현황') ? null : workbook.SheetNames[0]);
    
    if (grievanceSheetName) {
       const worksheet = workbook.Sheets[grievanceSheetName];
       const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
       let parsedData = [];
       rows.forEach((cols, idx) => {
         if (idx === 0) return;
         if (!cols || cols.length === 0 || cols.every(cell => !cell)) return;
         parsedData.push({
           corpName: cols[1] || cols[0] || '-',
           location: cols[2] || '-',
           field: cols[3] || '-',
           grievanceContent: cols[4] || '-',
           consultContent: cols[5] || '-',
           processContent: cols[6] || '-',
           rawStatus: (cols[7] || cols[cols.length - 1] || '검토').toString().trim()
         });
       });
       
       if (parsedData.length > 0) {
         renderGrievanceTable(parsedData);
         isDataUpdated = true;
       } else {
         loadGrievanceData(); 
       }
    } else {
       loadGrievanceData(); 
    }

    // 데이터가 업데이트 되었으면 UI 다시 그리기
    if (isDataUpdated) {
      renderSidebarMenu();
      renderDashboardCards();
      initCharts();
      console.log('GitHub data.xlsx 파일로부터 데이터를 성공적으로 업데이트 했습니다.');
    }

  } catch (err) {
    console.warn(err.message);
    loadGrievanceData(); // 엑셀 파일이 없으면 기존 로직(data.js 및 로컬스토리지) 실행
  }
}
let currentDetailBizId = null;
let singleFiles = {};

function toggleSidebarDesktop() {
  const aside = document.querySelector('aside');
  aside.classList.toggle('desktop-collapsed');
}

function toggleSidebar() {
  const aside = document.querySelector('aside');
  aside.classList.toggle('collapsed');
}

function initApp() {
  loadFilesFromStorage();
  renderSidebarMenu();
  renderDashboardCards();
  initCharts();
  
  // 기존 방식인 loadGrievanceData() 대신 깃허브 엑셀 자동 로드 호출
  // 엑셀이 없으면 내부적으로 기존 데이터를 로드합니다.
  fetchExcelFromGitHub();
}

function renderSidebarMenu() {
  document.getElementById('dynamic-nav-support').innerHTML = SUPPORT_TEAM_DATA.map(s => `
    <li class="nav-item" id="nav-${s.id}" onclick="openProjectDetail('${s.id}')">
      <span>${s.title}</span>
      <span class="nav-badge" style="background:#0284c7; color:#fff;">${s.badge}</span>
    </li>
  `).join('');

  document.getElementById('dynamic-nav-projects').innerHTML = PROJECTS_DATA.map(p => `
    <li class="nav-item" id="nav-${p.id}" onclick="openProjectDetail('${p.id}')">
      <span>${p.title}</span>
      <span class="nav-badge">${p.badge}</span>
    </li>
  `).join('');
}

function renderDashboardCards() {
  const allData = [...SUPPORT_TEAM_DATA, ...PROJECTS_DATA];
  document.getElementById('searchableCardsGrid').innerHTML = allData.map(p => `
    <div class="content-box searchable-card" data-category="${p.team}" data-keywords="${p.title} ${p.desc} ${p.finance.agency}">
      <div class="box-header">
        <span class="box-title">${p.title}</span>
        <span class="box-badge ${p.team === '지원팀' ? 'badge-blue' : 'badge-green'}">${p.badge}</span>
      </div>
      <div class="box-desc">• ${p.desc}</div>
      <div style="font-size:0.78rem; color:#64748b; margin-bottom:8px;">
        <b>수행:</b> ${p.finance.agency} | <b>팀:</b> ${p.team === '지원팀' ? '기업지원팀' : '기업SOS팀'}
      </div>
      <button class="tab-btn" style="padding:4px 8px; font-size:0.78rem; background:#eff6ff; color:#1e40af; border-radius:4px;" onclick="openProjectDetail('${p.id}')">상세보기 &rarr;</button>
    </div>
  `).join('');
}

function switchView(viewId) {
  document.querySelectorAll('.view-panel').forEach(el => el.classList.add('hidden'));
  document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));

  const target = document.getElementById(`view-${viewId}`);
  if (target) target.classList.remove('hidden');

  const meta = {
    'dashboard': { title: '예산·실적 종합 대시보드', desc: '기업지원팀 및 기업SOS팀 소관 지원사업의 2026~2027년 예산 편성 현황 및 실적 통계를 분석합니다.' },
    'grievance': { title: '💬 기업애로사항 접수·관리', desc: '관내 기업체 현장 애로사항 접수 목록 및 상태별 통계를 조회하고 엑셀 파일로 관리합니다.' },
    'manual':    { title: '📖 업무매뉴얼', desc: '맞춤형 기업애로 해결사업(e호조) 및 기업환경개선사업(보탬e) 업무 지침 가이드' }
  };

  if (meta[viewId]) {
    document.getElementById('view-title').innerText = meta[viewId].title;
    document.getElementById('view-desc').innerText = meta[viewId].desc;
  }
  if (viewId === 'dashboard') {
    document.querySelector('.nav-list .nav-item:first-of-type').classList.add('active');
  }
  
  if (window.innerWidth <= 768) {
    document.querySelector('aside').classList.add('collapsed');
  }
}

function openProjectDetail(bizId) {
  currentDetailBizId = bizId;
  const allData = [...SUPPORT_TEAM_DATA, ...PROJECTS_DATA];
  const proj = allData.find(p => p.id === bizId);
  if (!proj) return;

  document.querySelectorAll('.view-panel').forEach(el => el.classList.add('hidden'));
  document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));

  const navEl = document.getElementById(`nav-${bizId}`);
  if (navEl) navEl.classList.add('active');

  document.getElementById('view-title').innerText = proj.title;
  document.getElementById('view-desc').innerText = proj.desc;
  document.getElementById('detail-proj-desc').innerText = proj.desc;
  document.getElementById('detail-proj-agency').innerText = proj.finance.agency;

  const fmt = (v) => v === 0 ? '-' : v.toLocaleString() + '천원';
  document.getElementById('fin-total').innerText = proj.finance.total === 0 ? '비예산' : fmt(proj.finance.total);
  document.getElementById('fin-city').innerText = fmt(proj.finance.city);
  document.getElementById('fin-prov').innerText = fmt(proj.finance.prov);
  document.getElementById('fin-nat').innerText = fmt(proj.finance.nat);
  document.getElementById('fin-matching').innerText = proj.finance.matching;
  document.getElementById('fin-exec').innerText = proj.finance.agency;

  document.getElementById('detail-tab-plan').innerHTML = proj.plans.map(pl => `
    <div class="content-box">
      <div class="box-header">
        <span class="box-title">${pl.title}</span>
        <span class="box-badge badge-green">${pl.budget}</span>
      </div>
      <div class="box-desc">• ${pl.desc}</div>
    </div>
  `).join('');

  const historyCustom = document.getElementById('detail-history-custom');
  const table = document.getElementById('detail-history-table');
  historyCustom.innerHTML = '';
  table.innerHTML = '';

  if (proj.team === '지원팀') {
    historyCustom.innerHTML = `
      <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:16px; margin-bottom:14px;">
        <h4 style="font-size:0.95rem; color:#0369a1; margin-bottom:10px;">📋 기업지원팀 소관 사업 안내</h4>
        <p style="font-size:0.85rem; color:#475569;">본 사업은 남양주시청 기업지원팀 소관 2026년 추진 사업입니다.</p>
      </div>
    `;
    table.innerHTML = `<tr><td style="text-align:center; padding:20px; color:#94a3b8;">등록된 추진 실적 데이터가 없습니다.</td></tr>`;
  } else if (bizId === 'biz10') {
    historyCustom.innerHTML = `
      <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:16px; margin-bottom:14px;">
        <h4 style="font-size:0.95rem; color:#1e3a8a; margin-bottom:10px;">📊 인증 완료 기업 현황 (총 99개사)</h4>
        <div style="display:flex; gap:8px; flex-wrap:wrap;">
          <span class="box-badge badge-blue">제조업: 42</span>
          <span class="box-badge badge-green">도·소매업: 17</span>
          <span class="box-badge badge-amber">요식업: 3</span>
          <span class="box-badge badge-purple">정보통신업: 4</span>
          <span class="box-badge badge-slate">기타 분야: 44</span>
        </div>
      </div>
    `;
  }

  if (proj.history) {
    table.innerHTML = `
      <thead><tr>${proj.history.headers.map(h => `<th>${h}</th>`).join('')}</tr></thead>
      <tbody>${proj.history.rows.map(row => `<tr>${row.map(c => `<td>${c}</td>`).join('')}</tr>`).join('')}</tbody>
    `;
  }

  const isSarangbang = (bizId === 'biz9');
  document.getElementById('sarangbang-application-section').classList.toggle('hidden', !isSarangbang);
  document.getElementById('doc-card-title').innerText = isSarangbang 
    ? '📎 기업사랑방 온라인 상담 신청 및 핵심 문서 보관함' 
    : '📎 사업 관련 핵심 문서 관리 (공문 · 추진계획 · 추진현황)';

  renderAttachmentRows(bizId);
  switchDetailTab('plan');
  document.getElementById('view-project-detail').classList.remove('hidden');
  
  if (window.innerWidth <= 768) {
    document.querySelector('aside').classList.add('collapsed');
  }
}

function switchDetailTab(type) {
  const isPlan = (type === 'plan');
  document.getElementById('tab-btn-plan').classList.toggle('active', isPlan);
  document.getElementById('tab-btn-history').classList.toggle('active', !isPlan);
  document.getElementById('detail-tab-plan').classList.toggle('hidden', !isPlan);
  document.getElementById('detail-tab-history').classList.toggle('hidden', isPlan);
}

async function handleSarangbangSubmit(event) {
  event.preventDefault();

  const corpName = document.getElementById('sb_corpName').value.trim();
  const bizNum = document.getElementById('sb_bizNum').value.trim();
  const ceoName = document.getElementById('sb_ceoName').value.trim();
  const phone = document.getElementById('sb_phone').value.trim();
  const field = document.getElementById('sb_field').value;
  const address = document.getElementById('sb_address').value.trim();
  const content = document.getElementById('sb_content').value.trim();

  const now = new Date();
  const receiptNo = `SARANG-${now.getFullYear()}-${Math.floor(1000 + Math.random() * 9000)}`;

  const fileText = 
`===================================================================
 [남양주시] 지식나눔 기업사랑방 온라인 상담 사전접수 신청서
===================================================================

[접수 정보]
- 접수 번호 : ${receiptNo}
- 접수 일시 : ${now.toLocaleDateString('ko-KR')} ${now.toLocaleTimeString('ko-KR')}
- 처리 부서 : 남양주시청 지역경제과 기업SOS팀 (☎ 031-590-8220)

-------------------------------------------------------------------
[신청 기업 정보]
-------------------------------------------------------------------
- 상 호 명(기업명) : ${corpName}
- 사업자 등록번호  : ${bizNum}
- 대  표  자  명   : ${ceoName}
- 신청인 연락처    : ${phone}
- 사업장 소재지    : ${address}

-------------------------------------------------------------------
[상담 신청 내용]
-------------------------------------------------------------------
- 희망 상담 분야   : ${field}
- 상담 요청 내용 :
${content}

-------------------------------------------------------------------
[안내 사항]
-------------------------------------------------------------------
1. 본 신청서는 기업사랑방 기초 무료상담을 위한 사전 접수 내역입니다.
2. 실무 적용(소송, 등기, 기장 등 실제 대행 행위)에 따른 비용은 지원되지 않습니다.
3. 접수 후 기업SOS팀 담당자가 접수 내용을 검토하여 매칭 일정을 안내해 드립니다.
===================================================================`;

  const safeCorpName = corpName.replace(/[^a-zA-Z0-9가-힣]/g, '_');
  const fileName = `기업사랑방_상담접수증_${safeCorpName}_${receiptNo}.txt`;

  const blob = new Blob(["\uFEFF" + fileText], { type: 'text/plain;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(url);

  singleFiles['biz9_status'] = {
    name: fileName,
    size: Math.round(blob.size / 1024) + ' KB',
    date: now.toISOString().slice(0, 10),
    dataUrl: 'data:text/plain;charset=utf-8;base64,' + btoa(unescape(encodeURIComponent("\uFEFF" + fileText)))
  };
  saveFilesToStorage();
  renderAttachmentRows('biz9');

  alert(`[접수 완료] ${corpName}님의 상담 신청서가 접수되었습니다.\n\nPC에 '${fileName}'이 저장되었으며, 하단 현황 보관함에 자동 등록되었습니다.`);
  document.getElementById('sarangbangForm').reset();
}

const FILE_CATS = [
  { key: 'official', name: '사업관련 공문', badge: 'badge-blue' },
  { key: 'plan', name: '추진계획', badge: 'badge-green' },
  { key: 'status', name: '추진현황 / 접수내역', badge: 'badge-amber' }
];

function loadFilesFromStorage() {
  try {
    const saved = localStorage.getItem('bizFiles_v6');
    if (saved) singleFiles = JSON.parse(saved);
  } catch (e) {
    console.warn('LocalStorage 로드 실패:', e);
  }
}

function saveFilesToStorage() {
  try {
    localStorage.setItem('bizFiles_v6', JSON.stringify(singleFiles));
  } catch (e) {
    alert('브라우저 저장공간(LocalStorage) 한도를 초과했습니다. 첨부파일을 정리해 주세요.');
  }
}

function renderAttachmentRows(bizId) {
  document.getElementById('detail-attachments-container').innerHTML = FILE_CATS.map(cat => {
    const fKey = `${bizId}_${cat.key}`;
    const file = singleFiles[fKey];
    const isRegistered = Boolean(file && file.name);

    return `
      <div class="file-manage-row">
        <div class="file-info-col">
          <div style="display:flex; align-items:center; gap:8px;">
            <span class="box-badge ${cat.badge}">${cat.name}</span>
            <span style="font-weight:700; font-size:0.95rem;">${cat.name}</span>
          </div>
          <div style="font-size:0.83rem; margin-top:4px;">
            ${isRegistered 
              ? `<span style="color:#15803d; font-weight:700;">● 등록완료:</span> <b>${file.name}</b> <span style="color:#64748b;">(${file.size} \vert{}${file.date})</span>` 
              : `<span style="color:#94a3b8;">미등록 상태입니다. [파일 업로드] 버튼을 눌러 등록하세요.</span>`}
          </div>
        </div>
        <div class="file-actions-col">
          <input type="file" id="file-${fKey}" style="display:none;" onchange="handleFileUpload(event, '${bizId}', '${cat.key}')">
          <button class="btn-upload-file" onclick="document.getElementById('file-${fKey}').click()">📤 업로드</button>
          <button class="btn-download-file ${isRegistered ? '' : 'disabled'}" onclick="downloadFile('${bizId}', '${cat.key}')">💾 다운로드</button>
          ${isRegistered ? `<button class="btn-delete-file" onclick="deleteFile('${bizId}', '${cat.key}')">삭제</button>` : ''}
        </div>
      </div>
    `;
  }).join('');
}

function handleFileUpload(event, bizId, catKey) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(e) {
    singleFiles[`${bizId}_${catKey}`] = {
      name: file.name,
      size: file.size > 1048576 ? (file.size / 1048576).toFixed(1) + ' MB' : Math.round(file.size / 1024) + ' KB',
      date: new Date().toISOString().slice(0, 10),
      dataUrl: e.target.result
    };
    saveFilesToStorage();
    renderAttachmentRows(bizId);
    alert(`'${file.name}' 파일이 정상 등록되었습니다.`);
  };
  reader.readAsDataURL(file);
  event.target.value = '';
}

function downloadFile(bizId, catKey) {
  const file = singleFiles[`${bizId}_${catKey}`];
  if (!file) return alert('등록된 파일이 없습니다.');
  const a = document.createElement('a');
  a.href = file.dataUrl;
  a.download = file.name;
  a.click();
}

function deleteFile(bizId, catKey) {
  if (!confirm('등록된 파일을 삭제하시겠습니까?')) return;
  delete singleFiles[`${bizId}_${catKey}`];
  saveFilesToStorage();
  renderAttachmentRows(bizId);
}

function filterDashboard() {
  const q = document.getElementById('dashboardSearch').value.toLowerCase().trim();
  const cat = document.getElementById('categoryFilter').value;
  const allData = [...SUPPORT_TEAM_DATA, ...PROJECTS_DATA];
  let count = 0;

  document.querySelectorAll('.searchable-card').forEach((card, idx) => {
    const item = allData[idx];
    const text = card.innerText.toLowerCase();
    const match = (!q || text.includes(q)) && (cat === 'all' || item.team === cat);

    card.style.display = match ? 'block' : 'none';
    if (match) count++;
  });
  document.getElementById('searchResultCount').innerText = `총 ${count}개 사업 표시 중`;
}

function handleExcelUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      
      let parsedData = [];

      rows.forEach((cols, idx) => {
        if (!cols || cols.length === 0 || cols.every(cell => !cell)) return;

        const corpName = cols[1] || cols[0] || '-';
        const location = cols[2] || '-';
        const field = cols[3] || '-';
        const grievanceContent = cols[4] || '-';
        const consultContent = cols[5] || '-';
        const processContent = cols[6] || '-';
        const rawStatus = (cols[7] || cols[cols.length - 1] || '검토').toString().trim();

        parsedData.push({ corpName, location, field, grievanceContent, consultContent, processContent, rawStatus });
      });

      localStorage.setItem('grievanceData_v10', JSON.stringify(parsedData));
      renderGrievanceTable(parsedData);
      alert(`엑셀 업로드 및 갱신이 완료되었습니다. (총 ${parsedData.length}건)`);
    } catch (err) {
      console.error(err);
      alert('엑셀 파일을 읽는 중 오류가 발생했습니다. 파일 형식을 확인해 주세요.');
    }
  };
  reader.readAsArrayBuffer(file);
  event.target.value = '';
}

function renderGrievanceTable(dataList) {
  const tbody = document.getElementById('grievance-tbody');
  const summaryBox = document.getElementById('grievance-summary-content');
  tbody.innerHTML = '';

  if (!dataList || dataList.length === 0) {
    tbody.innerHTML = `<tr><td colspan="8" style="text-align:center; padding:30px; color:#94a3b8;">등록된 애로사항 데이터가 없습니다. 상단 [엑셀 업로드] 버튼으로 데이터를 등록하세요.</td></tr>`;
    document.getElementById('grievance-count').innerText = `총 접수: 0건`;
    summaryBox.innerHTML = `<div style="grid-column: 1 / -1; text-align:center; color:#94a3b8; padding:10px;">등록된 데이터가 없습니다.</div>`;
    return;
  }

  const targetStatuses = ['대기', '진행 중', '검토', '장기 검토', '불가', '완료'];
  let statusCounts = {};
  targetStatuses.forEach(st => statusCounts[st] = 0);

  dataList.forEach((item, idx) => {
    const cnt = idx + 1;
    let stKey = item.rawStatus;
    if (stKey.includes('장기')) stKey = '장기 검토';
    if (statusCounts[stKey] !== undefined) {
      statusCounts[stKey]++;
    } else {
      statusCounts[stKey] = (statusCounts[stKey] || 0) + 1;
    }

    let badgeStyle = 'background: #fef9c3; color: #a16207; padding: 4px 10px; border-radius: 6px; font-weight: 700; font-size: 0.78rem; display: inline-block;';

    if (item.rawStatus.includes('진행 중')) {
      badgeStyle = 'background: #dcfce7; color: #15803d; padding: 4px 10px; border-radius: 6px; font-weight: 700; font-size: 0.78rem; display: inline-block;';
    } else if (item.rawStatus.includes('대기')) {
      badgeStyle = 'background: #ffedd5; color: #c2410c; padding: 4px 10px; border-radius: 6px; font-weight: 700; font-size: 0.78rem; display: inline-block;';
    } else if (item.rawStatus.includes('검토') || item.rawStatus.includes('장기')) {
      badgeStyle = 'background: #fef9c3; color: #a16207; padding: 4px 10px; border-radius: 6px; font-weight: 700; font-size: 0.78rem; display: inline-block;';
    } else if (item.rawStatus.includes('불가') || item.rawStatus.includes('반려')) {
      badgeStyle = 'background: #fee2e2; color: #dc2626; padding: 4px 10px; border-radius: 6px; font-weight: 700; font-size: 0.78rem; display: inline-block;';
    } else if (item.rawStatus.includes('완료') || item.rawStatus.includes('조치')) {
      badgeStyle = 'background: #dbeafe; color: #1d4ed8; padding: 4px 10px; border-radius: 6px; font-weight: 700; font-size: 0.78rem; display: inline-block;';
    }

    tbody.innerHTML += `
      <tr>
        <td style="text-align:center; font-weight:700; color:var(--primary);">${cnt}</td>
        <td><b>${item.corpName}</b></td>
        <td>${item.location}</td>
        <td>${item.field}</td>
        <td>${item.grievanceContent}</td>
        <td>${item.consultContent}</td>
        <td>${item.processContent}</td>
        <td style="text-align:center;"><span style="${badgeStyle}">${item.rawStatus}</span></td>
      </tr>
    `;
  });

  document.getElementById('grievance-count').innerText = `총 접수: ${dataList.length}건`;

  summaryBox.innerHTML = Object.entries(statusCounts).map(([status, count]) => `
    <div class="content-box" style="text-align:center; padding:14px; background:#f8fafc; border:1px solid var(--border); border-radius:8px;">
      <div style="font-size:0.8rem; color:#64748b; font-weight:700; margin-bottom:4px;">${status}</div>
      <div style="font-size:1.4rem; font-weight:800; color:var(--primary);">${count}건</div>
    </div>
  `).join('');
}

function loadGrievanceData() {
  try {
    const saved = localStorage.getItem('grievanceData_v10');
    if (saved) {
      const dataList = JSON.parse(saved);
      renderGrievanceTable(dataList);
    } else {
      renderGrievanceTable(DEFAULT_GRIEVANCE_DATA);
    }
  } catch (e) {
    console.warn('애로사항 데이터 로드 실패:', e);
    renderGrievanceTable(DEFAULT_GRIEVANCE_DATA);
  }
}

function clearGrievanceData() {
  if (!confirm('기본 데이터(20건)로 복원하시겠습니까?')) return;
  localStorage.removeItem('grievanceData_v10');
  renderGrievanceTable(DEFAULT_GRIEVANCE_DATA);
  alert('기본 데이터로 복원되었습니다.');
}

function showManual(type) {
  const box = document.getElementById('manual-detail-box');
  box.classList.remove('hidden');
  const isEhojo = (type === 'ehojo');
  document.getElementById('manual-detail-title').innerText = isEhojo 
    ? '💰 e호조 회계지출 프로세스 매뉴얼' 
    : '📦 보탬e 보조금 교부정산 매뉴얼';
  document.getElementById('manual-detail-desc').innerHTML = isEhojo
    ? '세부사업계획 등록 ➔ 예산과목 배정 ➔ 지출품의 ➔ 지출원인행위 확정 ➔ 준공검사 및 대가지급'
    : '공모 및 사업신청 접수 ➔ 심의 및 교부결정 통보 ➔ 사업수행 및 집행내역 증빙 등록 ➔ 정산검증';
}

function initCharts() {
  const supportChartData = SUPPORT_TEAM_DATA.filter(p => p.finance.total > 0);
  const supLabels = supportChartData.map(p => p.shortTitle);
  const supTotals = supportChartData.map(p => p.finance.total);
  const supCities = supportChartData.map(p => p.finance.city);
  const supOthers = supportChartData.map(p => p.finance.total - p.finance.city);

  new Chart(document.getElementById('supportPieChart'), {
    type: 'doughnut',
    data: {
      labels: supLabels,
      datasets: [{ data: supTotals, backgroundColor: ['#0284c7', '#38bdf8', '#0369a1', '#0ea5e9', '#6366f1', '#8b5cf6', '#ec4899', '#f43f5e', '#10b981', '#14b8a6', '#f59e0b', '#fbbf24', '#64748b', '#475569'] }]
    },
    options: { 
      responsive: true, 
      maintainAspectRatio: false, 
      plugins: { legend: { position: 'bottom', labels: { boxWidth: 8, font: { size: 9 } } } } 
    }
  });

  new Chart(document.getElementById('supportBarChart'), {
    type: 'bar',
    data: {
      labels: supLabels,
      datasets: [
        { label: '시비(남양주)', data: supCities, backgroundColor: '#0284c7', borderRadius: 4 },
        { label: '도비·국비 매칭', data: supOthers, backgroundColor: '#cbd5e1', borderRadius: 4 }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 10 } } } },
      scales: { 
        x: { stacked: true },
        y: { stacked: true, beginAtZero: true, ticks: { callback: v => (v / 1000).toLocaleString() + '백만' } } 
      }
    }
  });

  const sosChartData = PROJECTS_DATA.filter(p => p.finance.total > 0);
  const labels = sosChartData.map(p => p.shortTitle);
  const totals = sosChartData.map(p => p.finance.total);
  const cities = sosChartData.map(p => p.finance.city);
  const others = sosChartData.map(p => p.finance.total - p.finance.city);

  new Chart(document.getElementById('budgetPieChart'), {
    type: 'doughnut',
    data: {
      labels: labels,
      datasets: [{ data: totals, backgroundColor: ['#2563eb', '#10b981', '#f59e0b', '#8b5cf6', '#06b6d4', '#ec4899', '#64748b', '#f97316'] }]
    },
    options: { 
      responsive: true, 
      maintainAspectRatio: false, 
      plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 10 } } } } 
    }
  });

  new Chart(document.getElementById('budgetBarChart'), {
    type: 'bar',
    data: {
      labels: labels,
      datasets: [
        { label: '시비(남양주)', data: cities, backgroundColor: '#1d4ed8', borderRadius: 4 },
        { label: '도비·국비 매칭', data: others, backgroundColor: '#94a3b8', borderRadius: 4 }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 10 } } } },
      scales: { 
        x: { stacked: true },
        y: { stacked: true, beginAtZero: true, ticks: { callback: v => (v / 1000).toLocaleString() + '백만' } } 
      }
    }
  });
}

window.addEventListener('DOMContentLoaded', () => {
  if (window.innerWidth <= 768) {
    document.querySelector('aside').classList.add('collapsed');
  }
  initApp();
});
