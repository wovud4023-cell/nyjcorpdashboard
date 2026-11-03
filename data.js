/**
 * ==========================================================================
 * [데이터 소스] 기업지원팀 14개 사업 및 기업SOS팀 10개 핵심 지원사업 통합 마스터 데이터
 * ==========================================================================
 */
const SUPPORT_TEAM_DATA = [
  {
    id: 'sup1', team: '지원팀', title: '1. 제품디자인 및 마케팅 지원', shortTitle: '제품디자인', badge: '총 29,900천원',
    finance: { total: 29900, city: 21400, prov: 8500, nat: 0, matching: '도비 간접매칭', agency: '경기도경제과학진흥원' },
    desc: '제품/시각/포장 디자인 개발 65% 이내 지원 (자부담 30% 이상)',
    plans: [{ title: '디자인 개발 지원', budget: '시비 21,400천원 / 도비 8,500천원', desc: '제품, 시각, 포장 디자인 개발 비용 65% 이내 지원\n• 자부담: 30% 이상' }]
  },
  {
    id: 'sup2', team: '지원팀', title: '2. 개발생산판로 맞춤형 지원', shortTitle: '개발생산판로', badge: '총 194,440천원',
    finance: { total: 194440, city: 119900, prov: 74540, nat: 0, matching: '도비 간접매칭', agency: '경기도경제과학진흥원' },
    desc: '창안개발, 제품생산, 판로개척 60% 이내 지원 (자부담 40% 이상)',
    plans: [{ title: '맞춤형 패키지 지원', budget: '시비 119,900천원 / 도비 74,540천원', desc: '창안개발, 제품생산, 판로개척 비용 60% 이내 지원\n• 자부담: 40% 이상' }]
  },
  {
    id: 'sup3', team: '지원팀', title: '3. 내일채움공제', shortTitle: '내일채움공제', badge: '총 180,000천원',
    finance: { total: 180000, city: 180000, prov: 0, nat: 0, matching: '시비 100%', agency: '중소벤처기업진흥공단' },
    desc: '공제부금 기업부담금 일부 지원 (기업 14만원, 근로자 10만원 시 10만원)',
    plans: [{ title: '핵심인력 성과보상 공제', budget: '시비 180,000천원', desc: '공제부금 기업부담금 일부 지원' }]
  },
  {
    id: 'sup4', team: '지원팀', title: '4. 가구산업 홍보비 지원', shortTitle: '가구산업홍보', badge: '총 60,022천원',
    finance: { total: 60022, city: 36000, prov: 18000, nat: 0, matching: '도비 직접매칭', agency: '마석가구공단연합회' },
    desc: 'SBS 라디오, 버스(93, 1330-4번) 광고 지원 (자부담 10%)',
    plans: [{ title: '매체 광고 지원', budget: '시비 36,000천원 / 도비 18,000천원', desc: 'SBS 라디오, 버스 노선 광고 집행\n• 자부담: 10%' }]
  },
  {
    id: 'sup5', team: '지원팀', title: '5. 기업 맞춤형 마케팅 지원', shortTitle: '기업맞춤마케팅', badge: '총 70,000천원',
    finance: { total: 70000, city: 70000, prov: 0, nat: 0, matching: '시비 100%', agency: '경기대진테크노파크' },
    desc: '마케팅 비용 70% 이내 지원 (자부담 30% 이상)',
    plans: [{ title: '맞춤형 마케팅', budget: '시비 70,000천원', desc: '마케팅 소요 비용 70% 이내 지원\n• 자부담: 30% 이상' }]
  },
  {
    id: 'sup6', team: '지원팀', title: '6. 매출채권보험', shortTitle: '매출채권보험', badge: '총 96,000천원',
    finance: { total: 96000, city: 96000, prov: 0, nat: 0, matching: '시비 100%', agency: '신용보증기금' },
    desc: '보험료 20% 지원 (2,000천원 한도, 자부담 10% 이상)',
    plans: [{ title: '보험료 지원', budget: '시비 96,000천원', desc: '매출채권보험료 20% 지원 (한도 2,000천원)\n• 자부담: 10% 이상' }]
  },
  {
    id: 'sup7', team: '지원팀', title: '7. 남양주미래산업 성장전략 정책포럼', shortTitle: '미래산업포럼', badge: '총 22,000천원',
    finance: { total: 22000, city: 22000, prov: 0, nat: 0, matching: '시비 100%', agency: '브레인파크' },
    desc: '남양주 산업 진단 및 미래 설계 (신규 사업)',
    plans: [{ title: '정책포럼 개최', budget: '시비 22,000천원', desc: '남양주 관내 산업 진단 및 미래 성장 전략 포럼 운영' }]
  },
  {
    id: 'sup8', team: '지원팀', title: '8. 시장개척단', shortTitle: '시장개척단', badge: '총 140,000천원',
    finance: { total: 140000, city: 140000, prov: 0, nat: 0, matching: '시비 100%', agency: '경기도경제과학진흥원' },
    desc: '왕복항공료 50%, 상담장 임차, 통역(1개사당 1인), 바이어매칭 지원',
    plans: [{ title: '해외 시장개척단 파견', budget: '시비 140,000천원', desc: '항공료, 상담장, 통역 및 바이어 매칭 패키지 지원' }]
  },
  {
    id: 'sup9', team: '지원팀', title: '9. 해외전시회 (단체관)', shortTitle: '해외전시회', badge: '총 117,000천원',
    finance: { total: 117000, city: 117000, prov: 0, nat: 0, matching: '시비 100%', agency: '경기도경제과학진흥원' },
    desc: '부스비·장치비 80%, 편도운송료 100%, 통역(1개사당 1인) 지원',
    plans: [{ title: '해외전시회 단체관 참가', budget: '시비 117,000천원', desc: '전시회 부스비, 장치비, 운송료, 통역 지원' }]
  },
  {
    id: 'sup10', team: '지원팀', title: '10. 수출물류비', shortTitle: '수출물류비', badge: '총 42,000천원',
    finance: { total: 42000, city: 42000, prov: 0, nat: 0, matching: '시비 100%', agency: '경기도경제과학진흥원' },
    desc: '수출거래에 따른 물류비 80% 지원',
    plans: [{ title: '수출물류비 지원', budget: '시비 42,000천원', desc: '수출 물류 소요 비용의 80% 지원' }]
  },
  {
    id: 'sup11', team: '지원팀', title: '11. 수출지원센터', shortTitle: '수출지원센터', badge: '총 20,000천원',
    finance: { total: 20000, city: 20000, prov: 0, nat: 0, matching: '시비 100%', agency: '경기동부상공회의소' },
    desc: '수출관련 컨설팅 지원',
    plans: [{ title: '수출 컨설팅 운영', budget: '시비 20,000천원', desc: '관내 기업 대상 수출 실무 및 애로사항 컨설팅 제공' }]
  },
  {
    id: 'sup12', team: '지원팀', title: '12. 국내전시회 지원', shortTitle: '국내전시회', badge: '총 81,000천원',
    finance: { total: 81000, city: 40500, prov: 40500, nat: 0, matching: '도비 직접매칭', agency: '경기도경제과학진흥원' },
    desc: '부스비·장치비 100% 지원',
    plans: [{ title: '국내전시회 참가 지원', budget: '시비 40,500천원 / 도비 40,500천원', desc: '국내 전시회 참가 기업의 부스비 및 장치비 전액 지원' }]
  },
  {
    id: 'sup13', team: '지원팀', title: '13. 기업인 역량강화 교육지원', shortTitle: '기업인역량강화', badge: '총 70,000천원',
    finance: { total: 70000, city: 70000, prov: 0, nat: 0, matching: '시비 100%', agency: '경기동부상공회의소' },
    desc: '조찬강연 포럼, CEO아카데미 과정, 청년기업인 미래포럼 운영',
    plans: [{ title: '교육 및 포럼 지원', budget: '시비 70,000천원', desc: '조찬 포럼, CEO 아카데미, 청년기업인 미래포럼 개최' }]
  },
  {
    id: 'sup14', team: '지원팀', title: '14. 호우피해 복구지원 (중소기업)', shortTitle: '호우피해복구', badge: '총 12,000천원',
    finance: { total: 12000, city: 3600, prov: 2400, nat: 6000, matching: '국·도비 매칭', agency: '남양주시 (직접수행)' },
    desc: '호우피해 중소기업에 재난지원금 지급 (신규 사업)',
    plans: [{ title: '재난지원금 지원', budget: '국비 6,000천원 / 도비 2,400천원 / 시비 3,600천원', desc: '호우피해를 입은 관내 중소기업 대상 재난지원금 지급' }]
  }
];

const PROJECTS_DATA = [
  {
    id: 'biz1', team: 'SOS팀', title: '1. 맞춤형기업애로민원해결사업', shortTitle: '맞춤형애로', category: '인프라', badge: '총 2.5억',
    finance: { total: 250000, city: 250000, prov: 0, nat: 0, matching: '해당없음', agency: '남양주시 (직접수행)' },
    keywords: '도로 상하수도 긴급 기동보수 단가계약 기반시설 직접수행',
    desc: '기업밀집지역 위험도로·상하수도·안전시설 정비 및 긴급 기동보수',
    plans: [
      { title: '민원처리 연간 단가계약', budget: '180,000천원', desc: '기업밀집지역 기반시설(도로, 관로 등) 재해복구 및 긴급 유지보수' },
      { title: '도로재포장 및 기반시설 보수', budget: '70,000천원', desc: '노후도로 아스콘 재포장, 관로설치 및 안전시설 정비' }
    ],
    history: {
      headers: ['사업명', '기간', '내용', '사업비(천원)', '상태'],
      rows: [
        ['상반기 민원처리 단가계약', '2026.2~7', '기반시설 긴급정비', '81,283', '완료'],
        ['창현리 노후도로 포장공사', '2026.9~11', '도로 노면 포장 정비', '42,810', '진행중']
      ]
    }
  },
  {
    id: 'biz2', team: 'SOS팀', title: '2. 지식재산창출지원사업', shortTitle: '지식재산', category: '경영지원', badge: '총 0.85억',
    finance: { total: 85494, city: 60000, prov: 0, nat: 25494, matching: '간접매칭', agency: '경기대진테크노파크 (위탁)' },
    keywords: '특허 디자인 브랜드 출원비 특허맵 경기대진테크노파크 국비',
    desc: '관내 중소기업 맞춤형 지식재산(특허, 디자인, 브랜드, 출원비) 창출 지원',
    plans: [
      { title: '특허', budget: '9,000천원 이내', desc: '특허맵, 특허기술 홍보영상 제작 지원' },
      { title: '디자인', budget: '16,000천원 이내', desc: '디자인맵(일반), 제품디자인개발, 제품디자인목업 등 지원' },
      { title: '브랜드', budget: '14,000천원 이내', desc: '신규브랜드, 리뉴얼브랜드 개발 지원' },
      { title: '국내출원비용', budget: '2,100천원 이내', desc: '국내 출원 시 소요되는 대리인 비용, 출연료 지원' },
      { title: '해외출원비용', budget: '2,200천원 이내', desc: '해외 출원 시 소요되는 대리인 비용, 번역료, 국제출원 수수료 지원' }
    ]
  },
  {
    id: 'biz3', team: 'SOS팀', title: '3. 기술닥터사업', shortTitle: '기술닥터', category: '기술R&D', badge: '총 1.44억',
    finance: { total: 143900, city: 100730, prov: 43170, nat: 0, matching: '직접매칭', agency: '경기테크노파크 (위탁)' },
    keywords: '기술닥터 현장애로 중기애로 상용화 단계별검증 원스톱 경기테크노파크 도비',
    desc: '산·학·연 전문가(기술닥터)를 활용한 맞춤형 기술지원 및 단계별 상용화 지원',
    plans: [
      { title: '1단계: 현장애로기술지원', budget: '11건 (무료)', desc: '기업 현장 10회 방문 1:1 밀착 기술애로 해결' },
      { title: '2단계: 중기애로기술지원', budget: '4건 (최대 2천만)', desc: '시제품 제작, 공정개선 등 심화 성과물 도출 (자부담 20%)' },
      { title: '3단계: 상용화지원', budget: '30,000천원 이내', desc: '개발 기술의 사업화 및 상용화 촉진을 위한 심화 지원' },
      { title: '단계별검증지원', budget: '6,000천원 이내', desc: '기술의 성능 및 신뢰성 검증 지원' }
    ]
  },
  {
    id: 'biz4', team: 'SOS팀', title: '4. 국가기술거래플랫폼사업', shortTitle: '국가기술거래', category: '기술R&D', badge: '총 1.76억',
    finance: { total: 176500, city: 176500, prov: 0, nat: 0, matching: '해당없음', agency: '경기대진테크노파크 (위탁)' },
    keywords: '공공기술 이전 사업화 기술코칭 경기대진테크노파크',
    desc: '첨단전략산업 중심 우수 공공기술 이전 및 사업화, 전문가 기술코칭 지원',
    plans: [
      { title: '기술이전 및 사업화 지원', budget: '6개사 (최대 2.2천만)', desc: '수요기술 매칭, 기술료 및 시제품 상용화 패키지 지원' },
      { title: '기술코칭', budget: '최대 1,500천원 (3개사)', desc: '전문가 매칭을 통한 기술 애로사항 및 사업화 코칭 지원' }
    ]
  },
  {
    id: 'biz5', team: 'SOS팀', title: '5. 기업환경 개선사업', shortTitle: '기업환경개선', category: '인프라', badge: '총 5.62억',
    finance: { total: 561545, city: 367182, prov: 194363, nat: 0, matching: '직접매칭', agency: '남양주시 (직접수행)' },
    keywords: '밀집지역 기반시설 노동복지 작업환경 소방안전 직접매칭 직접수행',
    desc: '중소기업 밀집지역 기반시설 정비 및 개별기업 노동·작업환경, 소방안전 시설 개선',
    plans: [
      { title: '기반시설 정비 및 작업환경 개선', budget: '한도 2억 / 2천만원', desc: '도로포장, 상하수도, 화장실, 환기집진시설 정비' }
    ]
  },
  {
    id: 'biz6', team: 'SOS팀', title: '6. 스마트공장 구축 및 컨설팅', shortTitle: '스마트공장', category: '기술R&D', badge: '총 5.30억',
    finance: { total: 530400, city: 371280, prov: 159120, nat: 0, matching: '직접매칭', agency: '경기테크노파크 (위탁)' },
    keywords: '스마트공장 솔루션 자동화 설비 현장컨설팅 직접매칭 경기테크노파크',
    desc: '관내 중소 제조기업 기초단계 스마트공장 솔루션·설비 구축 및 현장 컨설팅',
    plans: [
      { title: '스마트공장 구축 지원', budget: '기업당 50,000천원', desc: '기초단계 솔루션 및 연계 장비 구축 (총 사업비 70% 지원)' }
    ]
  },
  {
    id: 'biz7', team: 'SOS팀', title: '7. 경기도 기술개발사업', shortTitle: '기술개발(R&D)', category: '기술R&D', badge: '총 4.5억',
    finance: { total: 450000, city: 150000, prov: 300000, nat: 0, matching: '간접매칭', agency: '경기도경제과학진흥원 (위탁)' },
    keywords: 'R&D 첨단기술 탄소중립 기술개발 연구전담부서 경과원 경기도경제과학진흥원',
    desc: '기업수요 중심의 첨단기술 및 탄소중립 산업기술 연구개발(R&D) 지원',
    plans: [
      { title: 'R&D 기술개발 과제 지원', budget: '과제당 최대 1.5억', desc: '연구개발 인건비, 시제품 제작비, 연구장비비 지원 (3개사)' }
    ]
  },
  {
    id: 'biz8', team: 'SOS팀', title: '8. 경복대 창업보육센터 운영지원', shortTitle: '창업보육', category: '경영지원', badge: '총 0.33억',
    finance: { total: 33333, city: 23333, prov: 10000, nat: 0, matching: '직접매칭', agency: '남양주시 (직접수행)' },
    keywords: '경복대 창업보육 입주기업 시제품 인증 마케팅 직접매칭 직접수행',
    desc: '경복대학교 창업보육센터 입주기업 대상 국내외 인증, 시제품, 홍보마케팅 지원',
    plans: [
      { title: '입주기업 사업화 지원', budget: '총 33,333천원', desc: '인증, 시제품, 홍보마케팅, 특허출원 패키지 지원' }
    ]
  },
  {
    id: 'biz9', team: 'SOS팀', title: '9. 지식나눔 기업사랑방', shortTitle: '기업사랑방', category: '경영지원', badge: '비예산',
    finance: { total: 0, city: 0, prov: 0, nat: 0, matching: '해당없음', agency: '남양주시 (직접수행)' },
    keywords: '지식나눔 기업사랑방 전문가 상담(법률, 회계, 세무 등 10개 분야) 비예산 무료상담',
    desc: '관내 중소기업 맞춤형 지식공유(법률, 회계, 세무, 창업지도, 가치평가 등 10개 분야)',
    plans: [
      { title: '전문가 1:1 상담 데이', budget: '비예산(무료)', desc: '10개 분야 상담 : 법률, 특허, 노무, 경영, 회계, 세무, 무역, 법무, 감정평가, 창업지도' }
    ]
  },
  {
    id: 'biz10', team: 'SOS팀', title: '10. 청년기업 인증제도', shortTitle: '청년기업', category: '경영지원', badge: '비예산 (99개사)',
    finance: { total: 0, city: 0, prov: 0, nat: 0, matching: '해당없음', agency: '남양주시 (직접수행)' },
    keywords: '청년기업 인증 공공구매 수의계약 가점 육성자금 비예산 99개사 인센티브 9건',
    desc: '관내 만 19세~39세 청년 창업자 발굴 및 공공구매·수의계약·자금 우대 지원',
    plans: [
      { 
        title: '청년기업 인증 현황 (총 99개사)', 
        budget: '인증완료', 
        desc: '• 요식업: 3개사\n• 도매 및 소매업: 17개사\n• 제조업: 42개사\n• 정보통신업: 4개사\n• 기타: 44개사' 
      },
      { 
        title: '청년기업 전용 인센티브 (총 9건)', 
        budget: '우대지원', 
        desc: '1. 소상공인 경영환경 개선사업 가점 1점\n2. 국가기술거래플랫폼 지원사업 가점 2점\n3. 청년기업 제품 우수구매 권장\n4. 글로벌 시장개척단 / 해외전시회 가점 5점\n5. 경복대 창업보육센터 입주 가점 5점\n6. 청년 소상공인 특례보증 지원 출연금 우대\n7. 중소기업 육성자금 지원사업 금리우대\n8. 기업 맞춤형 마케팅 지원사업 가점 3점\n9. 위생업소 환경개선 지원사업 가점 1점' 
      }
    ]
  }
];

const DEFAULT_GRIEVANCE_DATA = [
  { corpName: "삼진씨앤에프", location: "화도읍 차산리 24-12일원", field: "주변 인프라", grievanceContent: "도로파손", consultContent: "도로재포장", processContent: "토지사용동의서 접수 어려움상태", rawStatus: "대기" },
  { corpName: "쌍용스텐", location: "화도읍 차산리 19-4 일원", field: "주변 인프라", grievanceContent: "대형차량 회전 불가", consultContent: "구거 복개 및 재포장", processContent: "주민숙원사업 통한 해결 예정(서류 전달 완료)-지역경제과\n2027년 본예산 신청완료-화도읍 도시건축과", rawStatus: "진행 중" },
  { corpName: "경진씨앤씨", location: "녹촌리 83-7 일원", field: "주변 인프라", grievanceContent: "도로통행 불편", consultContent: "도로통행 불편 해소", processContent: "주차금지구역 선정 계획 중(도로관리과)", rawStatus: "검토" },
  { corpName: "진벌리", location: "진접읍 진벌리 185-1", field: "주변 인프라", grievanceContent: "도로통행 불편", consultContent: "교각확장", processContent: "주민숙원사업 서류 전달 완료-진벌1리 이장님\n진접읍 도시건축과-22381(2026.8.18.)호-[기업애로 사항으로 지역경제과 공문 시행]\n지역경제과-25378(2026.8.28.)호-[주민숙원사업 검토 및 선정 결과 공유 요청]", rawStatus: "검토" },
  { corpName: "철마기업인회", location: "진접읍 진벌리 252-10 일원", field: "주변 인프라", grievanceContent: "도시계획도로 지정 및 개설공사 건의", consultContent: "도시계획도로 지정 및 개설공사 건의", processContent: "시의 재정 여건, 시급성, 교통현황 등 종합적으로 고려하여 중 장기검토 필요-도로건설과", rawStatus: "장기 검토" },
  { corpName: "녹촌기업인회", location: "녹촌리 303-17~창현리 577-7", field: "주변 인프라", grievanceContent: "교통정체 예상", consultContent: "녹촌IC인근(농어촌도로102호선) 확장건의", processContent: "실시설계 진행 중, 사업비 확보하여 보상절차 진행 예정-도로건설과", rawStatus: "진행 중" },
  { corpName: "삭다니기업인회", location: "양지리 543-60 일원", field: "주변 인프라", grievanceContent: "도로통행 불편", consultContent: "가로등 설치", processContent: "사유지 설치 불가 / 한전, 소관부서 협의 지속 필요-도로시설관리과", rawStatus: "불가" },
  { corpName: "용정기업인회", location: "용정기업인회 입구 경사로", field: "주변 인프라", grievanceContent: "도로통행 불편", consultContent: "열선 설치", processContent: "장기적 검토 필요(예산, 사업성 등) - 진건읍 도시건축과\n염사분사기 설치로 검토 중", rawStatus: "검토" },
  { corpName: "동영화섬", location: "화도읍 월산리 327번지 일원", field: "주변 인프라", grievanceContent: "상수도 공급 요청", consultContent: "상수도 공급 요청", processContent: "개별 급수는 신청인 부담 원칙으로 불가(수도과)", rawStatus: "불가" },
  { corpName: "차산기업인회", location: "화도읍 재재기로71번길 11-1", field: "주변 인프라", grievanceContent: "도로통행 불편", consultContent: "도로 재포장", processContent: "도로 꺼짐에 따른 재포장 요청 >> 구거 지역 포함되어 화도읍 도시건축과 조치", rawStatus: "완료" },
  { corpName: "바람골길기업인회", location: "진접읍 금곡전기", field: "주변 인프라", grievanceContent: "수도연결 등", consultContent: "도로 포장 및 수도 연결", processContent: "진접 차량기지 등 건설로 인하여 지역 식수 오염 문제가 발생. 도로건설과, 수도과 협의하여 도시계획도로 건설과 동시에 수도 설치 중. 추가 구간 예산 확보 예정", rawStatus: "진행 중" },
  { corpName: "삼광비닐", location: "차산리 238", field: "안전", grievanceContent: "보강토옹벽 배부름 현상", consultContent: "안전 진단 및 조치", processContent: "옹벽 배면부 상수관 파열로 인한 관로 보수 이후 보강토옹벽 1단 하부 이격현상과 옹벽 전반 밀림\n시민안전관 안전관리자문단 점검 실시, 계측 관리(도로관리사업소_지표투과레이더) 요청", rawStatus: "진행 중" },
  { corpName: "중산시스템", location: "진접읍 부마로 227", field: "안전", grievanceContent: "토사 유출", consultContent: "토사 유출 방지 요청", processContent: "7.30. 토지주동의완료 / 실시설계용역: 8.~12. / 사방공사 시행 2027. 2. ~ 5. [산림녹지과]", rawStatus: "진행 중" },
  { corpName: "스위트컵", location: "진접읍 팔야로 149", field: "주변 인프라", grievanceContent: "차량 진입 불가", consultContent: "교량 폭 확장", processContent: "교각 폭이 좁아 대형차량 진출입이 어려움. 관련부서 협의 완료 및 자비로 교량 확장 공사 진행 중", rawStatus: "진행 중" },
  { corpName: "철마기업인회", location: "진벌리 281-8", field: "주변 인프라", grievanceContent: "도로통행 불편", consultContent: "도시계획도로 지정 및 인도설치건의", processContent: "해당 구역은 사유지이며, 통행불편에 따른 도시계획도로 지정 및 인도설치 필요-도로건설과 장기 검토", rawStatus: "장기 검토" },
  { corpName: "하늘농가(진송,용정기업인회)", location: "용정리 229-12외 11필지\n송능리 70-10외 2필지", field: "주변 인프라", grievanceContent: "교통정체 예상", consultContent: "주차장 조성", processContent: "공원녹지->주차장용지 사업 시행자(경기주택도시공사)의 반영 어려움 회신\n이에 따라 지속협의계획(미래도시과)", rawStatus: "불가" },
  { corpName: "차산기업인회", location: "차산리 187-1", field: "주변 인프라", grievanceContent: "도로통행 불편", consultContent: "2차선 확장 요청", processContent: "인도교 설치 공사 수행 중. 교량 확장 검토 결과 비용대비 효과가 낮아 추진 불가-화도읍사무소", rawStatus: "불가" },
  { corpName: "가곡공장단지기업인회", location: "가곡공장단지 기업 도로 일원", field: "안전", grievanceContent: "도로통행 불편", consultContent: "염화칼슘 지원 요청", processContent: "협의 완료", rawStatus: "완료" },
  { corpName: "성신미공사", location: "오남리 108-13 일원", field: "주변 인프라", grievanceContent: "도로통행 불편", consultContent: "도시계획도로 지정 및 개설공사 건의", processContent: "지역경제과-29589(2026.10.14.)호 도시계획도로 지정 및 개설요청 공문 발송\n도로건설과-20409(2026.10.26.)호 주 도로(오남리 102-2번지 일원) 공사 완료에 따라 불가.", rawStatus: "불가" },
  { corpName: "동부상공회의소", location: "진접읍 내각리 178-4 일원", field: "주변 인프라", grievanceContent: "수압 저하", consultContent: "수도관 가압펌프 설치 요청", processContent: "지역경제과-29444(2026.10.13.)호 수압 실측 및 가압펌프 설치 등 개선방안 공문 발송\n수도과-19945(2026.10.14.)호 수압 측정 결과 정상. 보완 필요성 낮음", rawStatus: "불가" }
];
