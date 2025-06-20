// Google Apps Script 백엔드 코드

// 스프레드시트 ID와 시트 이름 설정
const SPREADSHEET_ID = '1Ykzi-v6zPW4hgUaRYsSloStlWPQN_HgDoPvMVcGmdmc';
const SHEETS = {
  users: '사용자',
  teams: '팀',
  channels: '채널',
  evaluations: '월별평가',
  members: '담당자'
};

// 스프레드시트 접근 헬퍼
function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (error) {
    console.error('스프레드시트 접근 오류:', error);
    try {
      return SpreadsheetApp.getActiveSpreadsheet();
    } catch (e) {
      throw new Error('스프레드시트에 접근할 수 없습니다. 권한을 확인해주세요.');
    }
  }
}

// 웹앱 진입점
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('성과 평가 시스템')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 스프레드시트 데이터 가져오기 헬퍼 함수
function getSheetData(sheetName) {
  try {
    const spreadsheet = getSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      console.error(`시트를 찾을 수 없습니다: ${sheetName}`);
      return [];
    }
    
    const range = sheet.getDataRange();
    if (range.getNumRows() <= 1) {
      return [];
    }
    
    const data = range.getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    return rows.map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });
  } catch (error) {
    console.error(`getSheetData 오류 (${sheetName}):`, error);
    return [];
  }
}

// 현재 사용자 정보 가져오기 (권한 정보 포함)
function getCurrentUser(email) {
  const users = getSheetData(SHEETS.users);
  const user = users.find(user => user['사용자ID'] === email);
  
  if (user) {
    user['평가권한'] = [];
    
    if (user['1차평가권한'] === 'Y') {
      user['평가권한'].push('1차');
    }
    if (user['2차평가권한'] === 'Y') {
      user['평가권한'].push('2차');
    }
    if (user['3차평가권한'] === 'Y') {
      user['평가권한'].push('3차');
    }
    
    if (user['소속팀ID'] && user['소속팀ID'].includes(',')) {
      user['소속팀목록'] = user['소속팀ID'].split(',').map(t => t.trim());
    } else {
      user['소속팀목록'] = user['소속팀ID'] ? [user['소속팀ID']] : [];
    }
  }
  
  return user || null;
}

// 역할별 데이터 가져오기
function getDataByRole(userId, role, evaluationMonth) {
  console.log('getDataByRole 호출:', {userId, role, evaluationMonth});
  
  try {
    // 데이터 로드
    const users = getSheetData(SHEETS.users);
    const teams = getSheetData(SHEETS.teams);
    const channels = getSheetData(SHEETS.channels);
    const evaluations = getSheetData(SHEETS.evaluations);
    const members = getSheetData(SHEETS.members);
    
    console.log('로드된 원시 데이터:', {
      users: users.length,
      teams: teams.length,
      channels: channels.length,
      evaluations: evaluations.length,
      members: members.length
    });
    
    // 결과 객체 초기화
    const result = {
      teams: [],
      channels: [],
      evaluations: [],
      users: [],
      members: []
    };
    
    // 평가월 필터링
    const monthEvaluations = evaluationMonth ? 
      evaluations.filter(e => e['평가월'] === evaluationMonth) : 
      evaluations;
    
    if (role === '팀장') {
      const currentUser = users.find(u => u['사용자ID'] === userId);
      console.log('현재 사용자 찾기:', currentUser);
      
      if (currentUser && currentUser['소속팀ID']) {
        const userTeams = currentUser['소속팀ID'].split(',').map(t => t.trim());
        console.log('팀장의 소속팀들:', userTeams);
        
        result.teams = teams.filter(t => userTeams.includes(t['팀ID']));
        result.channels = channels.filter(c => userTeams.includes(c['소속팀ID']) && c['활성상태'] !== false);
        result.users = users.filter(u => {
          const uTeams = u['소속팀ID'] ? u['소속팀ID'].split(',').map(t => t.trim()) : [];
          return uTeams.some(t => userTeams.includes(t));
        });
        result.members = members.filter(m => userTeams.includes(m['소속팀ID']));
        result.evaluations = monthEvaluations.filter(e => {
          const channelTeam = channels.find(c => c['채널ID'] === e['채널ID'])?.['소속팀ID'];
          return userTeams.includes(channelTeam);
        });
      }
    } else {
      // 관리자와 최종관리자는 모든 데이터
      result.teams = teams;
      result.channels = channels.filter(c => c['활성상태'] !== false);
      result.users = users;
      result.members = members;
      result.evaluations = monthEvaluations;
    }
    
    console.log('필터링된 결과:', {
      teams: result.teams.length,
      channels: result.channels.length,
      users: result.users.length,
      members: result.members.length,
      evaluations: result.evaluations.length
    });
    
    // 중요: 결과 반환
    return result;
    
  } catch (error) {
    console.error('getDataByRole 오류:', error);
    console.error('오류 상세:', error.stack);
    
    // 오류 발생시에도 빈 객체 반환
    return {
      teams: [],
      channels: [],
      evaluations: [],
      users: [],
      members: []
    };
  }
}

// addMember 함수 수정 (현재 시트 구조에 맞게)
function addMember(memberData) {
  try {
    const spreadsheet = getSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEETS.members);
    
    if (!sheet) {
      throw new Error('담당자 시트를 찾을 수 없습니다.');
    }
    
    // 현재 시트의 실제 헤더
    const headers = ['담당자ID', '이름', '역할', '소속팀ID', '활성상태', '생성일시'];
    
    // 헤더가 없으면 생성
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
    }
    
    // 기존 담당자 ID들을 가져와서 최대값 찾기
    const data = sheet.getDataRange().getValues();
    let maxNum = 0;
    
    for (let i = 1; i < data.length; i++) {
      const memberId = data[i][0];
      if (memberId && memberId.toString().startsWith('MEM')) {  // MEM으로 시작
        const num = parseInt(memberId.toString().substring(3));
        if (!isNaN(num) && num > maxNum) {
          maxNum = num;
        }
      }
    }
    
    const newId = 'MEM' + String(maxNum + 1).padStart(3, '0');
    const now = new Date();
    
    sheet.appendRow([
      newId, 
      memberData.name, 
      memberData.role, 
      memberData.teamId,
      'Y',  // 활성상태
      now   // 생성일시
    ]);
    
    return { success: true, memberId: newId };
  } catch (error) {
    console.error('addMember 오류:', error);
    return { success: false, message: error.toString() };
  }
}
// 담당자 추가 함수 - 새로 추가
function addMember(memberData) {
  try {
    const spreadsheet = getSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEETS.members);
    
    if (!sheet) {
      throw new Error('담당자 시트를 찾을 수 없습니다.');
    }
    
    // 헤더가 없으면 생성
    if (sheet.getLastRow() === 0) {
      const headers = ['담당자ID', '이름', '역할', '소속팀ID', '소속채널ID'];
      sheet.appendRow(headers);
    }
    
    // 기존 담당자 ID들을 가져와서 최대값 찾기
    const data = sheet.getDataRange().getValues();
    let maxNum = 0;
    
    for (let i = 1; i < data.length; i++) {
      const memberId = data[i][0];
      if (memberId && memberId.toString().startsWith('M')) {
        const num = parseInt(memberId.toString().substring(1));
        if (!isNaN(num) && num > maxNum) {
          maxNum = num;
        }
      }
    }
    
    const newId = 'M' + String(maxNum + 1).padStart(3, '0');
    sheet.appendRow([
      newId, 
      memberData.name, 
      memberData.role, 
      memberData.teamId,
      memberData.channelId || ''
    ]);
    
    return { success: true, memberId: newId };
  } catch (error) {
    console.error('addMember 오류:', error);
    return { success: false, message: error.toString() };
  }
}
// 사용자 추가 함수 - 시스템 관리자만
function addUser(userData) {
  try {
    const spreadsheet = getSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEETS.users);
    
    if (!sheet) {
      throw new Error('사용자 시트를 찾을 수 없습니다.');
    }
    
    // 헤더가 없으면 생성
    if (sheet.getLastRow() === 0) {
      const headers = ['사용자ID', '사용자계정', '이름', '역할', '소속팀ID', '1차평가권한', '2차평가권한', '3차평가권한'];
      sheet.appendRow(headers);
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 권한 설정
    const permissions = {
      '1차평가권한': userData.permissions?.perm1 !== undefined ? 
        (userData.permissions.perm1 ? 'Y' : 'N') :
        (userData.role === '팀장' ? 'Y' : 'N'),
      '2차평가권한': userData.permissions?.perm2 !== undefined ? 
        (userData.permissions.perm2 ? 'Y' : 'N') :
        (userData.role === '관리자' ? 'Y' : 'N'),
      '3차평가권한': userData.permissions?.perm3 !== undefined ? 
        (userData.permissions.perm3 ? 'Y' : 'N') :
        (userData.role === '최종관리자' ? 'Y' : 'N')
    };
    
    const teamIds = Array.isArray(userData.teamId) ? 
      userData.teamId.join(',') : 
      userData.teamId;
    
    const rowData = headers.map(header => {
      switch(header) {
        case '사용자ID': return userData.userId;
        case '사용자계정': return userData.userId;  // 시스템 관리자는 이메일이 ID
        case '이름': return userData.name;
        case '역할': return userData.role;
        case '소속팀ID': return teamIds;
        case '1차평가권한': return permissions['1차평가권한'];
        case '2차평가권한': return permissions['2차평가권한'];
        case '3차평가권한': return permissions['3차평가권한'];
        default: return '';
      }
    });
    
    sheet.appendRow(rowData);
    return { success: true, userId: userData.userId };
  } catch (error) {
    console.error('addUser 오류:', error);
    return { success: false, message: error.toString() };
  }
}

// 평가 진행 상태 조회 함수 - 새로 추가
function getEvaluationStatus(evaluationMonth) {
  try {
    const evaluations = getSheetData(SHEETS.evaluations);
    const channels = getSheetData(SHEETS.channels);
    const teams = getSheetData(SHEETS.teams);
    
    const monthEvaluations = evaluations.filter(e => e['평가월'] === evaluationMonth);
    
    const statusByChannel = {};
    const statusByTeam = {};
    
    // 채널별 진행상태 계산
    channels.forEach(channel => {
      const channelEvals = monthEvaluations.filter(e => e['채널ID'] === channel['채널ID']);
      
      if (channelEvals.length === 0) {
        statusByChannel[channel['채널ID']] = '미시작';
      } else {
        const has1st = channelEvals.some(e => e['1차평가상태'] === '제출완료');
        const has2nd = channelEvals.some(e => e['2차평가상태'] === '제출완료');
        const has3rd = channelEvals.some(e => e['최종평가상태'] === '제출완료');
        
        if (has3rd) statusByChannel[channel['채널ID']] = '완료';
        else if (has2nd) statusByChannel[channel['채널ID']] = '3차대기';
        else if (has1st) statusByChannel[channel['채널ID']] = '2차대기';
        else statusByChannel[channel['채널ID']] = '1차진행중';
      }
    });
    
    // 팀별 진행상태 계산
    teams.forEach(team => {
      const teamChannels = channels.filter(c => c['소속팀ID'] === team['팀ID']);
      const statuses = teamChannels.map(c => statusByChannel[c['채널ID']] || '미시작');
      
      if (statuses.every(s => s === '완료')) statusByTeam[team['팀ID']] = '완료';
      else if (statuses.some(s => s === '3차대기')) statusByTeam[team['팀ID']] = '3차대기';
      else if (statuses.some(s => s === '2차대기')) statusByTeam[team['팀ID']] = '2차대기';
      else if (statuses.some(s => s === '1차진행중')) statusByTeam[team['팀ID']] = '1차진행중';
      else statusByTeam[team['팀ID']] = '미시작';
    });
    
    return { channelStatus: statusByChannel, teamStatus: statusByTeam };
  } catch (error) {
    console.error('getEvaluationStatus 오류:', error);
    return { channelStatus: {}, teamStatus: {} };
  }
}

// 초기 데이터 설정 함수 - 담당자 시트 포함
function initializeTestData() {
  try {
    const spreadsheet = getSpreadsheet();
    
    // 0. 담당자 시트 생성
    let memberSheet = spreadsheet.getSheetByName(SHEETS.members);
    if (!memberSheet) {
      spreadsheet.insertSheet(SHEETS.members);
      memberSheet = spreadsheet.getSheetByName(SHEETS.members);
      memberSheet.appendRow(['담당자ID', '이름', '역할', '소속팀ID', '소속채널ID']);
    }
    
    // 1. 팀 데이터 추가
    const teamSheet = spreadsheet.getSheetByName(SHEETS.teams);
    if (teamSheet.getLastRow() <= 1) {
      teamSheet.appendRow(['T01', '콘텐츠1팀', 'lead01@company.kr']);
      teamSheet.appendRow(['T02', '콘텐츠2팀', 'lead02@company.kr']);
      teamSheet.appendRow(['T03', '콘텐츠3팀', 'admin001@company.kr']);
      console.log('팀 데이터 추가 완료');
    }
    
    // 2. 채널 데이터 추가
    const channelSheet = spreadsheet.getSheetByName(SHEETS.channels);
    if (channelSheet.getLastRow() <= 1) {
      channelSheet.appendRow(['CH001', '메인 채널', 'T01', true]);
      channelSheet.appendRow(['CH002', '서브 채널', 'T01', true]);
      channelSheet.appendRow(['CH003', '뉴스 채널', 'T02', true]);
      channelSheet.appendRow(['CH004', '엔터 채널', 'T03', true]);
      channelSheet.appendRow(['CH005', '스포츠 채널', 'T03', true]);
      console.log('채널 데이터 추가 완료');
    }
    
    // 3. 사용자 데이터 추가 (시스템 관리자만)
    const userSheet = spreadsheet.getSheetByName(SHEETS.users);
    if (userSheet.getLastRow() <= 1) {
      userSheet.appendRow(['admin001@company.kr', 'admin001@company.kr', '김관리자', '팀장', 'T03', 'Y', 'N', 'N']);
      userSheet.appendRow(['admin002@company.kr', 'admin002@company.kr', '이관리자', '관리자', 'T01,T02,T03', 'N', 'Y', 'N']);
      userSheet.appendRow(['admin003@company.kr', 'admin003@company.kr', '박관리자', '최종관리자', 'T01,T02,T03', 'N', 'N', 'Y']);
      console.log('사용자 데이터 추가 완료');
    }
    
    // 4. 담당자 데이터 추가
    if (memberSheet.getLastRow() <= 1) {
      memberSheet.appendRow(['M001', '홍길동', 'PD', 'T03', 'CH004']);
      memberSheet.appendRow(['M002', '김영희', '편집자', 'T03', 'CH004']);
      memberSheet.appendRow(['M003', '이철수', '작가', 'T01', 'CH001']);
      memberSheet.appendRow(['M004', '박민수', '디자이너', 'T01', 'CH001']);
      memberSheet.appendRow(['M005', '정수진', 'PD', 'T02', 'CH003']);
      console.log('담당자 데이터 추가 완료');
    }
    
    return { success: true, message: '초기 데이터 설정 완료' };
  } catch (error) {
    console.error('초기 데이터 설정 오류:', error);
    return { success: false, message: error.toString() };
  }
}

// 나머지 함수들은 기존 코드와 동일...

// ============ 인증 관련 함수 ============

// 로그인 처리
function login(email, password) {
  try {
    const users = getSheetData(SHEETS.users);
    const user = users.find(u => 
      u['이메일'] === email && 
      u['비밀번호'] === password && 
      u['활성상태'] !== 'N'
    );
    
    if (user) {
      // 권한 정보 추가
      user['평가권한'] = [];
      if (user['1차평가권한'] === 'Y') user['평가권한'].push('1차');
      if (user['2차평가권한'] === 'Y') user['평가권한'].push('2차');
      if (user['3차평가권한'] === 'Y') user['평가권한'].push('3차');
      
      // 복수 팀 소속 처리
      if (user['소속팀ID'] && user['소속팀ID'].includes(',')) {
        user['소속팀목록'] = user['소속팀ID'].split(',').map(t => t.trim());
      } else {
        user['소속팀목록'] = user['소속팀ID'] ? [user['소속팀ID']] : [];
      }
      
      // 세션에 사용자 정보 저장
      const userProperties = PropertiesService.getUserProperties();
      userProperties.setProperty('currentUser', JSON.stringify({
        사용자ID: user['사용자ID'],
        이메일: user['이메일'],
        이름: user['이름'],
        평가권한: user['평가권한'],
        소속팀목록: user['소속팀목록']
      }));
      
      return {
        success: true,
        user: user
      };
    } else {
      return {
        success: false,
        message: '이메일 또는 비밀번호가 일치하지 않습니다.'
      };
    }
  } catch (error) {
    console.error('로그인 오류:', error);
    return {
      success: false,
      message: '로그인 처리 중 오류가 발생했습니다.'
    };
  }
}

// 세션 확인
function checkSession() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const currentUserStr = userProperties.getProperty('currentUser');
    
    if (currentUserStr) {
      const currentUser = JSON.parse(currentUserStr);
      // 전체 사용자 정보 다시 로드
      const users = getSheetData(SHEETS.users);
      const user = users.find(u => u['사용자ID'] === currentUser.사용자ID);
      
      if (user && user['활성상태'] !== 'N') {
        // 권한 정보 재설정
        user['평가권한'] = currentUser.평가권한;
        user['소속팀목록'] = currentUser.소속팀목록;
        return {
          success: true,
          user: user
        };
      }
    }
    
    return {
      success: false
    };
  } catch (error) {
    console.error('세션 확인 오류:', error);
    return {
      success: false
    };
  }
}

// 로그아웃
function logout() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty('currentUser');
  return { success: true };
}

// ============ API 함수들 ============

// 현재 사용자 정보 가져오기
function getCurrentUser(userId) {
  const users = getSheetData(SHEETS.users);
  const user = users.find(u => u['사용자ID'] === userId);
  
  if (user) {
    // 권한 정보 추가
    user['평가권한'] = [];
    if (user['1차평가권한'] === 'Y') user['평가권한'].push('1차');
    if (user['2차평가권한'] === 'Y') user['평가권한'].push('2차');
    if (user['3차평가권한'] === 'Y') user['평가권한'].push('3차');
    
    // 복수 팀 소속 처리
    if (user['소속팀ID'] && user['소속팀ID'].includes(',')) {
      user['소속팀목록'] = user['소속팀ID'].split(',').map(t => t.trim());
    } else {
      user['소속팀목록'] = user['소속팀ID'] ? [user['소속팀ID']] : [];
    }
  }
  
  return user || null;
}

// 역할별 데이터 가져오기
function getDataByRole(userId, role, evaluationMonth) {
  console.log('getDataByRole 호출:', {userId, role, evaluationMonth});
  
  try {
    const result = {
      teams: [],
      channels: [],
      evaluations: [],
      users: [],      // 관리자들
      members: [],    // 담당자들
      progress: []
    };
    
    // 데이터 로드
    const users = getSheetData(SHEETS.users);
    const members = getSheetData(SHEETS.members);  // 담당자 데이터
    const teams = getSheetData(SHEETS.teams);
    const channels = getSheetData(SHEETS.channels);
    const evaluations = getSheetData(SHEETS.evaluations);
    const progress = getSheetData(SHEETS.progress);
    
    // 활성 팀과 채널만 필터링
    const activeTeams = teams.filter(t => t['활성상태'] !== 'N');
    const activeChannels = channels.filter(c => c['활성상태'] !== 'N');
    
    // 평가월 필터링
    const monthEvaluations = evaluationMonth ? 
      evaluations.filter(e => e['평가월'] === evaluationMonth) : 
      evaluations;
    
    const monthProgress = evaluationMonth ? 
      progress.filter(p => p['평가월'] === evaluationMonth) : 
      progress;
    
    if (role === '팀장') {
      // 팀장은 자신의 팀 데이터만
      const currentUser = users.find(u => u['사용자ID'] === userId);
      if (currentUser && currentUser['소속팀ID']) {
        const userTeams = currentUser['소속팀ID'].split(',').map(t => t.trim());
        
        result.teams = activeTeams.filter(t => userTeams.includes(t['팀ID']));
        result.channels = activeChannels.filter(c => userTeams.includes(c['소속팀ID']));
        result.users = users.filter(u => {
          const uTeams = u['소속팀ID'] ? u['소속팀ID'].split(',').map(t => t.trim()) : [];
          return uTeams.some(t => userTeams.includes(t));
        });
        result.members = members.filter(m => {
          const mTeams = m['소속팀ID'] ? m['소속팀ID'].split(',').map(t => t.trim()) : [];
          return mTeams.some(t => userTeams.includes(t));
        });
        result.evaluations = monthEvaluations.filter(e => {
          const channelTeam = activeChannels.find(c => c['채널ID'] === e['채널ID'])?.['소속팀ID'];
          return userTeams.includes(channelTeam);
        });
        result.progress = monthProgress.filter(p => userTeams.includes(p['팀ID']));
      }
    } else {
      // 관리자와 최종관리자는 모든 데이터
      result.teams = activeTeams;
      result.channels = activeChannels;
      result.users = users;
      result.members = members;
      result.evaluations = monthEvaluations;
      result.progress = monthProgress;
    }
    
    return result;
  } catch (error) {
    console.error('getDataByRole 오류:', error);
    throw new Error('데이터를 불러올 수 없습니다: ' + error.toString());
  }
}

// 평가 데이터 일괄 저장
function saveEvaluationBatch(evaluationDataArray) {
  try {
    console.log('saveEvaluationBatch 시작, 데이터 개수:', evaluationDataArray.length);
    
    const spreadsheet = getSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEETS.evaluations);
    
    if (!sheet) {
      throw new Error('평가 시트를 찾을 수 없습니다.');
    }
    
    // 헤더 확인 및 생성
    const defaultHeaders = [
      '평가ID', '평가월', '채널ID', '담당자ID', '담당자이름', '투입MM', 
      '1차기여도', '1차성과내용', '1차코멘트', '1차평가상태', '1차평가자ID', '1차평가일시',
      '2차기여도', '2차코멘트', '2차평가상태', '2차평가자ID', '2차평가일시', 
      '사업효율', '팀장성과금', '팀장성과', '실지급성과금', '최종코멘트', 
      '최종평가상태', '최종평가자ID', '최종평가일시', '채널코멘트'
    ];
    
    let headers;
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(defaultHeaders);
      headers = defaultHeaders;
    } else {
      const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
      headers = headerRange.getValues()[0];
      
      // 빈 헤더 확인
      if (!headers[0]) {
        sheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
        headers = defaultHeaders;
      }
    }
    
    const dataRange = sheet.getDataRange();
    const existingData = dataRange.getValues();
    
    // 기존 데이터를 Map으로 변환
    const existingMap = new Map();
    for (let i = 1; i < existingData.length; i++) {
      const evalId = existingData[i][0];
      if (evalId) {
        const existingRow = {};
        headers.forEach((header, index) => {
          existingRow[header] = existingData[i][index];
        });
        existingMap.set(evalId, { rowIndex: i + 1, data: existingRow });
      }
    }
    
    const updates = [];
    const appends = [];
    const currentTime = new Date();
    
    evaluationDataArray.forEach(evalData => {
      // 평가ID가 없으면 생성
      if (!evalData['평가ID']) {
        evalData['평가ID'] = generateEvaluationId(evalData['평가월'] || '');
      }
      
      const existing = existingMap.get(evalData['평가ID']);
      
      if (existing) {
        // 기존 데이터와 병합
        const mergedData = { ...existing.data };
        Object.keys(evalData).forEach(key => {
          if (evalData[key] !== undefined && evalData[key] !== '') {
            mergedData[key] = evalData[key];
          }
        });
        
        // 평가자와 시간 기록
        if (evalData['1차평가상태']) {
          mergedData['1차평가자ID'] = evalData['평가자ID'];
          mergedData['1차평가일시'] = currentTime;
        }
        if (evalData['2차평가상태']) {
          mergedData['2차평가자ID'] = evalData['평가자ID'];
          mergedData['2차평가일시'] = currentTime;
        }
        if (evalData['최종평가상태']) {
          mergedData['최종평가자ID'] = evalData['평가자ID'];
          mergedData['최종평가일시'] = currentTime;
        }
        
        const rowData = headers.map(header => mergedData[header] || '');
        updates.push({ row: existing.rowIndex, data: rowData });
      } else {
        // 신규 데이터
        if (evalData['1차평가상태']) {
          evalData['1차평가자ID'] = evalData['평가자ID'];
          evalData['1차평가일시'] = currentTime;
        }
        if (evalData['2차평가상태']) {
          evalData['2차평가자ID'] = evalData['평가자ID'];
          evalData['2차평가일시'] = currentTime;
        }
        if (evalData['최종평가상태']) {
          evalData['최종평가자ID'] = evalData['평가자ID'];
          evalData['최종평가일시'] = currentTime;
        }
        
        const rowData = headers.map(header => evalData[header] || '');
        appends.push(rowData);
      }
    });
    
    // 업데이트 실행
    updates.forEach(update => {
      sheet.getRange(update.row, 1, 1, update.data.length).setValues([update.data]);
    });
    
    // 추가 실행
    if (appends.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, appends.length, appends[0].length).setValues(appends);
    }
    
    // 진행상태 업데이트
    updateProgressStatus(evaluationDataArray[0]['평가월']);
    
    return { success: true, message: `${updates.length}개 업데이트, ${appends.length}개 추가 완료` };
  } catch (error) {
    console.error('saveEvaluationBatch 오류:', error);
    return { success: false, message: error.toString() };
  }
}

// 평가 진행상태 업데이트
function updateProgressStatus(evaluationMonth) {
  try {
    const spreadsheet = getSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEETS.progress);
    
    if (!sheet) {
      // 진행상태 시트가 없으면 생성
      const newSheet = spreadsheet.insertSheet(SHEETS.progress);
      const headers = ['상태ID', '평가월', '팀ID', '채널ID', '1차진행상태', '2차진행상태', '3차진행상태', '최종완료여부'];
      newSheet.appendRow(headers);
    }
    
    // 해당 월의 평가 데이터 분석
    const evaluations = getSheetData(SHEETS.evaluations).filter(e => e['평가월'] === evaluationMonth);
    const channels = getSheetData(SHEETS.channels);
    const teams = getSheetData(SHEETS.teams);
    
    // 채널별 진행상태 계산
    const channelProgress = {};
    
    channels.forEach(channel => {
      const channelId = channel['채널ID'];
      const teamId = channel['소속팀ID'];
      const channelEvals = evaluations.filter(e => e['채널ID'] === channelId);
      
      if (channelEvals.length > 0) {
        const total = channelEvals.length;
        const eval1Complete = channelEvals.filter(e => e['1차평가상태'] === '제출완료').length;
        const eval2Complete = channelEvals.filter(e => e['2차평가상태'] === '제출완료').length;
        const eval3Complete = channelEvals.filter(e => e['최종평가상태'] === '제출완료').length;
        
        channelProgress[channelId] = {
          teamId: teamId,
          status1: `${eval1Complete}/${total}`,
          status2: `${eval2Complete}/${total}`,
          status3: `${eval3Complete}/${total}`,
          complete: eval3Complete === total ? 'Y' : 'N'
        };
      }
    });
    
    // 진행상태 시트 업데이트
    const progressSheet = spreadsheet.getSheetByName(SHEETS.progress);
    if (progressSheet && progressSheet.getLastRow() > 1) {
      // 기존 데이터 삭제 (해당 월)
      const data = progressSheet.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][1] === evaluationMonth) {
          progressSheet.deleteRow(i + 1);
        }
      }
    }
    
    // 새 데이터 추가
    Object.keys(channelProgress).forEach(channelId => {
      const progress = channelProgress[channelId];
      const stateId = 'PROG-' + new Date().getTime() + '-' + channelId;
      progressSheet.appendRow([
        stateId,
        evaluationMonth,
        progress.teamId,
        channelId,
        progress.status1,
        progress.status2,
        progress.status3,
        progress.complete
      ]);
    });
    
  } catch (error) {
    console.error('진행상태 업데이트 오류:', error);
  }
}

// 새 평가ID 생성
function generateEvaluationId(evaluationMonth) {
  try {
    const spreadsheet = getSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEETS.evaluations);
    
    // YYYYMM 형식 생성
    const monthMatch = evaluationMonth.match(/(\d{4})년\s*(\d{1,2})월/);
    const yyyymm = monthMatch ? 
      monthMatch[1] + monthMatch[2].padStart(2, '0') : 
      new Date().getFullYear() + (new Date().getMonth() + 1).toString().padStart(2, '0');
    
    // 해당 월의 기존 평가ID 찾기
    const prefix = `EVAL-${yyyymm}-`;
    let maxNum = 0;
    
    if (sheet && sheet.getLastRow() > 0) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const evalId = data[i][0];
        if (evalId && evalId.toString().startsWith(prefix)) {
          const numStr = evalId.toString().substring(prefix.length);
          const num = parseInt(numStr);
          if (!isNaN(num) && num > maxNum) {
            maxNum = num;
          }
        }
      }
    }
    
    return prefix + (maxNum + 1).toString().padStart(3, '0');
  } catch (error) {
    // 에러 발생 시 타임스탬프 기반 ID 생성
    return 'EVAL-' + new Date().getTime();
  }
}

// 평가 월 목록 가져오기
function getEvaluationMonths() {
  try {
    const evaluations = getSheetData(SHEETS.evaluations);
    const progress = getSheetData(SHEETS.progress);
    
    // 평가 데이터와 진행상태 데이터에서 월 추출
    const monthsFromEval = evaluations.map(e => e['평가월']).filter(m => m);
    const monthsFromProgress = progress.map(p => p['평가월']).filter(m => m);
    
    // 중복 제거 및 정렬
    const allMonths = [...new Set([...monthsFromEval, ...monthsFromProgress])];
    const sortedMonths = allMonths.sort().reverse();
    
    // 현재 월 추가 (없는 경우)
    const currentDate = new Date();
    const currentYearMonth = currentDate.getFullYear() + '년 ' + (currentDate.getMonth() + 1) + '월';
    
    if (!sortedMonths.includes(currentYearMonth)) {
      sortedMonths.unshift(currentYearMonth);
    }
    
    return sortedMonths;
  } catch (error) {
    console.error('getEvaluationMonths 오류:', error);
    const currentDate = new Date();
    const currentYearMonth = currentDate.getFullYear() + '년 ' + (currentDate.getMonth() + 1) + '월';
    return [currentYearMonth];
  }
}
// 평가 진행 상태 조회 함수 - 새로 추가
function getEvaluationStatus(evaluationMonth) {
  try {
    const evaluations = getSheetData(SHEETS.evaluations);
    const channels = getSheetData(SHEETS.channels);
    const teams = getSheetData(SHEETS.teams);
    
    const monthEvaluations = evaluations.filter(e => e['평가월'] === evaluationMonth);
    
    const statusByChannel = {};
    const statusByTeam = {};
    
    // 채널별 진행상태 계산
    channels.forEach(channel => {
      const channelEvals = monthEvaluations.filter(e => e['채널ID'] === channel['채널ID']);
      
      if (channelEvals.length === 0) {
        statusByChannel[channel['채널ID']] = '미시작';
      } else {
        const has1st = channelEvals.some(e => e['1차평가상태'] === '제출완료');
        const has2nd = channelEvals.some(e => e['2차평가상태'] === '제출완료');
        const has3rd = channelEvals.some(e => e['최종평가상태'] === '제출완료');
        
        if (has3rd) statusByChannel[channel['채널ID']] = '완료';
        else if (has2nd) statusByChannel[channel['채널ID']] = '3차대기';
        else if (has1st) statusByChannel[channel['채널ID']] = '2차대기';
        else statusByChannel[channel['채널ID']] = '1차진행중';
      }
    });
    
    // 팀별 진행상태 계산
    teams.forEach(team => {
      const teamChannels = channels.filter(c => c['소속팀ID'] === team['팀ID']);
      const statuses = teamChannels.map(c => statusByChannel[c['채널ID']] || '미시작');
      
      if (statuses.every(s => s === '완료')) statusByTeam[team['팀ID']] = '완료';
      else if (statuses.some(s => s === '3차대기')) statusByTeam[team['팀ID']] = '3차대기';
      else if (statuses.some(s => s === '2차대기')) statusByTeam[team['팀ID']] = '2차대기';
      else if (statuses.some(s => s === '1차진행중')) statusByTeam[team['팀ID']] = '1차진행중';
      else statusByTeam[team['팀ID']] = '미시작';
    });
    
    return { channelStatus: statusByChannel, teamStatus: statusByTeam };
  } catch (error) {
    console.error('getEvaluationStatus 오류:', error);
    return { channelStatus: {}, teamStatus: {} };
  }
}
// 채널 추가
function addChannel(channelData) {
  try {
    const spreadsheet = getSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEETS.channels);
    
    if (!sheet) {
      throw new Error('채널 시트를 찾을 수 없습니다.');
    }
    
    // 기존 채널 ID들을 가져와서 최대값 찾기
    const data = sheet.getDataRange().getValues();
    let maxNum = 0;
    
    for (let i = 1; i < data.length; i++) {
      const channelId = data[i][0];
      if (channelId && channelId.toString().startsWith('CH')) {
        const num = parseInt(channelId.toString().substring(2));
        if (!isNaN(num) && num > maxNum) {
          maxNum = num;
        }
      }
    }
    
    const newId = 'CH' + String(maxNum + 1).padStart(3, '0');
    sheet.appendRow([newId, channelData.name, channelData.teamId, 'Y']);
    return { success: true, channelId: newId };
  } catch (error) {
    console.error('addChannel 오류:', error);
    return { success: false, message: error.toString() };
  }
}

// 팀 추가
function addTeam(teamData) {
  try {
    const spreadsheet = getSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEETS.teams);
    
    if (!sheet) {
      throw new Error('팀 시트를 찾을 수 없습니다.');
    }
    
    // 기존 팀 ID들을 가져와서 최대값 찾기
    const data = sheet.getDataRange().getValues();
    let maxNum = 0;
    
    for (let i = 1; i < data.length; i++) {
      const teamId = data[i][0];
      if (teamId && teamId.toString().startsWith('T')) {
        const num = parseInt(teamId.toString().substring(1));
        if (!isNaN(num) && num > maxNum) {
          maxNum = num;
        }
      }
    }
    
    const newId = 'T' + String(maxNum + 1).padStart(2, '0');
    sheet.appendRow([newId, teamData.name, teamData.leaderId || '', 'Y']);
    return { success: true, teamId: newId };
  } catch (error) {
    console.error('addTeam 오류:', error);
    return { success: false, message: error.toString() };
  }
}

// 담당자 추가 (평가받는 사람)
function addMember(memberData) {
  try {
    const spreadsheet = getSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEETS.members);
    
    if (!sheet) {
      throw new Error('담당자 시트를 찾을 수 없습니다.');
    }
    
    // 헤더가 없으면 생성
    if (sheet.getLastRow() === 0) {
      const headers = ['담당자ID', '이름', '역할', '소속팀ID', '활성상태', '생성일시'];
      sheet.appendRow(headers);
    }
    
    // 담당자ID 생성 (MEM + 숫자)
    const data = sheet.getDataRange().getValues();
    let maxNum = 0;
    
    for (let i = 1; i < data.length; i++) {
      const memberId = data[i][0];
      if (memberId && memberId.toString().startsWith('MEM')) {
        const num = parseInt(memberId.toString().substring(3));
        if (!isNaN(num) && num > maxNum) {
          maxNum = num;
        }
      }
    }
    
    const memberId = 'MEM' + String(maxNum + 1).padStart(3, '0');
    
    sheet.appendRow([
      memberId,
      memberData.name,
      memberData.role,
      memberData.teamId,
      'Y',
      new Date()
    ]);
    
    return { 
      success: true, 
      memberId: memberId,
      member: {
        담당자ID: memberId,
        이름: memberData.name,
        역할: memberData.role,
        소속팀ID: memberData.teamId
      }
    };
  } catch (error) {
    console.error('addMember 오류:', error);
    return { success: false, message: error.toString() };
  }
}

// 사용자 추가 (관리자 - 로그인 가능한 사람만)
function addUser(userData) {
  try {
    const spreadsheet = getSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEETS.users);
    
    if (!sheet) {
      throw new Error('사용자 시트를 찾을 수 없습니다.');
    }
    
    // 헤더가 없으면 생성
    if (sheet.getLastRow() === 0) {
      const headers = ['사용자ID', '이메일', '이름', '비밀번호', '역할', '소속팀ID', 
                       '1차평가권한', '2차평가권한', '3차평가권한', '활성상태', '생성일시'];
      sheet.appendRow(headers);
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 시스템 관리자는 이메일을 ID로 사용
    const userId = userData.userId;
    
    // 권한 설정
    const permissions = {
      '1차평가권한': userData.permissions?.perm1 !== undefined ? 
        (userData.permissions.perm1 ? 'Y' : 'N') :
        (userData.role === '팀장' ? 'Y' : 'N'),
      '2차평가권한': userData.permissions?.perm2 !== undefined ? 
        (userData.permissions.perm2 ? 'Y' : 'N') :
        (userData.role === '관리자' ? 'Y' : 'N'),
      '3차평가권한': userData.permissions?.perm3 !== undefined ? 
        (userData.permissions.perm3 ? 'Y' : 'N') :
        (userData.role === '최종관리자' ? 'Y' : 'N')
    };
    
    const rowData = headers.map(header => {
      switch(header) {
        case '사용자ID': return userId;
        case '이메일': return userId;
        case '이름': return userData.name;
        case '비밀번호': return userData.password || '1234';
        case '역할': return userData.role;
        case '소속팀ID': return userData.teamId;
        case '1차평가권한': return permissions['1차평가권한'];
        case '2차평가권한': return permissions['2차평가권한'];
        case '3차평가권한': return permissions['3차평가권한'];
        case '활성상태': return 'Y';
        case '생성일시': return new Date();
        default: return '';
      }
    });
    
    sheet.appendRow(rowData);
    return { success: true, userId: userId };
  } catch (error) {
    console.error('addUser 오류:', error);
    return { success: false, message: error.toString() };
  }
}

// 테스트 함수
function testConnection() {
  try {
    const spreadsheet = getSpreadsheet();
    const sheets = spreadsheet.getSheets();
    const sheetNames = sheets.map(s => s.getName());
    
    console.log('연결 성공! 시트 목록:', sheetNames);
    
    // 각 시트의 데이터 수 확인
    const result = {};
    for (const [key, sheetName] of Object.entries(SHEETS)) {
      try {
        const data = getSheetData(sheetName);
        result[key] = data.length;
      } catch (e) {
        result[key] = 'Error: ' + e.toString();
      }
    }
    
    return {
      success: true,
      sheets: sheetNames,
      dataCount: result
    };
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// 세션 테스트
function testSession() {
  const userProperties = PropertiesService.getUserProperties();
  const currentUser = userProperties.getProperty('currentUser');
  
  if (currentUser) {
    console.log('현재 세션:', JSON.parse(currentUser));
    return '세션이 활성화되어 있습니다.';
  } else {
    console.log('세션 없음');
    return '세션이 없습니다.';
  }
}

// 샘플 데이터 생성
function createSampleData() {
  const spreadsheet = getSpreadsheet();
  
  // 팀장 추가
  const userSheet = spreadsheet.getSheetByName(SHEETS.users);
  userSheet.appendRow([
    'team1@company.com', 'team1@company.com', '이팀장', 'team123',
    '팀장', 'T01', 'Y', 'N', 'N', 'Y', new Date()
  ]);
  
  // 관리자 추가
  userSheet.appendRow([
    'manager@company.com', 'manager@company.com', '박관리', 'manager123',
    '관리자', 'T01,T02', 'N', 'Y', 'N', 'Y', new Date()
  ]);
  
  // 담당자 추가
  const memberSheet = spreadsheet.getSheetByName(SHEETS.members);
  memberSheet.appendRow(['MEM003', '김작가', '작가', 'T01', 'Y', new Date()]);
  memberSheet.appendRow(['MEM004', '최편집', '편집자', 'T01', 'Y', new Date()]);
  memberSheet.appendRow(['MEM005', '정PD', 'PD', 'T02', 'Y', new Date()]);
  
  // 채널 추가
  const channelSheet = spreadsheet.getSheetByName(SHEETS.channels);
  channelSheet.appendRow(['CH003', '교육채널', 'T02', 'Y']);
  
  return '샘플 데이터 생성 완료!';
}

// 모든 세션 초기화
function clearAllSessions() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteAllProperties();
  return '모든 세션이 초기화되었습니다.';
}

// 초기 시트 설정 (처음 실행 시)
function initSheets() {
  const spreadsheet = getSpreadsheet();
  
  // 1. 사용자 시트 (관리자만)
  let sheet = spreadsheet.getSheetByName(SHEETS.users);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEETS.users);
    sheet.appendRow(['사용자ID', '이메일', '이름', '비밀번호', '역할', '소속팀ID', 
                     '1차평가권한', '2차평가권한', '3차평가권한', '활성상태', '생성일시']);
    
    // 샘플 데이터
    sheet.appendRow(['admin@company.com', 'admin@company.com', '김관리자', 'admin123', 
                     '최종관리자', 'T01', 'Y', 'Y', 'Y', 'Y', new Date()]);
  }
  
  // 2. 담당자 시트 (평가받는 사람들)
  sheet = spreadsheet.getSheetByName(SHEETS.members);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEETS.members);
    sheet.appendRow(['담당자ID', '이름', '역할', '소속팀ID', '활성상태', '생성일시']);
    
    // 샘플 데이터
    sheet.appendRow(['MEM001', '박담당', 'PD', 'T01', 'Y', new Date()]);
    sheet.appendRow(['MEM002', '이편집', '편집자', 'T01', 'Y', new Date()]);
  }
  
  // 3. 팀 시트
  sheet = spreadsheet.getSheetByName(SHEETS.teams);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEETS.teams);
    sheet.appendRow(['팀ID', '팀이름', '팀장ID', '활성상태']);
    
    // 샘플 데이터
    sheet.appendRow(['T01', '콘텐츠1팀', 'team1@company.com', 'Y']);
    sheet.appendRow(['T02', '콘텐츠2팀', 'team2@company.com', 'Y']);
  }
  
  // 4. 채널 시트
  sheet = spreadsheet.getSheetByName(SHEETS.channels);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEETS.channels);
    sheet.appendRow(['채널ID', '채널이름', '소속팀ID', '활성상태']);
    
    // 샘플 데이터
    sheet.appendRow(['CH001', '메인채널', 'T01', 'Y']);
    sheet.appendRow(['CH002', '서브채널', 'T01', 'Y']);
  }
  
  // 5. 월별평가 시트
  sheet = spreadsheet.getSheetByName(SHEETS.evaluations);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEETS.evaluations);
    sheet.appendRow([
      '평가ID', '평가월', '채널ID', '담당자ID', '담당자이름', '투입MM', 
      '1차기여도', '1차성과내용', '1차코멘트', '1차평가상태', '1차평가자ID', '1차평가일시',
      '2차기여도', '2차코멘트', '2차평가상태', '2차평가자ID', '2차평가일시', 
      '사업효율', '팀장성과금', '팀장성과', '실지급성과금', '최종코멘트', 
      '최종평가상태', '최종평가자ID', '최종평가일시', '채널코멘트'
    ]);
  }
  
  // 6. 평가진행상태 시트
  sheet = spreadsheet.getSheetByName(SHEETS.progress);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEETS.progress);
    sheet.appendRow(['상태ID', '평가월', '팀ID', '채널ID', '1차진행상태', '2차진행상태', '3차진행상태', '최종완료여부']);
  }
  
  return '시트 초기화 완료!';
}

// Apps Script 편집기에서 실행할 시트 정리 함수

function cleanupAndFixSheets() {
  const spreadsheet = getSpreadsheet();
  
  // 1. 기존 사용자 시트 백업
  const oldUserSheet = spreadsheet.getSheetByName('사용자');
  if (oldUserSheet) {
    oldUserSheet.setName('사용자_백업_' + new Date().getTime());
  }
  
  // 2. 새 사용자 시트 생성
  const userSheet = spreadsheet.insertSheet('사용자');
  const userHeaders = ['사용자ID', '이메일', '이름', '비밀번호', '역할', '소속팀ID', 
                       '1차평가권한', '2차평가권한', '3차평가권한', '활성상태', '생성일시'];
  userSheet.appendRow(userHeaders);
  
  // 3. 관리자 계정 추가
  userSheet.appendRow([
    'admin@company.com',
    'admin@company.com',
    '어드민',
    'admin123',
    '최종관리자',
    'T01,T02,T03,T04,T05,T06',
    'Y',
    'Y',
    'Y',
    'Y',
    new Date()
  ]);
  
  // 팀장들 추가
  userSheet.appendRow([
    'team1@company.com',
    'team1@company.com',
    '1팀팀장',
    'team123',
    '팀장',
    'T01',
    'Y',
    'N',
    'N',
    'Y',
    new Date()
  ]);
  
  userSheet.appendRow([
    'manager@company.com',
    'manager@company.com',
    '박관리',
    'manager123',
    '관리자',
    'T01,T02',
    'N',
    'Y',
    'N',
    'Y',
    new Date()
  ]);
  
  // 4. 담당자 시트 확인/생성
  let memberSheet = spreadsheet.getSheetByName('담당자');
  if (!memberSheet) {
    memberSheet = spreadsheet.insertSheet('담당자');
    const memberHeaders = ['담당자ID', '이름', '역할', '소속팀ID', '활성상태', '생성일시'];
    memberSheet.appendRow(memberHeaders);
  }
  
  // 5. 팀 시트 확인
  let teamSheet = spreadsheet.getSheetByName('팀');
  if (!teamSheet) {
    teamSheet = spreadsheet.insertSheet('팀');
    const teamHeaders = ['팀ID', '팀이름', '팀장ID', '활성상태'];
    teamSheet.appendRow(teamHeaders);
    
    // 샘플 팀 데이터
    teamSheet.appendRow(['T01', '팀원1', 'team1@company.com', 'Y']);
    teamSheet.appendRow(['T02', '팀원2', 'team2@company.com', 'Y']);
    teamSheet.appendRow(['T03', '팀원3', 'team3@company.com', 'Y']);
    teamSheet.appendRow(['T04', '팀원4', 'team4@company.com', 'Y']);
    teamSheet.appendRow(['T05', '팀원5', 'team5@company.com', 'Y']);
    teamSheet.appendRow(['T06', '팀원6', 'team6@company.com', 'Y']);
  }
  
  // 6. 채널 시트 확인
  let channelSheet = spreadsheet.getSheetByName('채널');
  if (!channelSheet) {
    channelSheet = spreadsheet.insertSheet('채널');
    const channelHeaders = ['채널ID', '채널이름', '소속팀ID', '활성상태'];
    channelSheet.appendRow(channelHeaders);
    
    // 샘플 채널 데이터
    channelSheet.appendRow(['CH001', '메인채널', 'T01', 'Y']);
    channelSheet.appendRow(['CH002', '서브채널', 'T01', 'Y']);
  }
  
  console.log('시트 정리 완료!');
  return '시트가 정리되었습니다. 이제 로그인을 시도해보세요.';
}

// 기존 백업 시트에서 필요한 데이터 마이그레이션
function migrateOldData() {
  const spreadsheet = getSpreadsheet();
  const backupSheets = spreadsheet.getSheets().filter(s => s.getName().includes('사용자_백업'));
  
  if (backupSheets.length === 0) {
    return '백업 시트가 없습니다.';
  }
  
  const oldSheet = backupSheets[0];
  const oldData = oldSheet.getDataRange().getValues();
  
  const memberSheet = spreadsheet.getSheetByName('담당자');
  
  // 이전 데이터에서 담당자들 추출
  for (let i = 1; i < oldData.length; i++) {
    const row = oldData[i];
    if (!row[0]) continue; // 빈 행 무시
    
    // 이메일이 없는 경우 담당자로 분류
    if (!row[1] || row[1] === '') {
      const memberId = 'MEM' + String(i).padStart(3, '0');
      const name = row[2] || '이름없음';
      const role = row[4] || 'PD';
      const teamId = row[5] || 'T01';
      
      memberSheet.appendRow([memberId, name, role, teamId, 'Y', new Date()]);
    }
  }
  
  return '데이터 마이그레이션 완료!';
}

// ============ 로그인 관련 함수들 ============
// Code.gs의 맨 아래에 이 함수들을 추가하세요

// 로그인 처리
function login(email, password) {
  try {
    console.log('로그인 시도:', email);
    
    const users = getSheetData(SHEETS.users);
    console.log('전체 사용자 수:', users.length);
    
    // 이메일 또는 사용자ID로 찾기
    const user = users.find(u => 
      (u['이메일'] === email || u['사용자ID'] === email) && 
      u['비밀번호'] === password && 
      u['활성상태'] !== 'N'
    );
    
    if (user) {
      console.log('로그인 성공:', user['이름']);
      
      // 권한 정보 추가
      user['평가권한'] = [];
      if (user['1차평가권한'] === 'Y') user['평가권한'].push('1차');
      if (user['2차평가권한'] === 'Y') user['평가권한'].push('2차');
      if (user['3차평가권한'] === 'Y') user['평가권한'].push('3차');
      
      // 복수 팀 소속 처리
      if (user['소속팀ID'] && user['소속팀ID'].includes(',')) {
        user['소속팀목록'] = user['소속팀ID'].split(',').map(t => t.trim());
      } else {
        user['소속팀목록'] = user['소속팀ID'] ? [user['소속팀ID']] : [];
      }
      
      // 세션에 사용자 정보 저장
      const userProperties = PropertiesService.getUserProperties();
      userProperties.setProperty('currentUser', JSON.stringify({
        사용자ID: user['사용자ID'],
        이메일: user['이메일'],
        이름: user['이름'],
        역할: user['역할'],
        평가권한: user['평가권한'],
        소속팀목록: user['소속팀목록']
      }));
      
      return {
        success: true,
        user: user
      };
    } else {
      console.log('로그인 실패: 사용자를 찾을 수 없음');
      return {
        success: false,
        message: '이메일 또는 비밀번호가 일치하지 않습니다.'
      };
    }
  } catch (error) {
    console.error('로그인 오류:', error);
    return {
      success: false,
      message: '로그인 처리 중 오류가 발생했습니다: ' + error.toString()
    };
  }
}

// 세션 확인
function checkSession() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const currentUserStr = userProperties.getProperty('currentUser');
    
    if (currentUserStr) {
      const currentUser = JSON.parse(currentUserStr);
      console.log('세션 사용자:', currentUser.이름);
      
      // 전체 사용자 정보 다시 로드
      const users = getSheetData(SHEETS.users);
      const user = users.find(u => 
        u['사용자ID'] === currentUser.사용자ID || 
        u['이메일'] === currentUser.이메일
      );
      
      if (user && user['활성상태'] !== 'N') {
        // 권한 정보 재설정
        user['평가권한'] = currentUser.평가권한;
        user['소속팀목록'] = currentUser.소속팀목록;
        
        return {
          success: true,
          user: user
        };
      }
    }
    
    return {
      success: false
    };
  } catch (error) {
    console.error('세션 확인 오류:', error);
    return {
      success: false
    };
  }
}

// 로그아웃
function logout() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    userProperties.deleteProperty('currentUser');
    return { success: true };
  } catch (error) {
    console.error('로그아웃 오류:', error);
    return { success: false };
  }
}

// 샘플 사용자 생성 (테스트용)
function createSampleUsers() {
  const spreadsheet = getSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEETS.users);
  
  if (!sheet) {
    console.error('사용자 시트를 찾을 수 없습니다.');
    return '사용자 시트가 없습니다.';
  }
  
  // 헤더가 없으면 생성
  if (sheet.getLastRow() === 0) {
    const headers = ['사용자ID', '이메일', '이름', '비밀번호', '역할', '소속팀ID', 
                     '1차평가권한', '2차평가권한', '3차평가권한', '활성상태', '생성일시'];
    sheet.appendRow(headers);
  }
  
  // 관리자 계정 추가
  const sampleUsers = [
    ['admin@company.com', 'admin@company.com', '김관리자', 'admin123', 
     '최종관리자', 'T01,T02,T03', 'Y', 'Y', 'Y', 'Y', new Date()],
    ['team1@company.com', 'team1@company.com', '이팀장', 'team123', 
     '팀장', 'T01', 'Y', 'N', 'N', 'Y', new Date()],
    ['manager@company.com', 'manager@company.com', '박매니저', 'manager123', 
     '관리자', 'T01,T02', 'N', 'Y', 'N', 'Y', new Date()]
  ];
  
  sampleUsers.forEach(user => {
    sheet.appendRow(user);
  });
  
  return '샘플 사용자가 생성되었습니다.';
}

// 사용자 시트 확인 및 수정
function checkAndFixUserSheet() {
  const spreadsheet = getSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEETS.users);
  
  if (!sheet) {
    // 시트가 없으면 생성
    const newSheet = spreadsheet.insertSheet(SHEETS.users);
    const headers = ['사용자ID', '이메일', '이름', '비밀번호', '역할', '소속팀ID', 
                     '1차평가권한', '2차평가권한', '3차평가권한', '활성상태', '생성일시'];
    newSheet.appendRow(headers);
    console.log('사용자 시트 생성 완료');
  }
  
  const data = sheet.getDataRange().getValues();
  console.log('사용자 시트 데이터:', data.length + '행');
  
  // 첫 번째 행(헤더) 확인
  if (data.length > 0) {
    const headers = data[0];
    console.log('현재 헤더:', headers);
    
    // 필수 컬럼 확인
    const requiredColumns = ['사용자ID', '이메일', '이름', '비밀번호'];
    const missingColumns = requiredColumns.filter(col => !headers.includes(col));
    
    if (missingColumns.length > 0) {
      console.error('누락된 컬럼:', missingColumns);
      return '필수 컬럼이 누락되었습니다: ' + missingColumns.join(', ');
    }
  }
  
  // 로그인 가능한 사용자 확인
  let validUsers = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][3]) { // 이메일과 비밀번호가 있는 경우
      validUsers++;
    }
  }
  
  console.log('로그인 가능한 사용자 수:', validUsers);
  
  if (validUsers === 0) {
    console.log('로그인 가능한 사용자가 없습니다. 샘플 사용자를 생성합니다.');
    createSampleUsers();
  }
  
  return '사용자 시트 확인 완료';
}

// 1. 현재 상태 확인
function quickTest() {
  console.log('=== 시트 확인 ===');
  const result = testConnection();
  console.log(result);
  
  console.log('\n=== 사용자 시트 확인 ===');
  checkAndFixUserSheet();
  
  console.log('\n=== 로그인 테스트 ===');
  const loginResult = login('admin@company.com', 'admin123');
  console.log('로그인 결과:', loginResult);
  
  return '테스트 완료';
}

// 2. 샘플 데이터 생성
function setupTestData() {
  // 사용자 생성
  createSampleUsers();
  
  // 팀 데이터 확인
  const teamSheet = getSpreadsheet().getSheetByName(SHEETS.teams);
  if (!teamSheet || teamSheet.getLastRow() <= 1) {
    teamSheet.appendRow(['T01', '콘텐츠1팀', 'team1@company.com', 'Y']);
    teamSheet.appendRow(['T02', '콘텐츠2팀', '', 'Y']);
    teamSheet.appendRow(['T03', '콘텐츠3팀', '', 'Y']);
  }
  
  // 채널 데이터 확인
  const channelSheet = getSpreadsheet().getSheetByName(SHEETS.channels);
  if (!channelSheet || channelSheet.getLastRow() <= 1) {
    channelSheet.appendRow(['CH001', '메인채널', 'T01', 'Y']);
    channelSheet.appendRow(['CH002', '서브채널', 'T01', 'Y']);
  }
  
  return '테스트 데이터 생성 완료';
}

// 초기 데이터 설정 함수 - 담당자 시트 포함
function initializeTestData() {
  try {
    const spreadsheet = getSpreadsheet();
    
    // 0. 담당자 시트 생성
    let memberSheet = spreadsheet.getSheetByName(SHEETS.members);
    if (!memberSheet) {
      spreadsheet.insertSheet(SHEETS.members);
      memberSheet = spreadsheet.getSheetByName(SHEETS.members);
      memberSheet.appendRow(['담당자ID', '이름', '역할', '소속팀ID', '소속채널ID']);
    }
    
    // 1. 팀 데이터 추가
    const teamSheet = spreadsheet.getSheetByName(SHEETS.teams);
    if (teamSheet && teamSheet.getLastRow() <= 1) {
      teamSheet.appendRow(['T01', '콘텐츠1팀', 'lead01@company.kr']);
      teamSheet.appendRow(['T02', '콘텐츠2팀', 'lead02@company.kr']);
      teamSheet.appendRow(['T03', '콘텐츠3팀', 'admin001@company.kr']);
      console.log('팀 데이터 추가 완료');
    }
    
    // 2. 채널 데이터 추가
    const channelSheet = spreadsheet.getSheetByName(SHEETS.channels);
    if (channelSheet && channelSheet.getLastRow() <= 1) {
      channelSheet.appendRow(['CH001', '메인 채널', 'T01', true]);
      channelSheet.appendRow(['CH002', '서브 채널', 'T01', true]);
      channelSheet.appendRow(['CH003', '뉴스 채널', 'T02', true]);
      channelSheet.appendRow(['CH004', '엔터 채널', 'T03', true]);
      channelSheet.appendRow(['CH005', '스포츠 채널', 'T03', true]);
      console.log('채널 데이터 추가 완료');
    }
    
    // 3. 사용자 데이터 추가 (시스템 관리자만)
    const userSheet = spreadsheet.getSheetByName(SHEETS.users);
    if (userSheet && userSheet.getLastRow() <= 1) {
      userSheet.appendRow(['admin001@company.kr', 'admin001@company.kr', '김관리자', '팀장', 'T03', 'Y', 'N', 'N']);
      userSheet.appendRow(['admin002@company.kr', 'admin002@company.kr', '이관리자', '관리자', 'T01,T02,T03', 'N', 'Y', 'N']);
      userSheet.appendRow(['admin003@company.kr', 'admin003@company.kr', '박관리자', '최종관리자', 'T01,T02,T03', 'N', 'N', 'Y']);
      console.log('사용자 데이터 추가 완료');
    }
    
    // 4. 담당자 데이터 추가
    if (memberSheet && memberSheet.getLastRow() <= 1) {
      memberSheet.appendRow(['M001', '홍길동', 'PD', 'T03', 'CH004']);
      memberSheet.appendRow(['M002', '김영희', '편집자', 'T03', 'CH004']);
      memberSheet.appendRow(['M003', '이철수', '작가', 'T01', 'CH001']);
      memberSheet.appendRow(['M004', '박민수', '디자이너', 'T01', 'CH001']);
      memberSheet.appendRow(['M005', '정수진', 'PD', 'T02', 'CH003']);
      console.log('담당자 데이터 추가 완료');
    }
    
    return { success: true, message: '초기 데이터 설정 완료' };
  } catch (error) {
    console.error('초기 데이터 설정 오류:', error);
    return { success: false, message: error.toString() };
  }
}

// 모든 시트 초기화 함수 (주의: 모든 데이터 삭제)
function resetAllSheets() {
  try {
    const spreadsheet = getSpreadsheet();
    
    // 각 시트 초기화
    Object.entries(SHEETS).forEach(([key, sheetName]) => {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (sheet) {
        sheet.clear();
        
        // 헤더 추가
        switch(sheetName) {
          case '사용자':
            sheet.appendRow(['사용자ID', '사용자계정', '이름', '역할', '소속팀ID', '1차평가권한', '2차평가권한', '3차평가권한']);
            break;
          case '팀':
            sheet.appendRow(['팀ID', '팀이름', '팀장ID']);
            break;
          case '채널':
            sheet.appendRow(['채널ID', '채널이름', '소속팀ID', '활성상태']);
            break;
          case '월별평가':
            sheet.appendRow(['평가ID', '평가월', '채널ID', '담당자ID', '담당자이름', '투입MM', '1차기여도', '1차성과금', 
                           '1차코멘트', '1차평가상태', '2차기여도', '2차코멘트', '2차평가상태', 
                           '사업효율', '팀장성과금', '팀장성과', '실지급성과금', '최종코멘트', '최종평가상태', '채널코멘트']);
            break;
        }
      }
    });
    
    console.log('모든 시트 초기화 완료');
    
    // 초기 데이터 추가
    return initializeTestData();
  } catch (error) {
    console.error('시트 초기화 오류:', error);
    return { success: false, message: error.toString() };
  }
}
// 디버그 함수 - 맨 아래에 추가
function debugGetDataByRole() {
  console.log('=== 디버그 시작 ===');
  
  // 테스트 데이터
  const testUserId = 'admin001@company.kr';
  const testRole = '팀장';
  const testMonth = '2025년 6월';
  
  try {
    const result = getDataByRole(testUserId, testRole, testMonth);
    console.log('getDataByRole 결과:', JSON.stringify(result));
    console.log('teams 개수:', result.teams ? result.teams.length : 'null');
    console.log('channels 개수:', result.channels ? result.channels.length : 'null');
    
    // 각 시트 데이터 직접 확인
    console.log('\n=== 시트 데이터 직접 확인 ===');
    console.log('팀 데이터:', getSheetData(SHEETS.teams));
    console.log('채널 데이터:', getSheetData(SHEETS.channels));
    console.log('사용자 데이터:', getSheetData(SHEETS.users));
    console.log('담당자 데이터:', getSheetData(SHEETS.members));
    
  } catch (error) {
    console.error('디버그 중 오류:', error);
  }
}

// 시트 구조 확인 함수
function checkSheetStructure() {
  const spreadsheet = getSpreadsheet();
  
  Object.entries(SHEETS).forEach(([key, sheetName]) => {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      console.log(`${sheetName} 시트 헤더:`, headers);
    } else {
      console.log(`${sheetName} 시트를 찾을 수 없습니다.`);
    }
  });
}

// 팀 관리자 확인 함수 추가
function getTeamManagers(teamId) {
  const users = getSheetData(SHEETS.users);
  
  // 해당 팀에 속하고 1차 평가 권한이 있는 모든 사용자
  return users.filter(u => {
    const userTeams = u['소속팀ID'] ? u['소속팀ID'].split(',').map(t => t.trim()) : [];
    return userTeams.includes(teamId) && u['1차평가권한'] === 'Y';
  });
}
