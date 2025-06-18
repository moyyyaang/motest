// 사용자 권한 확인
function checkUserPermission(userId, permissionType) {
  const users = getSheetData(SHEETS.users);
  const user = users.find(u => u['사용자ID'] === userId);
  
  if (!user) return false;
  
  switch(permissionType) {
    case '1차':
      return user['1차평가권한'] === 'Y';
    case '2차':
      return user['2차평가권한'] === 'Y';
    case '3차':
      return user['3차평가권한'] === 'Y';
    default:
      return false;
  }
}

// 팀별 사용자 목록 가져오기 (복수 팀 소속 지원)
function getUsersByTeams(teamIds) {
  const users = getSheetData(SHEETS.users);
  const teamIdArray = Array.isArray(teamIds) ? teamIds : [teamIds];
  
  return users.filter(user => {
    const userTeams = user['소속팀ID'] ? user['소속팀ID'].split(',').map(t => t.trim()) : [];
    return userTeams.some(teamId => teamIdArray.includes(teamId));
  });
}// 평가ID 일괄 생성 (여러 개 한번에)
function generateEvaluationIds(evaluationMonth, count) {
  const ids = [];
  try {
    for (let i = 0; i < count; i++) {
      ids.push(generateEvaluationId(evaluationMonth));
    }
  } catch (error) {
    // 오류 시 타임스탬프 기반 ID 생성
    for (let i = 0; i < count; i++) {
      ids.push('EVAL-' + Date.now() + '-' + i);
    }
  }
  return ids;
}// Google Apps Script 백엔드 코드

// 스프레드시트 ID와 시트 이름 설정
const SPREADSHEET_ID = '1Ykzi-v6zPW4hgUaRYsSloStlWPQN_HgDoPvMVcGmdmc';
const SHEETS = {
  users: '사용자',
  teams: '팀',
  channels: '채널',
  evaluations: '월별평가'
};

// 스프레드시트 접근 헬퍼
function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (error) {
    console.error('스프레드시트 접근 오류:', error);
    // 권한 문제 시 활성 스프레드시트 사용 시도
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

// ============ API 함수들 ============

// 현재 사용자 정보 가져오기 (권한 정보 포함)
function getCurrentUser(email) {
  const users = getSheetData(SHEETS.users);
  const user = users.find(user => user['사용자ID'] === email);
  
  if (user) {
    // 권한 정보 추가 (권한 컬럼이 없을 수도 있으므로 기본값 처리)
    user['평가권한'] = [];
    
    // 권한 컬럼이 있는 경우
    if (user['1차평가권한'] === 'Y' || (user['1차평가권한'] === undefined && user['역할'] === '팀장')) {
      user['평가권한'].push('1차');
    }
    if (user['2차평가권한'] === 'Y' || (user['2차평가권한'] === undefined && user['역할'] === '관리자')) {
      user['평가권한'].push('2차');
    }
    if (user['3차평가권한'] === 'Y' || (user['3차평가권한'] === undefined && user['역할'] === '최종관리자')) {
      user['평가권한'].push('3차');
    }
    
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
      users: []
    };
    
    // 데이터 로드
    const users = getSheetData(SHEETS.users);
    const teams = getSheetData(SHEETS.teams);
    const channels = getSheetData(SHEETS.channels);
    const evaluations = getSheetData(SHEETS.evaluations);
    
    console.log('로드된 데이터:', {
      users: users.length,
      teams: teams.length,
      channels: channels.length,
      evaluations: evaluations.length
    });
    
    // 평가월 필터링
    const monthEvaluations = evaluationMonth ? 
      evaluations.filter(e => e['평가월'] === evaluationMonth) : 
      evaluations;
    
    if (role === '팀장') {
      // 팀장은 자신의 팀 데이터만
      const currentUser = users.find(u => u['사용자ID'] === userId);
      if (currentUser && currentUser['소속팀ID']) {
        // 복수 팀 소속 지원
        const userTeams = currentUser['소속팀ID'].split(',').map(t => t.trim());
        console.log('팀장의 소속팀들:', userTeams);
        
        result.teams = teams.filter(t => userTeams.includes(t['팀ID']));
        result.channels = channels.filter(c => userTeams.includes(c['소속팀ID']) && c['활성상태'] !== false);
        result.users = users.filter(u => {
          const uTeams = u['소속팀ID'] ? u['소속팀ID'].split(',').map(t => t.trim()) : [];
          return uTeams.some(t => userTeams.includes(t));
        });
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
      result.evaluations = monthEvaluations;
    }
    
    console.log('필터링된 결과:', {
      teams: result.teams.length,
      channels: result.channels.length,
      users: result.users.length,
      evaluations: result.evaluations.length
    });
    
    return result;
  } catch (error) {
    console.error('getDataByRole 오류:', error);
    throw new Error('데이터를 불러올 수 없습니다: ' + error.toString());
  }
}

// 평가 데이터 저장
function saveEvaluation(evaluationData) {
  try {
    const spreadsheet = getSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEETS.evaluations);
    
    if (!sheet) {
      throw new Error('평가 시트를 찾을 수 없습니다.');
    }
    
    // 헤더가 없으면 생성
    if (sheet.getLastRow() === 0) {
      const headers = ['평가ID', '평가월', '채널ID', '담당자ID', '담당자이름', '투입MM', '1차기여도', '1차성과금', 
                       '1차코멘트', '1차평가상태', '2차기여도', '2차코멘트', '2차평가상태', 
                       '사업효율', '팀장성과금', '팀장성과', '실지급성과금', '최종코멘트', '최종평가상태', '채널코멘트'];
      sheet.appendRow(headers);
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 기존 데이터 확인
    const evaluationId = evaluationData['평가ID'];
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === evaluationId) {
        rowIndex = i + 1; // 1-based index
        break;
      }
    }
    
    // 기존 데이터가 있으면 병합
    if (rowIndex > 0) {
      const existingData = {};
      headers.forEach((header, index) => {
        existingData[header] = values[rowIndex - 1][index];
      });
      // 새 데이터로 업데이트 (기존 데이터 유지하면서 새 데이터만 덮어쓰기)
      Object.keys(evaluationData).forEach(key => {
        if (evaluationData[key] !== undefined && evaluationData[key] !== '') {
          existingData[key] = evaluationData[key];
        }
      });
      evaluationData = existingData;
    }
    
    // 데이터 배열 생성
    const rowData = headers.map(header => evaluationData[header] || '');
    
    if (rowIndex > 0) {
      // 기존 데이터 업데이트
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      // 새 데이터 추가
      sheet.appendRow(rowData);
    }
    
    return { success: true, message: '저장되었습니다.' };
  } catch (error) {
    console.error('saveEvaluation 오류:', error);
    return { success: false, message: error.toString() };
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
    const defaultHeaders = ['평가ID', '평가월', '채널ID', '담당자ID', '담당자이름', '투입MM', '1차기여도', '1차성과금', 
                           '1차코멘트', '1차평가상태', '2차기여도', '2차코멘트', '2차평가상태', 
                           '사업효율', '팀장성과금', '팀장성과', '실지급성과금', '최종코멘트', '최종평가상태', '채널코멘트'];
    
    let headers;
    if (sheet.getLastRow() === 0) {
      // 헤더가 없으면 생성
      sheet.appendRow(defaultHeaders);
      headers = defaultHeaders;
      console.log('헤더 생성됨');
    } else {
      // 기존 헤더 읽기
      const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
      if (headerRange.getValues()[0][0]) {
        headers = headerRange.getValues()[0];
      } else {
        // 헤더가 비어있으면 기본 헤더 사용
        sheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
        headers = defaultHeaders;
      }
    }
    
    console.log('헤더:', headers);
    
    const dataRange = sheet.getDataRange();
    const existingData = dataRange.getValues();
    
    // 기존 데이터를 Map으로 변환 (빠른 검색을 위해)
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
        const rowData = headers.map(header => mergedData[header] || '');
        updates.push({ row: existing.rowIndex, data: rowData });
      } else {
        const rowData = headers.map(header => evalData[header] || '');
        appends.push(rowData);
      }
    });
    
    console.log('업데이트:', updates.length, '추가:', appends.length);
    
    // 업데이트 실행
    updates.forEach(update => {
      sheet.getRange(update.row, 1, 1, update.data.length).setValues([update.data]);
    });
    
    // 추가 실행
    if (appends.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, appends.length, appends[0].length).setValues(appends);
    }
    
    return { success: true, message: `${updates.length}개 업데이트, ${appends.length}개 추가 완료` };
  } catch (error) {
    console.error('saveEvaluationBatch 오류:', error);
    console.error('오류 상세:', error.stack);
    return { success: false, message: error.toString() };
  }
}

// 새 평가ID 생성 (EVAL-YYYYMM-001 형식)
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
    
    const newId = 'CH' + (maxNum + 1);
    sheet.appendRow([newId, channelData.name, channelData.teamId, true]);
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
    
    const newId = 'T' + new Date().getTime();
    sheet.appendRow([newId, teamData.name, teamData.leaderId || '']);
    return { success: true, teamId: newId };
  } catch (error) {
    console.error('addTeam 오류:', error);
    return { success: false, message: error.toString() };
  }
}

// 사용자 추가
function addUser(userData) {
  try {
    const spreadsheet = getSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEETS.users);
    
    if (!sheet) {
      throw new Error('사용자 시트를 찾을 수 없습니다.');
    }
    
    // 헤더가 없으면 생성
    if (sheet.getLastRow() === 0) {
      const headers = ['사용자ID', '이름', '역할', '소속팀ID', '1차평가권한', '2차평가권한', '3차평가권한'];
      sheet.appendRow(headers);
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 권한 설정 (전달된 값이 있으면 사용, 없으면 역할 기반)
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
    
    // 복수 팀 소속 지원 - 배열이면 쉼표로 구분하여 저장
    const teamIds = Array.isArray(userData.teamId) ? 
      userData.teamId.join(',') : 
      userData.teamId;
    
    const rowData = headers.map(header => {
      switch(header) {
        case '사용자ID': return userData.userId;
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

// 평가 월 목록 가져오기
function getEvaluationMonths() {
  try {
    const evaluations = getSheetData(SHEETS.evaluations);
    if (!evaluations || evaluations.length === 0) {
      console.log('평가 데이터가 없습니다.');
      return [];
    }
    
    const months = [...new Set(evaluations.map(e => e['평가월']).filter(m => m))]
      .sort()
      .reverse();
    
    console.log('평가월 목록:', months);
    return months;
  } catch (error) {
    console.error('getEvaluationMonths 오류:', error);
    return [];
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

// 디버그용 - 실행해서 로그 확인
function debugTest() {
  console.log('=== 연결 테스트 ===');
  console.log(testConnection());
  
  console.log('\n=== 평가월 목록 ===');
  console.log(getEvaluationMonths());
  
  console.log('\n=== 데이터 로드 테스트 ===');
  const testData = getDataByRole('member1@awesomeent.kr', '팀장', '2025년 6월');
  console.log('로드된 데이터 요약:', {
    teams: testData.teams.length,
    channels: testData.channels.length,
    evaluations: testData.evaluations.length,
    users: testData.users.length
  });
  
  // 복수 팀 소속 테스트
  console.log('\n=== 복수 팀 소속 테스트 ===');
  const users = getSheetData(SHEETS.users);
  users.forEach(user => {
    if (user['소속팀ID'] && user['소속팀ID'].includes(',')) {
      console.log(`${user['이름']}님은 복수 팀 소속:`, user['소속팀ID']);
    }
  });
}

// 평가 시트 초기화 (테스트용)
function initEvaluationSheet() {
  try {
    const spreadsheet = getSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SHEETS.evaluations);
    
    if (!sheet) {
      // 시트가 없으면 생성
      spreadsheet.insertSheet(SHEETS.evaluations);
      sheet = spreadsheet.getSheetByName(SHEETS.evaluations);
    }
    
    // 헤더 설정
    const headers = ['평가ID', '평가월', '채널ID', '담당자ID', '담당자이름', '투입MM', '1차기여도', '1차성과금', 
                     '1차코멘트', '1차평가상태', '2차기여도', '2차코멘트', '2차평가상태', 
                     '사업효율', '팀장성과금', '팀장성과', '실지급성과금', '최종코멘트', '최종평가상태', '채널코멘트'];
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      console.log('평가 시트 헤더 생성 완료');
    } else {
      console.log('평가 시트 이미 존재');
    }
    
    return { success: true, message: '평가 시트 초기화 완료' };
  } catch (error) {
    console.error('평가 시트 초기화 오류:', error);
    return { success: false, message: error.toString() };
  }
}

// 간단한 테스트 함수
function simpleTest() {
  return "서버 연결 성공!";
}
