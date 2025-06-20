// Google Apps Script 백엔드 코드 (개선버전)

// 스프레드시트 ID와 시트 이름 설정
const SPREADSHEET_ID = '1Ykzi-v6zPW4hgUaRYsSloStlWPQN_HgDoPvMVcGmdmc';
const SHEETS = {
  users: '사용자',
  teams: '팀',
  channels: '채널',
  evaluations: '월별평가',
  progress: '평가진행상태',
  members: '담당자'  // 담당자 전용 시트 추가
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
