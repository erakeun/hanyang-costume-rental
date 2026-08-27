/**
 * 하냥이 인형탈 대여 시스템 v1.0
 * - Apps Script Web App 단독 구성
 * - Google Sheets DB
 * - 신청 접수 / 중복 일정 차단 / 관리자 승인·거절 / 결과 메일
 *
 * 최초 1회:
 * 1) setupSystem() 실행
 * 2) 웹 앱으로 배포
 */

const CONFIG = Object.freeze({
  SITE_TITLE: '하냥이 인형탈 대여',
  DB_NAME: '하냥이 인형탈 대여 관리',
  SHEET_NAME: '신청DB',
  GUIDE_SHEET_NAME: '운영안내',
  ADMIN_EMAIL: 'hyerica@hanyang.ac.kr',
  SENDER_NAME: '한양대학교 ERICA 하냥이 인형탈 대여',
  TZ: 'Asia/Seoul',
  HEIGHT_MIN: 160,
  HEIGHT_MAX: 169,
  STATUS_PENDING: '접수',
  STATUS_APPROVED: '승인',
  STATUS_REJECTED: '거절',
  LOCK_TIMEOUT_MS: 15000
});

const HEADERS = Object.freeze([
  '신청번호',
  '접수일시',
  '상태',
  '신청기관 및 부서',
  '담당자 성명',
  '이메일',
  '연락처',
  '대여 시작일',
  '대여 종료일',
  '사용 목적',
  '사용 장소',
  '착용자 신장(cm)',
  '개인 신청 여부',
  '상업적 사용 여부',
  '야간 사용 여부',
  '우천 사용 여부',
  '동행자 여부',
  '교외 사용 여부',
  '키 경고 확인',
  '유의사항 동의',
  '검토 메모',
  '승인토큰 해시',
  '처리일시',
  '처리자',
  '거절사유'
]);

/* =========================================================
 * 최초 설치
 * ========================================================= */

function setupSystem() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty('SPREADSHEET_ID');
  let ss;

  if (spreadsheetId) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  } else {
    ss = SpreadsheetApp.create(CONFIG.DB_NAME);
    ss.setSpreadsheetTimeZone(CONFIG.TZ);
    props.setProperty('SPREADSHEET_ID', ss.getId());
    spreadsheetId = ss.getId();
  }

  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName(CONFIG.SHEET_NAME);
  }

  setupDbSheet_(sheet);

  let guide = ss.getSheetByName(CONFIG.GUIDE_SHEET_NAME);
  if (!guide) guide = ss.insertSheet(CONFIG.GUIDE_SHEET_NAME);
  setupGuideSheet_(guide);

  Logger.log('관리 스프레드시트: ' + ss.getUrl());
  Logger.log('관리 메일: ' + CONFIG.ADMIN_EMAIL);

  return {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetUrl: ss.getUrl(),
    adminEmail: CONFIG.ADMIN_EMAIL
  };
}

function setupDbSheet_(sheet) {
  sheet.clear();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);

  const header = sheet.getRange(1, 1, 1, HEADERS.length);
  header.setFontWeight('bold')
    .setBackground('#dbeafe')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.setRowHeight(1, 34);

  const widths = [
    165, 150, 80, 190, 120, 190, 130, 110, 110, 260, 220, 110,
    110, 120, 110, 110, 110, 110, 110, 110, 220, 240, 150, 160, 260
  ];
  widths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));

  if (sheet.getMaxRows() > 1) {
    sheet.getRange(2, 2, sheet.getMaxRows() - 1, 1)
      .setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange(2, 8, sheet.getMaxRows() - 1, 2)
      .setNumberFormat('yyyy-mm-dd');
    sheet.getRange(2, 23, sheet.getMaxRows() - 1, 1)
      .setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }

  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 2), HEADERS.length).createFilter();
}

function setupGuideSheet_(sheet) {
  sheet.clear();

  const rows = [
    ['하냥이 인형탈 대여 시스템 운영안내', ''],
    ['관리 메일', CONFIG.ADMIN_EMAIL],
    ['권장 신장', `${CONFIG.HEIGHT_MIN}~${CONFIG.HEIGHT_MAX}cm`],
    ['자동 차단', '개인 신청 / 상업적 사용 / 야간 사용 / 우천 사용 / 동행자 없음'],
    ['경고 후 신청 가능', '권장 신장 범위 이탈 / 교외 사용'],
    ['일정 차단 상태', '접수 또는 승인 상태와 날짜가 겹치면 신규 신청 차단'],
    ['승인 처리', '관리 메일의 승인·거절 검토 버튼 → 확인 페이지 → 최종 처리'],
    ['주의', '승인·거절 상태는 신청DB에서 임의로 수정하지 않는 것을 권장합니다.']
  ];

  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange(1, 1, 1, 2)
    .merge()
    .setFontWeight('bold')
    .setFontSize(16)
    .setBackground('#dbeafe');
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 620);
  sheet.getRange(1, 1, rows.length, 2).setVerticalAlignment('middle').setWrap(true);
}

/* =========================================================
 * 웹 앱 엔트리
 * ========================================================= */

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.mode === 'review' && p.token) {
    return renderReviewPage_(p.token, p.intent || '');
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(CONFIG.SITE_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const p = (e && e.parameter) || {};
    const action = String(p.action || '').toLowerCase();
    const token = String(p.token || '');

    if (!token || !['approve', 'reject'].includes(action)) {
      throw new Error('잘못된 처리 요청입니다.');
    }

    let reason = '';
    if (action === 'reject') {
      const reasonType = cleanText_(p.reasonType || '', 100);
      const customReason = cleanText_(p.customReason || '', 500);

      if (reasonType === '기타') {
        reason = customReason;
      } else {
        reason = reasonType;
        if (customReason) reason += ` - ${customReason}`;
      }

      if (!reason) {
        throw new Error('거절 사유를 입력해 주세요.');
      }
    }

    const result = processDecision_(token, action, reason);
    return renderDecisionResultPage_(result);
  } catch (err) {
    return renderSimplePage_(
      '처리할 수 없습니다',
      escapeHtml_(err && err.message ? err.message : '알 수 없는 오류가 발생했습니다.'),
      'error'
    );
  }
}

/* =========================================================
 * 공개 페이지 서버 함수
 * ========================================================= */

function getBootstrap() {
  ensureSetup_();

  const now = new Date();
  return {
    siteTitle: CONFIG.SITE_TITLE,
    currentYear: Number(Utilities.formatDate(now, CONFIG.TZ, 'yyyy')),
    currentMonth: Number(Utilities.formatDate(now, CONFIG.TZ, 'M')),
    today: Utilities.formatDate(now, CONFIG.TZ, 'yyyy-MM-dd'),
    heightMin: CONFIG.HEIGHT_MIN,
    heightMax: CONFIG.HEIGHT_MAX
  };
}

function getCalendarMonth(year, month) {
  ensureSetup_();

  year = Number(year);
  month = Number(month);

  if (!Number.isInteger(year) || !Number.isInteger(month) || month < 1 || month > 12) {
    throw new Error('달력 날짜가 올바르지 않습니다.');
  }

  const startOfMonth = new Date(year, month - 1, 1);
  const endOfMonth = new Date(year, month, 0);
  const sheet = getDbSheet_();

  const statusByDate = {};
  const lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

    rows.forEach(row => {
      const status = String(row[2] || '');
      if (![CONFIG.STATUS_PENDING, CONFIG.STATUS_APPROVED].includes(status)) return;

      const rowStart = normalizeDate_(row[7]);
      const rowEnd = normalizeDate_(row[8]);
      if (!rowStart || !rowEnd) return;
      if (rowEnd < startOfMonth || rowStart > endOfMonth) return;

      let cursor = new Date(Math.max(rowStart.getTime(), startOfMonth.getTime()));
      const finish = new Date(Math.min(rowEnd.getTime(), endOfMonth.getTime()));

      while (cursor <= finish) {
        const key = formatDateKey_(cursor);
        const nextStatus = status === CONFIG.STATUS_APPROVED ? 'APPROVED' : 'PENDING';
        const current = statusByDate[key];

        if (!current || nextStatus === 'APPROVED') {
          statusByDate[key] = nextStatus;
        }

        cursor.setDate(cursor.getDate() + 1);
      }
    });
  }

  return {
    year,
    month,
    days: Object.keys(statusByDate).sort().map(date => ({
      date,
      status: statusByDate[date]
    }))
  };
}

function checkAvailability(startDate, endDate) {
  ensureSetup_();

  const start = parseDateInput_(startDate);
  const end = parseDateInput_(endDate);

  if (!start || !end) {
    return { ok: false, available: false, message: '대여 시작일과 종료일을 확인해 주세요.' };
  }
  if (end < start) {
    return { ok: false, available: false, message: '대여 종료일은 시작일보다 빠를 수 없습니다.' };
  }

  const conflict = findConflict_(start, end, '');
  return conflict
    ? {
        ok: true,
        available: false,
        message: '선택한 기간에 이미 접수 또는 승인된 대여 일정이 있습니다.'
      }
    : {
        ok: true,
        available: true,
        message: '현재 기준으로 신청 가능한 기간입니다.'
      };
}

function submitApplication(payload) {
  ensureSetup_();

  const data = validateApplication_(payload || {});
  const lock = LockService.getScriptLock();

  lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);

  let record;
  try {
    const conflict = findConflict_(data.startDate, data.endDate, '');
    if (conflict) {
      throw new Error('선택한 기간에 방금 다른 신청이 접수되었거나 이미 확정된 일정이 있습니다. 날짜를 다시 확인해 주세요.');
    }

    const applicationId = createApplicationId_();
    const token = createToken_();
    const tokenHash = hashToken_(token);
    const submittedAt = new Date();

    const reviewNotes = [];
    if (data.height < CONFIG.HEIGHT_MIN || data.height > CONFIG.HEIGHT_MAX) {
      reviewNotes.push(`권장 신장(${CONFIG.HEIGHT_MIN}~${CONFIG.HEIGHT_MAX}cm) 범위 이탈`);
    }
    if (data.offCampus) {
      reviewNotes.push('교외 사용 예정 - 별도 승인 검토 필요');
    }

    const row = [
      safeSheetValue_(applicationId),
      submittedAt,
      CONFIG.STATUS_PENDING,
      safeSheetValue_(data.organization),
      safeSheetValue_(data.managerName),
      safeSheetValue_(data.email),
      safeSheetValue_(data.phone),
      data.startDate,
      data.endDate,
      safeSheetValue_(data.purpose),
      safeSheetValue_(data.location),
      data.height,
      data.individual ? '예' : '아니오',
      data.commercial ? '예' : '아니오',
      data.night ? '예' : '아니오',
      data.rain ? '예' : '아니오',
      data.companion ? '있음' : '없음',
      data.offCampus ? '예' : '아니오',
      data.heightAck ? '예' : '해당없음',
      data.termsAccepted ? '예' : '아니오',
      safeSheetValue_(reviewNotes.join(' / ')),
      tokenHash,
      '',
      '',
      ''
    ];

    const sheet = getDbSheet_();
    sheet.appendRow(row);
    const rowNumber = sheet.getLastRow();

    sheet.getRange(rowNumber, 2).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange(rowNumber, 8, 1, 2).setNumberFormat('yyyy-mm-dd');
    sheet.getRange(rowNumber, 23).setNumberFormat('yyyy-mm-dd hh:mm:ss');

    record = {
      rowNumber,
      applicationId,
      token,
      submittedAt,
      organization: data.organization,
      managerName: data.managerName,
      email: data.email,
      phone: data.phone,
      startDate: data.startDate,
      endDate: data.endDate,
      purpose: data.purpose,
      location: data.location,
      height: data.height,
      offCampus: data.offCampus,
      reviewNotes
    };
  } finally {
    lock.releaseLock();
  }

  let mailWarning = '';
  try {
    sendAdminRequest_(record);
    sendApplicantReceipt_(record);
  } catch (mailErr) {
    mailWarning = '신청은 저장되었지만 메일 알림 전송 중 오류가 발생했습니다. 담당부서에서 신청DB를 확인해 주세요.';
    appendReviewNote_(record.applicationId, '메일 발송 오류: ' + String(mailErr && mailErr.message ? mailErr.message : mailErr));
  }

  return {
    ok: true,
    applicationId: record.applicationId,
    message: '대여 신청이 접수되었습니다. 담당자 검토 후 결과가 이메일로 발송됩니다.',
    mailWarning
  };
}

/* =========================================================
 * 신청 검증
 * ========================================================= */

function validateApplication_(p) {
  const organization = cleanText_(p.organization, 120);
  const managerName = cleanText_(p.managerName, 50);
  const email = cleanText_(p.email, 180);
  const phone = cleanText_(p.phone, 40);
  const purpose = cleanText_(p.purpose, 800);
  const location = cleanText_(p.location, 300);

  const individual = toBool_(p.individual);
  const commercial = toBool_(p.commercial);
  const night = toBool_(p.night);
  const rain = toBool_(p.rain);
  const companion = toBool_(p.companion);
  const offCampus = toBool_(p.offCampus);
  const heightAck = toBool_(p.heightAck);
  const termsAccepted = toBool_(p.termsAccepted);

  const height = Number(p.height);
  const startDate = parseDateInput_(p.startDate);
  const endDate = parseDateInput_(p.endDate);

  if (!organization) throw new Error('신청 기관 및 부서를 입력해 주세요.');
  if (!managerName) throw new Error('담당자 성명을 입력해 주세요.');
  if (!isValidEmail_(email)) throw new Error('이메일 주소를 확인해 주세요.');
  if (!phone) throw new Error('연락처를 입력해 주세요.');
  if (!purpose) throw new Error('사용 목적 및 상세 내용을 입력해 주세요.');
  if (!location) throw new Error('사용 장소를 입력해 주세요.');

  if (!startDate || !endDate) throw new Error('대여 시작일과 종료일을 입력해 주세요.');
  if (endDate < startDate) throw new Error('대여 종료일은 시작일보다 빠를 수 없습니다.');

  const today = parseDateInput_(Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd'));
  if (startDate < today) throw new Error('과거 날짜로는 신청할 수 없습니다.');

  if (!Number.isFinite(height) || height < 100 || height > 220) {
    throw new Error('착용자 신장을 올바르게 입력해 주세요.');
  }

  // 자동 차단 규칙
  if (individual) throw new Error('개인 단위 대여는 신청할 수 없습니다. 단체 명의로 신청해 주세요.');
  if (commercial) throw new Error('상업적 목적의 사용은 신청할 수 없습니다.');
  if (night) throw new Error('야간 사용은 신청할 수 없습니다.');
  if (rain) throw new Error('우천 환경에서의 사용은 신청할 수 없습니다.');
  if (!companion) throw new Error('인형탈 사용 시 동행자가 반드시 필요합니다.');

  // 경고 후 진행 규칙
  if ((height < CONFIG.HEIGHT_MIN || height > CONFIG.HEIGHT_MAX) && !heightAck) {
    throw new Error(`착용자 신장이 권장 범위(${CONFIG.HEIGHT_MIN}~${CONFIG.HEIGHT_MAX}cm)를 벗어났습니다. 경고 내용을 확인해 주세요.`);
  }

  if (!termsAccepted) {
    throw new Error('대여 및 사용 유의사항에 동의해 주세요.');
  }

  return {
    organization,
    managerName,
    email,
    phone,
    purpose,
    location,
    individual,
    commercial,
    night,
    rain,
    companion,
    offCampus,
    heightAck,
    termsAccepted,
    height,
    startDate,
    endDate
  };
}

/* =========================================================
 * 승인 / 거절
 * ========================================================= */

function renderReviewPage_(token, intent) {
  ensureSetup_();

  try {
    const lookup = findRecordByToken_(token);
    if (!lookup) {
      return renderSimplePage_('신청을 찾을 수 없습니다', '승인 링크가 잘못되었거나 더 이상 사용할 수 없습니다.', 'error');
    }

    const record = lookup.record;
    if (record.status !== CONFIG.STATUS_PENDING) {
      return renderSimplePage_(
        '이미 처리된 신청입니다',
        `현재 상태: <strong>${escapeHtml_(record.status)}</strong>`,
        'info'
      );
    }

    const webAppUrl = ScriptApp.getService().getUrl() || '';
    const intentText = intent === 'reject'
      ? '거절 검토'
      : intent === 'approve'
        ? '승인 검토'
        : '신청 검토';

    const reviewNoteHtml = record.reviewNote
      ? `<div class="warning"><strong>검토 메모</strong><br>${escapeHtml_(record.reviewNote)}</div>`
      : '<div class="good">자동 사전검사에서 추가 경고 사항이 없습니다.</div>';

    const html = `
<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>${escapeHtml_(CONFIG.SITE_TITLE)} - ${escapeHtml_(intentText)}</title>
  <style>
    *{box-sizing:border-box}
    body{margin:0;background:#f4f7fb;color:#172033;font-family:Arial,"Noto Sans KR",sans-serif}
    .wrap{max-width:760px;margin:0 auto;padding:32px 18px 60px}
    .card{background:#fff;border:1px solid #dbe2ea;border-radius:18px;padding:24px;margin-bottom:18px;box-shadow:0 10px 28px rgba(15,23,42,.06)}
    h1{font-size:28px;margin:0 0 8px}
    h2{font-size:19px;margin:0 0 16px}
    .sub{color:#667085;margin-bottom:24px}
    .grid{display:grid;grid-template-columns:150px 1fr;border-top:1px solid #e5e7eb}
    .grid div{padding:11px 8px;border-bottom:1px solid #e5e7eb}
    .label{font-weight:700;color:#475467}
    .good,.warning{padding:14px;border-radius:12px;margin-top:16px}
    .good{background:#ecfdf3;border:1px solid #abefc6;color:#067647}
    .warning{background:#fffaeb;border:1px solid #fedf89;color:#93370d}
    .actions{display:grid;grid-template-columns:1fr 1fr;gap:18px}
    .approve{border-top:5px solid #12b76a}
    .reject{border-top:5px solid #f04438}
    button{width:100%;border:0;border-radius:12px;padding:14px 18px;font-size:16px;font-weight:800;cursor:pointer}
    .btn-approve{background:#12b76a;color:#fff}
    .btn-reject{background:#f04438;color:#fff}
    select,textarea{width:100%;border:1px solid #d0d5dd;border-radius:10px;padding:12px;font:inherit;margin:8px 0}
    textarea{min-height:100px;resize:vertical}
    .note{font-size:13px;color:#667085;line-height:1.55}
    @media(max-width:640px){
      .actions{grid-template-columns:1fr}
      .grid{grid-template-columns:110px 1fr}
      .wrap{padding:18px 12px 40px}
    }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <h1>하냥이 인형탈 대여 ${escapeHtml_(intentText)}</h1>
      <div class="sub">메일의 버튼은 검토 페이지로만 연결됩니다. 실제 상태 변경은 아래 최종 버튼을 눌렀을 때만 이루어집니다.</div>
      <div class="grid">
        <div class="label">신청번호</div><div>${escapeHtml_(record.applicationId)}</div>
        <div class="label">신청단체</div><div>${escapeHtml_(record.organization)}</div>
        <div class="label">담당자</div><div>${escapeHtml_(record.managerName)}</div>
        <div class="label">이메일</div><div>${escapeHtml_(record.email)}</div>
        <div class="label">연락처</div><div>${escapeHtml_(record.phone)}</div>
        <div class="label">대여기간</div><div>${escapeHtml_(record.startDateText)} ~ ${escapeHtml_(record.endDateText)}</div>
        <div class="label">사용목적</div><div>${nl2br_(record.purpose)}</div>
        <div class="label">사용장소</div><div>${escapeHtml_(record.location)}</div>
        <div class="label">착용자 신장</div><div>${escapeHtml_(String(record.height))}cm</div>
        <div class="label">교외 사용</div><div>${escapeHtml_(record.offCampus)}</div>
      </div>
      ${reviewNoteHtml}
    </div>

    <div class="actions">
      <div class="card approve">
        <h2>승인</h2>
        <p class="note">승인 시 신청자에게 승인 메일이 발송되고, 해당 기간은 달력에서 대여 확정 상태로 표시됩니다.</p>
        <form method="post" action="${escapeAttr_(webAppUrl)}">
          <input type="hidden" name="action" value="approve">
          <input type="hidden" name="token" value="${escapeAttr_(token)}">
          <button class="btn-approve" type="submit">최종 승인</button>
        </form>
      </div>

      <div class="card reject">
        <h2>거절</h2>
        <p class="note">거절 사유는 신청자에게 그대로 전달됩니다.</p>
        <form method="post" action="${escapeAttr_(webAppUrl)}">
          <input type="hidden" name="action" value="reject">
          <input type="hidden" name="token" value="${escapeAttr_(token)}">
          <select name="reasonType" required>
            <option value="">거절 사유 선택</option>
            <option>해당 기간 대여 불가</option>
            <option>사용 목적 부적합</option>
            <option>안전기준 미충족</option>
            <option>신청정보 미비</option>
            <option>기타</option>
          </select>
          <textarea name="customReason" maxlength="500" placeholder="필요 시 상세 사유를 입력해 주세요."></textarea>
          <button class="btn-reject" type="submit">거절 확정</button>
        </form>
      </div>
    </div>
  </div>
</body>
</html>`;

    return HtmlService.createHtmlOutput(html).setTitle(`${CONFIG.SITE_TITLE} - 신청 검토`);
  } catch (err) {
    return renderSimplePage_('검토 페이지 오류', escapeHtml_(err.message || String(err)), 'error');
  }
}

function processDecision_(token, action, reason) {
  const lock = LockService.getScriptLock();
  lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);

  let record;
  let newStatus;

  try {
    const lookup = findRecordByToken_(token);
    if (!lookup) throw new Error('신청을 찾을 수 없습니다.');

    record = lookup.record;
    if (record.status !== CONFIG.STATUS_PENDING) {
      throw new Error(`이미 처리된 신청입니다. 현재 상태: ${record.status}`);
    }

    if (action === 'approve') {
      const conflict = findConflict_(record.startDate, record.endDate, record.applicationId);
      if (conflict) {
        throw new Error(`다른 신청(${conflict.applicationId})과 일정이 겹쳐 승인할 수 없습니다.`);
      }
      newStatus = CONFIG.STATUS_APPROVED;
    } else {
      newStatus = CONFIG.STATUS_REJECTED;
    }

    const sheet = getDbSheet_();
    const row = lookup.rowNumber;

    const processor = Session.getActiveUser().getEmail() || '승인 링크 처리';
    sheet.getRange(row, 3).setValue(newStatus);
    sheet.getRange(row, 23).setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    sheet.getRange(row, 24).setValue(safeSheetValue_(processor));
    sheet.getRange(row, 25).setValue(action === 'reject' ? safeSheetValue_(reason) : '');

    record.status = newStatus;
    record.reason = reason || '';
  } finally {
    lock.releaseLock();
  }

  let mailWarning = '';
  try {
    sendApplicantDecision_(record);
  } catch (mailErr) {
    mailWarning = '상태 처리는 완료되었지만 신청자 결과 메일 발송에 실패했습니다.';
    appendReviewNote_(record.applicationId, '결과 메일 발송 오류: ' + String(mailErr && mailErr.message ? mailErr.message : mailErr));
  }

  return {
    ok: true,
    status: newStatus,
    applicationId: record.applicationId,
    organization: record.organization,
    email: record.email,
    mailWarning
  };
}

/* =========================================================
 * DB 조회
 * ========================================================= */

function findConflict_(startDate, endDate, ignoreApplicationId) {
  const sheet = getDbSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const applicationId = String(row[0] || '');
    const status = String(row[2] || '');

    if (ignoreApplicationId && applicationId === ignoreApplicationId) continue;
    if (![CONFIG.STATUS_PENDING, CONFIG.STATUS_APPROVED].includes(status)) continue;

    const rowStart = normalizeDate_(row[7]);
    const rowEnd = normalizeDate_(row[8]);
    if (!rowStart || !rowEnd) continue;

    if (startDate <= rowEnd && endDate >= rowStart) {
      return {
        rowNumber: i + 2,
        applicationId,
        status,
        startDate: rowStart,
        endDate: rowEnd
      };
    }
  }

  return null;
}

function findRecordByToken_(token) {
  const hash = hashToken_(token);
  const sheet = getDbSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][21] || '') === hash) {
      return {
        rowNumber: i + 2,
        record: rowToRecord_(rows[i])
      };
    }
  }

  return null;
}

function rowToRecord_(row) {
  return {
    applicationId: String(row[0] || ''),
    submittedAt: row[1],
    status: String(row[2] || ''),
    organization: String(row[3] || ''),
    managerName: String(row[4] || ''),
    email: String(row[5] || ''),
    phone: String(row[6] || ''),
    startDate: normalizeDate_(row[7]),
    endDate: normalizeDate_(row[8]),
    startDateText: formatDateKey_(normalizeDate_(row[7])),
    endDateText: formatDateKey_(normalizeDate_(row[8])),
    purpose: String(row[9] || ''),
    location: String(row[10] || ''),
    height: Number(row[11] || 0),
    individual: String(row[12] || ''),
    commercial: String(row[13] || ''),
    night: String(row[14] || ''),
    rain: String(row[15] || ''),
    companion: String(row[16] || ''),
    offCampus: String(row[17] || ''),
    heightAck: String(row[18] || ''),
    termsAccepted: String(row[19] || ''),
    reviewNote: String(row[20] || ''),
    reason: String(row[24] || '')
  };
}

function appendReviewNote_(applicationId, text) {
  try {
    const sheet = getDbSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0] === applicationId) {
        const cell = sheet.getRange(i + 2, 21);
        const old = String(cell.getValue() || '');
        cell.setValue(safeSheetValue_(old ? `${old} / ${text}` : text));
        return;
      }
    }
  } catch (err) {
    console.error(err);
  }
}

/* =========================================================
 * 메일
 * ========================================================= */

function sendAdminRequest_(r) {
  const webAppUrl = ScriptApp.getService().getUrl();
  if (!webAppUrl) throw new Error('웹 앱 배포 URL을 확인할 수 없습니다. 먼저 웹 앱으로 배포해 주세요.');

  const approveUrl = `${webAppUrl}?mode=review&intent=approve&token=${encodeURIComponent(r.token)}`;
  const rejectUrl = `${webAppUrl}?mode=review&intent=reject&token=${encodeURIComponent(r.token)}`;

  const notes = r.reviewNotes.length
    ? `<div style="margin:16px 0;padding:14px;border:1px solid #fedf89;background:#fffaeb;border-radius:10px;color:#93370d">
         <b>⚠ 검토 필요</b><br>${r.reviewNotes.map(escapeHtml_).join('<br>')}
       </div>`
    : `<div style="margin:16px 0;padding:14px;border:1px solid #abefc6;background:#ecfdf3;border-radius:10px;color:#067647">
         <b>✓ 자동 사전검사 통과</b>
       </div>`;

  const html = `
<div style="font-family:Arial,'Noto Sans KR',sans-serif;max-width:680px;margin:auto;color:#172033">
  <div style="padding:22px 24px;background:#0f172a;color:white;border-radius:16px 16px 0 0">
    <div style="font-size:13px;opacity:.8">하냥이 인형탈 대여 신청</div>
    <div style="font-size:24px;font-weight:800;margin-top:4px">${escapeHtml_(r.organization)}</div>
  </div>
  <div style="border:1px solid #dbe2ea;border-top:0;padding:22px 24px;border-radius:0 0 16px 16px">
    <p><b>신청번호</b> ${escapeHtml_(r.applicationId)}</p>
    <p><b>담당자</b> ${escapeHtml_(r.managerName)} / ${escapeHtml_(r.phone)}</p>
    <p><b>이메일</b> ${escapeHtml_(r.email)}</p>
    <p><b>대여기간</b> ${formatDateKey_(r.startDate)} ~ ${formatDateKey_(r.endDate)}</p>
    <p><b>사용장소</b> ${escapeHtml_(r.location)}</p>
    <p><b>착용자 신장</b> ${escapeHtml_(String(r.height))}cm</p>
    <p><b>교외 사용</b> ${r.offCampus ? '예' : '아니오'}</p>
    <p><b>사용목적</b><br>${nl2br_(r.purpose)}</p>
    ${notes}
    <div style="margin-top:22px">
      <a href="${escapeAttr_(approveUrl)}" style="display:inline-block;margin:0 8px 8px 0;padding:13px 20px;background:#12b76a;color:white;text-decoration:none;border-radius:10px;font-weight:800">✅ 승인 검토</a>
      <a href="${escapeAttr_(rejectUrl)}" style="display:inline-block;margin:0 0 8px 0;padding:13px 20px;background:#f04438;color:white;text-decoration:none;border-radius:10px;font-weight:800">❌ 거절 검토</a>
    </div>
    <p style="font-size:12px;line-height:1.55;color:#667085;margin-top:18px">
      메일의 버튼을 눌러도 즉시 승인·거절되지 않습니다. 검토 페이지에서 최종 버튼을 눌러야 상태가 변경됩니다.
    </p>
  </div>
</div>`;

  sendEmail_(
    CONFIG.ADMIN_EMAIL,
    `[하냥이 인형탈 대여신청] ${formatDateKey_(r.startDate)} / ${r.organization}`,
    `하냥이 인형탈 대여 신청이 접수되었습니다.\n신청번호: ${r.applicationId}\n신청단체: ${r.organization}\n대여기간: ${formatDateKey_(r.startDate)} ~ ${formatDateKey_(r.endDate)}\n\n승인 검토: ${approveUrl}\n거절 검토: ${rejectUrl}`,
    html
  );
}

function sendApplicantReceipt_(r) {
  const html = `
<div style="font-family:Arial,'Noto Sans KR',sans-serif;max-width:640px;margin:auto;color:#172033">
  <h2>하냥이 인형탈 대여 신청이 접수되었습니다.</h2>
  <p>담당자 검토 후 승인 또는 거절 결과를 이 이메일로 안내합니다.</p>
  <div style="padding:16px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px">
    <p><b>신청번호</b> ${escapeHtml_(r.applicationId)}</p>
    <p><b>신청단체</b> ${escapeHtml_(r.organization)}</p>
    <p><b>대여기간</b> ${formatDateKey_(r.startDate)} ~ ${formatDateKey_(r.endDate)}</p>
  </div>
  <p style="font-size:13px;color:#667085">문의: ${escapeHtml_(CONFIG.ADMIN_EMAIL)}</p>
</div>`;

  sendEmail_(
    r.email,
    `[하냥이 인형탈 대여] 신청 접수 - ${r.applicationId}`,
    `하냥이 인형탈 대여 신청이 접수되었습니다.\n신청번호: ${r.applicationId}\n대여기간: ${formatDateKey_(r.startDate)} ~ ${formatDateKey_(r.endDate)}\n담당자 검토 후 결과를 안내드립니다.`,
    html
  );
}

function sendApplicantDecision_(r) {
  const approved = r.status === CONFIG.STATUS_APPROVED;

  const title = approved
    ? '하냥이 인형탈 대여 신청이 승인되었습니다.'
    : '하냥이 인형탈 대여 신청이 승인되지 않았습니다.';

  const detail = approved
    ? `
      <p>대여 및 반납 일정과 사용 유의사항을 반드시 준수해 주세요.</p>
      <ul style="line-height:1.8">
        <li>사용 종료 후 오염·파손·변형 여부를 확인해 주세요.</li>
        <li>반납 전 세탁 또는 내부 위생 정비를 완료해 주세요.</li>
        <li>이동 시 반드시 동행자와 함께해 주세요.</li>
        <li>우천 등 장비 손상 또는 안전사고 우려 환경에서는 사용하지 마세요.</li>
      </ul>`
    : `<div style="padding:14px;background:#fef3f2;border:1px solid #fecdca;border-radius:10px;color:#b42318">
         <b>거절 사유</b><br>${nl2br_(r.reason)}
       </div>`;

  const html = `
<div style="font-family:Arial,'Noto Sans KR',sans-serif;max-width:640px;margin:auto;color:#172033">
  <h2>${escapeHtml_(title)}</h2>
  <div style="padding:16px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px">
    <p><b>신청번호</b> ${escapeHtml_(r.applicationId)}</p>
    <p><b>신청단체</b> ${escapeHtml_(r.organization)}</p>
    <p><b>대여기간</b> ${r.startDateText} ~ ${r.endDateText}</p>
  </div>
  ${detail}
  <p style="font-size:13px;color:#667085">문의: ${escapeHtml_(CONFIG.ADMIN_EMAIL)}</p>
</div>`;

  sendEmail_(
    r.email,
    approved
      ? `[하냥이 인형탈 대여] 승인 - ${r.applicationId}`
      : `[하냥이 인형탈 대여] 거절 - ${r.applicationId}`,
    approved
      ? `하냥이 인형탈 대여 신청이 승인되었습니다.\n신청번호: ${r.applicationId}\n대여기간: ${r.startDateText} ~ ${r.endDateText}`
      : `하냥이 인형탈 대여 신청이 승인되지 않았습니다.\n신청번호: ${r.applicationId}\n거절 사유: ${r.reason}`,
    html
  );
}

function sendEmail_(recipient, subject, body, htmlBody) {
  const options = {
    htmlBody,
    name: CONFIG.SENDER_NAME,
    replyTo: CONFIG.ADMIN_EMAIL
  };

  // hyerica@hanyang.ac.kr가 현재 실행 계정의 Gmail 발신 별칭이면 해당 주소로 발송.
  // 별칭이 아니면 실행 계정의 기본 주소로 발송하되 Reply-To는 부서 메일로 고정.
  try {
    const aliases = GmailApp.getAliases();
    if (aliases.includes(CONFIG.ADMIN_EMAIL)) {
      options.from = CONFIG.ADMIN_EMAIL;
    }
  } catch (err) {
    console.warn('Gmail alias 확인 실패: ' + err);
  }

  GmailApp.sendEmail(recipient, subject, body, options);
}

/* =========================================================
 * 결과 페이지
 * ========================================================= */

function renderDecisionResultPage_(result) {
  const approved = result.status === CONFIG.STATUS_APPROVED;
  const message = approved
    ? `신청번호 <strong>${escapeHtml_(result.applicationId)}</strong>을(를) 승인했습니다.`
    : `신청번호 <strong>${escapeHtml_(result.applicationId)}</strong>을(를) 거절했습니다.`;

  const warning = result.mailWarning
    ? `<div style="margin-top:14px;padding:12px;border-radius:10px;background:#fffaeb;color:#93370d">${escapeHtml_(result.mailWarning)}</div>`
    : '';

  return renderSimplePage_(
    approved ? '승인 완료' : '거절 완료',
    `${message}<br><br>신청자: ${escapeHtml_(result.email)}${warning}`,
    approved ? 'success' : 'info'
  );
}

function renderSimplePage_(title, bodyHtml, type) {
  const accent = type === 'success'
    ? '#12b76a'
    : type === 'error'
      ? '#f04438'
      : '#2563eb';

  const html = `
<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>${escapeHtml_(title)}</title>
  <style>
    *{box-sizing:border-box}
    body{margin:0;background:#f4f7fb;font-family:Arial,"Noto Sans KR",sans-serif;color:#172033}
    .wrap{max-width:620px;margin:12vh auto;padding:18px}
    .card{background:#fff;border:1px solid #dbe2ea;border-top:6px solid ${accent};border-radius:18px;padding:28px;box-shadow:0 12px 34px rgba(15,23,42,.08)}
    h1{margin:0 0 16px;font-size:28px}
    .body{line-height:1.75}
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card">
      <h1>${escapeHtml_(title)}</h1>
      <div class="body">${bodyHtml}</div>
    </div>
  </div>
</body>
</html>`;

  return HtmlService.createHtmlOutput(html).setTitle(title);
}

/* =========================================================
 * 공통 유틸
 * ========================================================= */

function ensureSetup_() {
  if (!PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')) {
    throw new Error('시스템 초기설정이 필요합니다. Apps Script 편집기에서 setupSystem()을 먼저 실행해 주세요.');
  }
}

function getDbSheet_() {
  ensureSetup_();
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`시트 "${CONFIG.SHEET_NAME}"을 찾을 수 없습니다.`);
  return sheet;
}

function cleanText_(value, maxLen) {
  let s = String(value == null ? '' : value)
    .replace(/\u0000/g, '')
    .trim();

  if (s.length > maxLen) s = s.slice(0, maxLen);
  return s;
}

function safeSheetValue_(value) {
  if (typeof value !== 'string') return value;
  // 스프레드시트 수식 삽입 방지
  return /^[\s]*[=+\-@]/.test(value) ? "'" + value : value;
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function toBool_(value) {
  return value === true || value === 'true' || value === '1' || value === 1 || value === 'yes' || value === '예';
}

function parseDateInput_(value) {
  const s = String(value || '').trim();
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (!m) return null;

  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);

  const date = new Date(y, mo - 1, d);
  if (date.getFullYear() !== y || date.getMonth() !== mo - 1 || date.getDate() !== d) return null;
  date.setHours(0, 0, 0, 0);
  return date;
}

function normalizeDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    const d = new Date(value);
    d.setHours(0, 0, 0, 0);
    return d;
  }

  if (typeof value === 'string') {
    return parseDateInput_(value);
  }

  return null;
}

function formatDateKey_(date) {
  if (!date) return '';
  return Utilities.formatDate(date, CONFIG.TZ, 'yyyy-MM-dd');
}

function createApplicationId_() {
  const stamp = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMMdd-HHmmss');
  const rand = Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase();
  return `HYC-${stamp}-${rand}`;
}

function createToken_() {
  return (
    Utilities.getUuid().replace(/-/g, '') +
    Utilities.getUuid().replace(/-/g, '')
  );
}

function hashToken_(token) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(token),
    Utilities.Charset.UTF_8
  );

  return bytes.map(b => {
    const v = b < 0 ? b + 256 : b;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function escapeAttr_(value) {
  return escapeHtml_(value);
}

function nl2br_(value) {
  return escapeHtml_(value).replace(/\r?\n/g, '<br>');
}
