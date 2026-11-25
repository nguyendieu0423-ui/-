import type { Candidate, ResignationRecord } from '../types/crm';
import {
  extractNormalizedMendanInfo,
  getEffectiveProposedCompany,
} from './crmHelpers';
import { normalizeAndAutoComputeResignationDates } from './resignationHelpers';

export interface ParsedScheduleJst {
  dateIso: string;
  startIso: string;
  endIso: string;
  displayDateTime: string;
  year: number;
  month: number;
  day: number;
}

export interface CalendarEventDraft {
  candidateId: string;
  entityType?: 'candidate' | 'resignation';
  candidateName: string;
  eventTitle: string;
  description: string;
  dateIso?: string;
  startIso: string;
  endIso: string;
  displayDateTime: string;
  existingEventId?: string;
  appointmentType: 'interview' | 'mendan' | 'resignation_deadline';
  reminderMinutes?: number[];
}

const UNSET_DATE_PLACEHOLDER_RX =
  /^(?:chưa\s*(?:đặt\s*hạn|cập\s*nhật|có|rõ)|未設定|未定|不明|n\/a|-+|none|null)$/i;

export function parseScheduleDateTimeJst(
  rawInput?: string,
  defaultHour = 10
): ParsedScheduleJst | null {
  const text = String(rawInput || '').trim();
  if (!text || UNSET_DATE_PLACEHOLDER_RX.test(text)) return null;

  let year: number | null = null;
  let month: number | null = null;
  let day: number | null = null;

  const fullYmdMatch = text.match(/(\d{4})\s*[年\/.\-]\s*(\d{1,2})\s*[月\/.\-]\s*(\d{1,2})/);
  const fullDmyMatch = text.match(/^(\d{1,2})\s*[\/.\-]\s*(\d{1,2})\s*[\/.\-]\s*(\d{4})(?:\s|$|T)/);

  if (fullYmdMatch) {
    year = Number(fullYmdMatch[1]);
    month = Number(fullYmdMatch[2]);
    day = Number(fullYmdMatch[3]);
  } else if (fullDmyMatch) {
    const p1 = Number(fullDmyMatch[1]);
    const p2 = Number(fullDmyMatch[2]);
    year = Number(fullDmyMatch[3]);
    if (p2 > 12 && p1 <= 12) {
      month = p1;
      day = p2;
    } else {
      day = p1;
      month = p2;
    }
  } else {
    const shortMatch = text.match(/(\d{1,2})\s*[月\/]\s*(\d{1,2})/);
    if (!shortMatch) return null;
    year = new Date().getFullYear();
    month = Number(shortMatch[1]);
    day = Number(shortMatch[2]);
  }

  if (!year || !month || !day || month < 1 || month > 12 || day < 1 || day > 31) {
    return null;
  }

  // Reject impossible calendar days such as 2026-02-30 or 2026-04-31
  const verifyUtc = new Date(Date.UTC(year, month - 1, day));
  if (
    verifyUtc.getUTCFullYear() !== year ||
    verifyUtc.getUTCMonth() !== month - 1 ||
    verifyUtc.getUTCDate() !== day
  ) {
    return null;
  }

  // Extract optional HH:mm or HH時mm分
  let hour = defaultHour;
  let minute = 0;
  const timeMatch = text.match(/(?:^|\s|T)(\d{1,2})\s*[:時]\s*(\d{1,2})?/);
  if (timeMatch) {
    const h = Number(timeMatch[1]);
    const m = timeMatch[2] ? Number(timeMatch[2]) : 0;
    if (h >= 0 && h <= 23 && m >= 0 && m <= 59) {
      hour = h;
      minute = m;
    }
  }

  const pad = (n: number) => String(n).padStart(2, '0');
  const startDate = new Date(Date.UTC(year, month - 1, day, hour - 9, minute, 0));
  const endDate = new Date(startDate.getTime() + 60 * 60 * 1000);

  const formatJstIso = (d: Date) => {
    const jst = new Date(d.getTime() + 9 * 60 * 60 * 1000);
    return `${jst.getUTCFullYear()}-${pad(jst.getUTCMonth() + 1)}-${pad(
      jst.getUTCDate()
    )}T${pad(jst.getUTCHours())}:${pad(jst.getUTCMinutes())}:00+09:00`;
  };

  const dateIso = `${year}-${pad(month)}-${pad(day)}`;
  return {
    dateIso,
    startIso: formatJstIso(startDate),
    endIso: formatJstIso(endDate),
    displayDateTime: `${dateIso} ${pad(hour)}:${pad(minute)} (JST)`,
    year,
    month,
    day,
  };
}

export function extractFacebookInfoForCalendar(person: Candidate): string {
  if (isValidSafeFacebookUrl(person.fbProfileUrl)) {
    return person.fbProfileUrl!.trim();
  }
  const combined = `${person.memo || ''}\n${person.fbName || ''}`;
  const urlMatch = combined.match(/https?:\/\/(?:www\.|m\.)?(?:facebook\.com|fb\.com|m\.me|messenger\.com)\/[^\s)]+/i);
  if (urlMatch && isValidSafeFacebookUrl(urlMatch[0])) return urlMatch[0];
  if (person.fbSenderId) {
    return `${person.fbName || 'Messenger'} (Sender ID: ${person.fbSenderId})`;
  }
  return person.fbName || 'Chưa cập nhật';
}

export function buildCandidateCalendarEventDraft(
  person: Candidate
): CalendarEventDraft | null {
  const proposedComp = getEffectiveProposedCompany(person);
  const interviewParsed = parseScheduleDateTimeJst(person.interviewDate);
  const normalizedMendan = extractNormalizedMendanInfo(person);
  const mendanParsed =
    parseScheduleDateTimeJst(person.mendanDate) ||
    parseScheduleDateTimeJst(normalizedMendan.dateLabel);

  const isCompanyInterviewStatus = /面接/.test(String(person.progressStatus || ''));

  let appointmentType: 'interview' | 'mendan' = 'mendan';
  let parsedSchedule: ParsedScheduleJst | null = null;
  let eventTitle = '';

  if (interviewParsed) {
    appointmentType = 'interview';
    parsedSchedule = interviewParsed;
    eventTitle = `[Phỏng vấn] ${person.name || 'Ứng viên'} - ${
      proposedComp || person.desiredIndustry || 'Doanh nghiệp'
    }`;
  } else if (mendanParsed && isCompanyInterviewStatus && proposedComp) {
    appointmentType = 'interview';
    parsedSchedule = mendanParsed;
    eventTitle = `[Phỏng vấn] ${person.name || 'Ứng viên'} - ${proposedComp}`;
  } else if (mendanParsed) {
    appointmentType = 'mendan';
    parsedSchedule = mendanParsed;
    eventTitle = `[Mendan] ${person.name || 'Ứng viên'} - ${
      person.desiredIndustry || 'Chưa rõ ngành'
    }`;
  } else {
    return null;
  }

  const fbDisplay = extractFacebookInfoForCalendar(person);
  const descriptionLines = [
    `Ứng viên: ${person.name || 'Chưa rõ'}`,
    person.fbName ? `Tên Facebook: ${person.fbName}` : '',
    `Loại lịch hẹn: ${
      appointmentType === 'interview' ? 'Phỏng vấn công ty (面接)' : 'Mendan sơ vấn'
    }`,
    `Trình độ JLPT: ${person.jlpt || 'Chưa rõ'}`,
    `Số điện thoại: ${person.phone || 'Chưa cập nhật'}`,
    `Facebook: ${fbDisplay}`,
    `Ngành / Khu vực: ${person.desiredIndustry || '-'} / ${person.desiredArea || '-'}`,
    proposedComp ? `Doanh nghiệp đề xuất: ${proposedComp}` : '',
    `Trạng thái CRM: ${person.progressStatus || '-'}`,
    '',
    'Ghi chú:',
    person.memo ? person.memo.trim() : 'Không có ghi chú',
  ].filter((line, idx, arr) => {
    if (line !== '') return true;
    return idx > 0 && arr[idx - 1] !== '';
  });

  return {
    candidateId: person.id,
    entityType: 'candidate',
    candidateName: person.name || 'Ứng viên',
    eventTitle,
    description: descriptionLines.join('\n'),
    dateIso: parsedSchedule.dateIso,
    startIso: parsedSchedule.startIso,
    endIso: parsedSchedule.endIso,
    displayDateTime: parsedSchedule.displayDateTime,
    existingEventId: person.googleCalendarEventId,
    appointmentType,
    reminderMinutes: [30, 10],
  };
}

export function buildResignationCalendarEventDraft(
  record: ResignationRecord
): CalendarEventDraft | null {
  if (record.paperworkStatus === 'Hoàn tất') {
    return null;
  }

  const normalized = normalizeAndAutoComputeResignationDates(record);

  // Priority order: 1. Hạn giấy tờ (提出期限 = Ngày nghỉ việc - 3 ngày) -> 2. Ngày nghỉ việc (退職日)
  const parsedDeadline = parseScheduleDateTimeJst(normalized.paperworkDeadline, 9);
  const parsedResignation = parseScheduleDateTimeJst(normalized.resignationDate, 9);
  const parsedTarget = parsedDeadline || parsedResignation;
  if (!parsedTarget) {
    return null;
  }

  const resigDateDisplay =
    parsedResignation?.dateIso || normalized.resignationDate || 'Chưa rõ';
  const dateSourceLabel = parsedDeadline
    ? 'Hạn nộp giấy tờ (提出期限 — trước ngày nghỉ việc 3 ngày)'
    : 'Ngày nghỉ việc (退職日)';

  const eventTitle = `[Hạn nộp giấy tờ nghỉ việc] ${
    normalized.fullName || 'Ứng viên'
  } - Nghỉ việc ngày ${resigDateDisplay}`;

  const descriptionLines = [
    `Họ tên: ${normalized.fullName || 'Chưa rõ'}`,
    `Quốc tịch: ${normalized.nationality || 'Chưa cập nhật'}`,
    `Trình độ tiếng Nhật: ${normalized.jlpt || 'Chưa cập nhật'}`,
    `Công ty cũ: ${normalized.previousCompany || 'Chưa cập nhật'}`,
    `Ngày nghỉ việc (退職日): ${resigDateDisplay}`,
    `Hạn nộp giấy tờ (提出期限 - 退職証明書・源泉徴収票): ${
      parsedDeadline?.dateIso || normalized.paperworkDeadline || 'Chưa đặt hạn'
    }`,
    `Mốc lịch đang theo dõi: ${dateSourceLabel}`,
    `Trạng thái giấy tờ: ${normalized.paperworkStatus || 'Chưa làm'}`,
    normalized.phone ? `SĐT / Liên hệ: ${normalized.phone}` : '',
    normalized.fbName ? `Tên Facebook: ${normalized.fbName}` : '',
    normalized.newJobNeed ? `Nhu cầu chuyển việc mới: ${normalized.newJobNeed}` : '',
    '',
    'Ghi chú bàn giao:',
    normalized.handoverNotes ? normalized.handoverNotes.trim() : 'Không có ghi chú bàn giao',
    '',
    '⏰ Nhắc nhở khuyến nghị: trước 1 ngày (24h) và trước 2 giờ.',
  ].filter((line, idx, arr) => {
    if (line !== '') return true;
    return idx > 0 && arr[idx - 1] !== '';
  });

  return {
    candidateId: normalized.id,
    entityType: 'resignation',
    candidateName: normalized.fullName,
    eventTitle,
    description: descriptionLines.join('\n'),
    dateIso: parsedTarget.dateIso,
    startIso: parsedTarget.startIso,
    endIso: parsedTarget.endIso,
    displayDateTime: parsedTarget.displayDateTime,
    existingEventId: normalized.googleCalendarEventId,
    appointmentType: 'resignation_deadline',
    reminderMinutes: [1440, 120], // Nhắc nhở trước 1 ngày (24h) và trước 2 giờ
  };
}

function formatGoogleTemplateTimestamp(isoOrYmd: string, fallbackHour = '090000'): string {
  const m = String(isoOrYmd || '').match(
    /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}))?)?/
  );
  if (!m) return '';
  const datePart = `${m[1]}${m[2]}${m[3]}`;
  if (m[4] && m[5]) {
    return `${datePart}T${m[4]}${m[5]}${m[6] || '00'}`;
  }
  return `${datePart}T${fallbackHour}`;
}

export function buildGoogleCalendarTemplateUrl(draft: CalendarEventDraft): string {
  const startCompact = formatGoogleTemplateTimestamp(
    draft.startIso || draft.dateIso || '',
    '090000'
  );
  const endCompact = formatGoogleTemplateTimestamp(
    draft.endIso || draft.dateIso || '',
    '100000'
  );

  const url = new URL('https://calendar.google.com/calendar/render');
  url.searchParams.set('action', 'TEMPLATE');
  url.searchParams.set('text', draft.eventTitle);
  if (startCompact && endCompact) {
    url.searchParams.set('dates', `${startCompact}/${endCompact}`);
  }
  url.searchParams.set('ctz', 'Asia/Tokyo');
  url.searchParams.set('details', draft.description);
  return url.toString();
}

export const GAS_CALENDAR_APP_SNIPPET = `// Thêm vào hàm doPost(e) trong Google Apps Script để tự động tạo lịch bằng CalendarApp:
if (data.action === 'createCalendarEvent' || data.action === 'createBulkCalendarEvents') {
  var items = data.action === 'createBulkCalendarEvents' ? (data.items || []) : [data];
  var cal = CalendarApp.getDefaultCalendar();
  items.forEach(function(item) {
    var start = new Date(item.startIso || (item.date + 'T09:00:00+09:00'));
    var end = new Date(item.endIso || (item.date + 'T10:00:00+09:00'));
    var ev = cal.createEvent(item.title || item.eventTitle, start, end, { description: item.description || '' });
    var mins = Array.isArray(item.reminderMinutes) ? item.reminderMinutes : [1440, 120];
    mins.forEach(function(m) { ev.addPopupReminder(Number(m)); });
  });
  return ContentService.createTextOutput(JSON.stringify({ ok: true, count: items.length })).setMimeType(ContentService.MimeType.JSON);
}`;

export function isValidGoogleCalendarUrl(url?: string): boolean {
  if (!url) return false;
  try {
    const parsed = new URL(url);
    return (
      parsed.protocol === 'https:' &&
      (parsed.hostname === 'calendar.google.com' ||
        parsed.hostname === 'www.google.com')
    );
  } catch {
    return false;
  }
}

const ALLOWED_FB_HOSTS = new Set([
  'facebook.com',
  'www.facebook.com',
  'm.facebook.com',
  'fb.com',
  'www.fb.com',
  'm.me',
  'messenger.com',
  'www.messenger.com',
  'business.facebook.com',
]);

export function isValidSafeFacebookUrl(url?: string): boolean {
  if (!url) return false;
  try {
    const parsed = new URL(url.trim());
    return parsed.protocol === 'https:' && ALLOWED_FB_HOSTS.has(parsed.hostname.toLowerCase());
  } catch {
    return false;
  }
}

export function normalizeFacebookPageTargetUrl(rawInput?: string): string | null {
  const trimmed = String(rawInput || '').trim();
  if (!trimmed) return null;
  if (/^https?:\/\//i.test(trimmed)) {
    const normalized = trimmed.replace(/^http:\/\//i, 'https://');
    return isValidSafeFacebookUrl(normalized) ? normalized : null;
  }
  const cleanHandle = trimmed.replace(/^@+/, '').replace(/^\/+|\/+$/g, '');
  if (!/^[A-Za-z0-9._-]{2,100}$/.test(cleanHandle)) {
    return null;
  }
  return `https://m.me/${cleanHandle}`;
}

export function getSafeFacebookMessengerUrl(person: Candidate): string | null {
  if (isValidSafeFacebookUrl(person.fbProfileUrl)) {
    return person.fbProfileUrl!.trim();
  }
  const combined = `${person.memo || ''}\n${person.fbName || ''}`;
  const urlMatch = combined.match(
    /https?:\/\/(?:www\.|m\.|business\.)?(?:facebook\.com|fb\.com|m\.me|messenger\.com)\/[^\s)]+/i
  );
  if (urlMatch && isValidSafeFacebookUrl(urlMatch[0])) {
    return urlMatch[0];
  }
  return null;
}
