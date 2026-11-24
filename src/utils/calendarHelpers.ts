import type { Candidate, ResignationRecord } from '../types/crm';
import {
  extractNormalizedMendanInfo,
  getEffectiveProposedCompany,
} from './crmHelpers';

export interface ParsedScheduleJst {
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
  startIso: string;
  endIso: string;
  displayDateTime: string;
  existingEventId?: string;
  appointmentType: 'interview' | 'mendan' | 'resignation_deadline';
  reminderMinutes?: number[];
}

export function parseScheduleDateTimeJst(rawInput?: string): ParsedScheduleJst | null {
  const text = String(rawInput || '').trim();
  if (!text) return null;

  let year: number | null = null;
  let month: number | null = null;
  let day: number | null = null;

  const fullMatch = text.match(/(\d{4})\s*[年\/\-]\s*(\d{1,2})\s*[月\/\-]\s*(\d{1,2})/);
  if (fullMatch) {
    year = Number(fullMatch[1]);
    month = Number(fullMatch[2]);
    day = Number(fullMatch[3]);
  } else {
    const shortMatch = text.match(/(\d{1,2})\s*[月\/]\s*(\d{1,2})/);
    if (!shortMatch) return null;
    year = new Date().getFullYear();
    month = Number(shortMatch[1]);
    day = Number(shortMatch[2]);
  }

  if (
    !year ||
    !month ||
    !day ||
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > 31
  ) {
    return null;
  }

  // Extract optional HH:mm or HH時mm分
  let hour = 10;
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

  return {
    startIso: formatJstIso(startDate),
    endIso: formatJstIso(endDate),
    displayDateTime: `${year}-${pad(month)}-${pad(day)} ${pad(hour)}:${pad(minute)} (JST)`,
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
  if (!record.paperworkDeadline || record.paperworkStatus === 'Hoàn tất') {
    return null;
  }
  const hasExplicitTime = /[:時]/.test(record.paperworkDeadline);
  const parsed = parseScheduleDateTimeJst(
    hasExplicitTime ? record.paperworkDeadline : `${record.paperworkDeadline} 09:00`
  );
  if (!parsed) return null;

  const eventTitle = `[Hạn giấy tờ nghỉ việc] ${record.fullName || 'Ứng viên'} - ${
    record.previousCompany || 'Công ty cũ'
  }`;

  const descriptionLines = [
    `Họ tên: ${record.fullName || 'Chưa rõ'}`,
    `Quốc tịch: ${record.nationality || 'Chưa rõ'}`,
    `Công ty cũ đang làm việc: ${record.previousCompany || 'Chưa rõ'}`,
    `Ngày nghỉ việc chính thức (退職日): ${record.resignationDate || 'Chưa cập nhật'}`,
    `Hạn chót làm & nộp giấy tờ (書類提出期限): ${record.paperworkDeadline}`,
    `Trạng thái giấy tờ: ${record.paperworkStatus || 'Chưa làm'}`,
    record.phone ? `SĐT: ${record.phone}` : '',
    record.newJobNeed ? `Nhu cầu chuyển việc mới: ${record.newJobNeed}` : '',
    '',
    'Ghi chú bàn giao:',
    record.handoverNotes ? record.handoverNotes.trim() : 'Không có ghi chú bàn giao',
  ].filter((line, idx, arr) => {
    if (line !== '') return true;
    return idx > 0 && arr[idx - 1] !== '';
  });

  return {
    candidateId: record.id,
    entityType: 'resignation',
    candidateName: record.fullName,
    eventTitle,
    description: descriptionLines.join('\n'),
    startIso: parsed.startIso,
    endIso: parsed.endIso,
    displayDateTime: parsed.displayDateTime,
    existingEventId: record.googleCalendarEventId,
    appointmentType: 'resignation_deadline',
    reminderMinutes: [4320, 1440, 120], // 3 ngày, 1 ngày, 2 giờ trước
  };
}

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
