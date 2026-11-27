import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import mammoth from 'mammoth';
import type {
  Candidate,
  ChannelType,
  ExcelAnchor,
  KyujinCompany,
  KyujinMatchEvaluation,
  KyujinMatchResult,
  KyujinSlot,
  MissingInfoItem,
  PtPartner,
  ReplyScenario,
  ResignationRecord,
  RirekishoDraft,
  RirekishoRow,
  SavedRirekishoRecord,
} from '../types/crm';
import { compressImageFileToBase64, fileToBase64Raw } from './imageCompression';
import {
  extractAnnualSalaryText,
  extractEstimatedNetSalaryText,
  extractKyujinSalaryFromRows,
  extractKyujinSalaryInfo,
  getKyujinSlotSalaryDisplayMeta,
  sanitizeExcelFloatText,
  stripKyujinFormBoilerplate,
} from './salaryParser';
import {
  normalizeResignationNationality,
  suggestNationalityFromName,
} from './resignationHelpers';
import { extractPdfWithOcrFallback } from './pdfExtractor';

export const copyTextSafely = async (text: string): Promise<boolean> => {
  const str = String(text || '');
  if (!str) return false;
  try {
    if (navigator?.clipboard?.writeText) {
      await navigator.clipboard.writeText(str);
      return true;
    }
  } catch {
    // Fallback for sandboxed iframes where Clipboard API is blocked
  }
  try {
    const textarea = document.createElement('textarea');
    textarea.value = str;
    textarea.setAttribute('readonly', '');
    textarea.style.position = 'fixed';
    textarea.style.top = '-9999px';
    textarea.style.left = '-9999px';
    document.body.appendChild(textarea);
    textarea.select();
    const ok = document.execCommand('copy');
    document.body.removeChild(textarea);
    return ok;
  } catch {
    return false;
  }
};

export const formatPublicArea = (rawArea = ''): string => {
  const s = String(rawArea || '')
    .replace(/〒\s*\d{3}-?\d{4}/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!s) return '';
  const m = s.match(
    /^((?:東京都|北海道|(?:大阪|京都)府|.{2,3}県)?[^0-9０-９丁目番地号\s]+?(?:市[^0-9０-９丁目番地号\s]+?区|区|市|町|村))/
  );
  if (m && m[1]) return m[1].trim();
  return s.replace(/[0-9０-９]+(?:丁目|番地|番|号|-|−).*$/, '').trim();
};

export const parseFlexibleDateToTimestamp = (raw = '', defaultYear = 2026): number => {
  const s = String(raw || '').trim();
  if (!s) return 0;

  const fullMatch = s.match(/(\d{4})\s*[年\/\-]\s*(\d{1,2})\s*[月\/\-]\s*(\d{1,2})/);
  if (fullMatch) {
    const y = parseInt(fullMatch[1], 10);
    const m = parseInt(fullMatch[2], 10) - 1;
    const d = parseInt(fullMatch[3], 10);
    const dt = new Date(y, m, d);
    return Number.isNaN(dt.getTime()) ? 0 : dt.getTime();
  }

  const shortMatch = s.match(/(?:^|[^\d])(\d{1,2})\s*[月\/]\s*(\d{1,2})(?:日|[^\d]|$)/);
  if (shortMatch) {
    const m = parseInt(shortMatch[1], 10) - 1;
    const d = parseInt(shortMatch[2], 10);
    if (m >= 0 && m <= 11 && d >= 1 && d <= 31) {
      const dt = new Date(defaultYear, m, d);
      return Number.isNaN(dt.getTime()) ? 0 : dt.getTime();
    }
  }

  return 0;
};

export const formatTimestampToIso = (ts: number): string => {
  if (!ts || ts <= 0) return '';
  const dt = new Date(ts);
  if (Number.isNaN(dt.getTime())) return '';
  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, '0');
  const d = String(dt.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
};

export const extractNormalizedMendanInfo = (
  person: Partial<Candidate> = {}
): { hasMendan: boolean; timestamp: number; dateLabel: string } => {
  const candidatesList: Array<{ ts: number; raw: string }> = [];

  if (person.mendanDate) {
    const ts = parseFlexibleDateToTimestamp(person.mendanDate, 2026);
    if (ts > 0) candidatesList.push({ ts, raw: person.mendanDate });
  }
  if (person.interviewDate) {
    const ts = parseFlexibleDateToTimestamp(person.interviewDate, 2026);
    if (ts > 0) candidatesList.push({ ts, raw: person.interviewDate });
  }

  const memo = String(person.memo || '');
  if (memo) {
    const patterns = [
      /(\d{4}\s*[年\/\-]\s*\d{1,2}\s*[月\/\-]\s*\d{1,2}\s*日?)\s*(?:に|の)?\s*(?:面談|面接)/g,
      /(\d{1,2}\s*[月\/]\s*\d{1,2}\s*日?)\s*(?:に|の)?\s*(?:面談|面接)/g,
      /(?:面談|面接)(?:日|予定)?[：:\s]*(\d{4}\s*[年\/\-]\s*\d{1,2}\s*[月\/\-]\s*\d{1,2}\s*日?|\d{1,2}\s*[月\/]\s*\d{1,2}\s*日?)/g,
    ];
    patterns.forEach((rx) => {
      let m: RegExpExecArray | null;
      while ((m = rx.exec(memo)) !== null) {
        const ts = parseFlexibleDateToTimestamp(m[1], 2026);
        if (ts > 0) candidatesList.push({ ts, raw: m[1] });
      }
    });
  }

  candidatesList.sort((a, b) => b.ts - a.ts);
  const latest = candidatesList[0] || null;
  const mendanStatuses = [
    '面談済み',
    '面談調整中',
    '1次面接予定',
    '2次面接予定',
    '面接調整中',
  ];
  const hasMendan =
    Boolean(latest) ||
    mendanStatuses.includes(String(person.progressStatus || '').trim()) ||
    /面談した|面談済|面接予定/.test(memo);

  return {
    hasMendan,
    timestamp: latest ? latest.ts : 0,
    dateLabel: latest ? formatTimestampToIso(latest.ts) : '',
  };
};

export const calcJapaneseAgeFromBirthDate = (birthDate = ''): string => {
  const s = String(birthDate || '').trim();
  const m = s.match(/(\d{4})\s*[年\/\-]\s*(\d{1,2})\s*[月\/\-]\s*(\d{1,2})/);
  if (!m) return '';
  const by = parseInt(m[1], 10);
  const bm = parseInt(m[2], 10);
  const bd = parseInt(m[3], 10);
  const now = new Date();
  let age = now.getFullYear() - by;
  const mDiff = now.getMonth() + 1 - bm;
  if (mDiff < 0 || (mDiff === 0 && now.getDate() < bd)) {
    age -= 1;
  }
  return age >= 15 && age <= 80 ? String(age) : '';
};

export const createInitialRirekishoDraft = (cand: Partial<Candidate> = {}): RirekishoDraft => {
  const now = new Date();
  const todayJp = `${now.getFullYear()}年 ${now.getMonth() + 1}月 ${now.getDate()}日現在`;
  const computedAge = cand.age || calcJapaneseAgeFromBirthDate(cand.birthDate) || '';
  return {
    dateStr: todayJp,
    furigana: cand.furigana || '',
    fullName: cand.name || '',
    birthDate: cand.birthDate || '',
    age: computedAge,
    gender: cand.gender || '男性',
    postalCode: cand.postalCode || '',
    address: cand.address || '',
    phone: cand.phone || '',
    email: cand.email || '',
    nationality:
      cand.nationality && cand.nationality !== '国籍不明' ? cand.nationality : '',
    visaExpiry: cand.visaExpiry || '',
    visaStatus: cand.visaStatus || '特定技能1号',
    educationRows: [
      { year: '', month: '', content: '' },
      { year: '', month: '', content: '' },
    ],
    workRows: cand.japanExperience
      ? [{ year: '', month: '', content: cand.japanExperience }]
      : [
          { year: '', month: '', content: '' },
          { year: '', month: '', content: '' },
        ],
    licenseRows: cand.passedExam
      ? [
          { year: '', month: '', content: cand.passedExam },
          cand.jlpt && cand.jlpt !== '要確認'
            ? { year: '', month: '', content: `日本語能力試験 ${cand.jlpt} 合格` }
            : { year: '', month: '', content: '' },
        ]
      : [
          { year: '', month: '', content: '' },
          { year: '', month: '', content: '' },
        ],
    nativeLanguage:
      cand.nationality === 'ベトナム'
        ? 'ベトナム語'
        : cand.nationality === 'ミャンマー'
        ? 'ミャンマー語'
        : cand.nationality === 'インドネシア'
        ? 'インドネシア語'
        : '母国語',
    availableLanguages: `日本語（${cand.jlpt || '日常会話'}）`,
    caSpeaking:
      cand.jlpt === 'N1' || cand.jlpt === 'N2' ? '日常会話以上（スムーズ）' : '日常会話可能',
    caListening:
      cand.jlpt === 'N1' || cand.jlpt === 'N2' ? '指示理解問題なし' : '基本的な指示理解可能',
    motivation: `${cand.desiredIndustry || '特定技能'}の分野でこれまでの経験と日本語力を活かし、長期的に貴社へ貢献したいと考え志望いたしました。`,
    selfPr: `${cand.japanExperience ? cand.japanExperience + 'の経験があり、' : ''}真面目で責任感が強く、職場のルールを守り周囲と協力して業務に取り組むことができます。`,
    personalRequests: `希望職種：${cand.desiredIndustry || '貴社規定に従います'}\n希望勤務地：${cand.desiredArea || '貴社規定に従います'}\n入社可能時期：${cand.startDate || '相談の上決定'}`,
    allergy: '特になし',
    dependentsCount: '0人',
    spouse: '無',
    spouseSupport: '無',
    photoDataUrl: '',
  };
};

const sanitizeRirekishoRows = (rows: unknown, fallback: RirekishoRow[]): RirekishoRow[] => {
  if (!Array.isArray(rows) || rows.length === 0) return fallback;
  return rows.map((r) => {
    const obj = r && typeof r === 'object' ? (r as Record<string, unknown>) : {};
    return {
      year: cleanLegacyStr(obj.year),
      month: cleanLegacyStr(obj.month),
      content: cleanLegacyStr(obj.content),
    };
  });
};

export const toSafeStringArray = (val: unknown): string[] => {
  if (Array.isArray(val)) {
    return val
      .map((item) => cleanLegacyStr(item))
      .filter(Boolean);
  }
  if (typeof val === 'string') {
    const trimmed = val.trim();
    if (!trimmed) return [];
    if (trimmed.startsWith('[')) {
      try {
        const parsed = JSON.parse(trimmed);
        if (Array.isArray(parsed)) {
          return parsed.map((item) => cleanLegacyStr(item)).filter(Boolean);
        }
      } catch {
        // Fallback to delimiter split
      }
    }
    return trimmed
      .split(/[,、，;\n]+/)
      .map((item) => item.trim())
      .filter(Boolean);
  }
  return [];
};

export const sanitizeSavedRirekishoRecord = (
  rawInput: unknown,
  fallbackCandidate?: Partial<Candidate>
): SavedRirekishoRecord | undefined => {
  if (!rawInput) return undefined;
  let resolvedInput: unknown = rawInput;
  if (typeof resolvedInput === 'string') {
    const trimmed = resolvedInput.trim();
    if (!trimmed) return undefined;
    try {
      resolvedInput = JSON.parse(trimmed);
    } catch {
      return undefined;
    }
  }
  if (!resolvedInput || typeof resolvedInput !== 'object' || Array.isArray(resolvedInput)) {
    return undefined;
  }
  const raw = resolvedInput as Record<string, unknown>;
  const nestedDraft =
    raw.draft && typeof raw.draft === 'object' && !Array.isArray(raw.draft)
      ? (raw.draft as Record<string, unknown>)
      : null;
  // Support legacy saved_cv where draft fields sit directly on the object
  const looksLikeFlatDraft =
    !nestedDraft &&
    (raw.fullName !== undefined ||
      raw.full_name !== undefined ||
      raw.name !== undefined ||
      raw.educationRows !== undefined ||
      raw.workRows !== undefined ||
      raw.motivation !== undefined ||
      raw.selfPr !== undefined ||
      raw.photoDataUrl !== undefined ||
      raw.photo !== undefined);
  const rawDraft = nestedDraft || (looksLikeFlatDraft ? raw : null);
  if (!rawDraft) return undefined;

  const base = createInitialRirekishoDraft(fallbackCandidate || {});
  const rawPhotoCandidate =
    typeof rawDraft.photoDataUrl === 'string'
      ? rawDraft.photoDataUrl.trim()
      : typeof rawDraft.photo === 'string'
      ? rawDraft.photo.trim()
      : '';
  const photoDataUrl = rawPhotoCandidate.startsWith('data:image/') ? rawPhotoCandidate : '';

  const draft: RirekishoDraft = {
    dateStr: cleanLegacyStr(rawDraft.dateStr ?? rawDraft.date_str) || base.dateStr,
    furigana: cleanLegacyStr(rawDraft.furigana),
    fullName: cleanLegacyStr(rawDraft.fullName ?? rawDraft.full_name ?? rawDraft.name) || base.fullName,
    birthDate: cleanLegacyStr(rawDraft.birthDate ?? rawDraft.birth_date),
    age: cleanLegacyStr(rawDraft.age),
    gender: cleanLegacyStr(rawDraft.gender),
    postalCode: cleanLegacyStr(rawDraft.postalCode ?? rawDraft.postal_code),
    address: cleanLegacyStr(rawDraft.address),
    phone: cleanLegacyStr(rawDraft.phone),
    email: cleanLegacyStr(rawDraft.email),
    nationality: cleanLegacyStr(rawDraft.nationality),
    visaExpiry: cleanLegacyStr(rawDraft.visaExpiry ?? rawDraft.visa_expiry),
    visaStatus: cleanLegacyStr(rawDraft.visaStatus ?? rawDraft.visa_status),
    educationRows: sanitizeRirekishoRows(rawDraft.educationRows ?? rawDraft.education_rows, base.educationRows),
    workRows: sanitizeRirekishoRows(rawDraft.workRows ?? rawDraft.work_rows, base.workRows),
    licenseRows: sanitizeRirekishoRows(rawDraft.licenseRows ?? rawDraft.license_rows, base.licenseRows),
    nativeLanguage: cleanLegacyStr(rawDraft.nativeLanguage ?? rawDraft.native_language),
    availableLanguages: cleanLegacyStr(rawDraft.availableLanguages ?? rawDraft.available_languages),
    caSpeaking: cleanLegacyStr(rawDraft.caSpeaking ?? rawDraft.ca_speaking),
    caListening: cleanLegacyStr(rawDraft.caListening ?? rawDraft.ca_listening),
    motivation: cleanLegacyStr(rawDraft.motivation),
    selfPr: cleanLegacyStr(rawDraft.selfPr ?? rawDraft.self_pr),
    personalRequests: cleanLegacyStr(rawDraft.personalRequests ?? rawDraft.personal_requests),
    allergy: cleanLegacyStr(rawDraft.allergy),
    dependentsCount: cleanLegacyStr(rawDraft.dependentsCount ?? rawDraft.dependents_count),
    spouse: cleanLegacyStr(rawDraft.spouse),
    spouseSupport: cleanLegacyStr(rawDraft.spouseSupport ?? rawDraft.spouse_support),
    photoDataUrl,
  };

  return {
    draft,
    sourceFileName: cleanLegacyStr(raw.sourceFileName) || undefined,
    exportedFileName: cleanLegacyStr(raw.exportedFileName) || undefined,
    exportedAt: cleanLegacyStr(raw.exportedAt) || undefined,
    updatedAt: cleanLegacyStr(raw.updatedAt) || new Date().toISOString(),
  };
};

export const getInitialRirekishoDraftForCandidate = (
  cand: Partial<Candidate> = {}
): RirekishoDraft => {
  if (cand.savedRirekisho?.draft) {
    const sanitized = sanitizeSavedRirekishoRecord(cand.savedRirekisho, cand);
    if (sanitized) return sanitized.draft;
  }
  return createInitialRirekishoDraft(cand);
};

export const getMissingInfoForCandidate = (person: Partial<Candidate> = {}): MissingInfoItem[] => {
  const missing: MissingInfoItem[] = [];
  if (!String(person.age || '').trim() && !String(person.birthDate || '').trim()) {
    missing.push({
      key: 'age',
      ja: '年齢・生年月日',
      vi: 'Tuổi / Ngày sinh',
      en: 'Age / Date of birth',
    });
  }
  if (!String(person.visaExpiry || '').trim()) {
    missing.push({
      key: 'visaExpiry',
      ja: '現在の在留期限',
      vi: 'Hạn visa hiện tại',
      en: 'Current visa expiry date',
    });
  }
  if (!String(person.jlpt || '').trim() || person.jlpt === '要確認') {
    missing.push({
      key: 'jlpt',
      ja: '日本語レベル（JLPT/JFT）',
      vi: 'Trình độ tiếng Nhật (JLPT/JFT)',
      en: 'Japanese level (JLPT/JFT)',
    });
  }
  if (!String(person.passedExam || '').trim()) {
    missing.push({
      key: 'passedExam',
      ja: '合格済みの特定技能試験・資格',
      vi: 'Bằng thi tay nghề Tokutei đã đỗ',
      en: 'Passed Tokutei skill exam / certificates',
    });
  }
  if (!String(person.japanExperience || '').trim()) {
    missing.push({
      key: 'japanExperience',
      ja: '日本での職歴・経験',
      vi: 'Kinh nghiệm làm việc tại Nhật',
      en: 'Work experience in Japan',
    });
  }
  if (!String(person.desiredArea || '').trim() || !String(person.startDate || '').trim()) {
    missing.push({
      key: 'areaStart',
      ja: '希望勤務エリア・転職可能時期',
      vi: 'Khu vực & Thời gian có thể đi làm',
      en: 'Preferred area & available start date',
    });
  }
  if (!String(person.phone || '').trim() && !String(person.email || '').trim()) {
    missing.push({
      key: 'contact',
      ja: '電話番号・LINE・メールアドレス',
      vi: 'SĐT / LINE / Email',
      en: 'Phone number / LINE / Email',
    });
  }
  return missing;
};

export const getEffectiveProposedCompany = (person: Partial<Candidate> = {}): string => {
  const direct = String(person.proposedCompany || '').trim();
  if (direct) return direct;
  const memo = String(person.memo || '');
  const m = memo.match(
    /([^\s、。,.\n（）()]+(?:株式会社|有限会社|医療法人|社会福祉法人|協同組合|施設|グループ|店)[^\s、。,.\n（）()]*)\s*(?:へ提案|に提案|へ推薦|に推薦|提案中|面接)/
  );
  if (m && m[1]) return m[1].trim();
  return '';
};

export const cleanLegacyStr = (val: unknown): string => {
  if (val === null || val === undefined) return '';
  const s = String(val).trim();
  if (!s || s.toLowerCase() === 'undefined' || s.toLowerCase() === 'null') return '';
  return s;
};

export const classifyProposedSubTab = (
  person: Partial<Candidate> = {}
): 'in_progress' | 'naitei' | 'jitai' => {
  const st = cleanLegacyStr(person.progressStatus);
  if (['内定', '内定承諾', '入社手続き中', '入社済み'].includes(st)) {
    return 'naitei';
  }
  if (
    ['内定辞退', '選考中辞退', '不合格', '辞退・終了'].includes(st) ||
    st.includes('辞退') ||
    st.includes('不合格') ||
    Boolean(extractDeclineReason(person))
  ) {
    return 'jitai';
  }
  return 'in_progress';
};

export const normalizeCandidateNationality = (raw: unknown): string => {
  const s = String(raw ?? '')
    .replace(/\s*[\(（][^()（）]*(?:xác\s*nhận|確認|từ\s*hội\s*thoại|ước\s*đoán|suy\s*luận|khai\s*báo)[^()（）]*[\)）]\s*$/i, '')
    .trim();
  if (!s) return '';
  return normalizeResignationNationality(s);
};

export const autoFixCandidatesNationalities = (
  list: Candidate[] = []
): { updated: Candidate[]; changedCount: number } => {
  let changedCount = 0;
  const updated = list.map((cand) => {
    const rawRec = cand as Candidate & {
      gasProtected?: boolean;
      manualProtectedFields?: string[];
    };
    const isProtected =
      Boolean(cand.nationalityManuallyEdited) ||
      Boolean(rawRec.gasProtected) ||
      Boolean(rawRec.manualProtectedFields?.includes('nationality'));
    if (isProtected) return cand;

    const currentNat = normalizeCandidateNationality(cand.nationality);
    if (currentNat && currentNat !== '国籍不明' && currentNat !== 'ベトナム') {
      return cand;
    }

    const inferred = suggestNationalityFromName(cand.name || cand.fbName || '');
    if (!inferred || inferred === currentNat) return cand;

    changedCount += 1;
    return {
      ...cand,
      nationality: inferred,
    };
  });
  return { updated, changedCount };
};

export const findDuplicateCandidate = (
  candidatesList: Candidate[] = [],
  draft: {
    id?: string;
    name?: string;
    fbName?: string;
    fbProfileUrl?: string;
    phone?: string;
  } = {}
): Candidate | null => {
  const normText = (v?: string) => String(v || '').trim().toLowerCase();
  const normPhone = (v?: string) => String(v || '').replace(/\D/g, '');
  const normUrl = (v?: string) =>
    String(v || '')
      .trim()
      .toLowerCase()
      .replace(/^https?:\/\/(?:www\.|m\.|mbasic\.)?/, '')
      .replace(/[?#].*$/, '')
      .replace(/\/+$/, '');

  const dName = normText(draft.name);
  const dFbName = normText(draft.fbName);
  const dUrl = normUrl(draft.fbProfileUrl);
  const dPhone = normPhone(draft.phone);

  for (const cand of candidatesList) {
    if (draft.id && cand.id === draft.id) continue;
    const cUrl = normUrl(cand.fbProfileUrl);
    if (dUrl && cUrl && dUrl === cUrl) return cand;

    const cPhone = normPhone(cand.phone);
    if (dPhone.length >= 9 && cPhone.length >= 9 && dPhone === cPhone) return cand;

    const cName = normText(cand.name);
    if (dName.length >= 3 && cName && dName === cName) return cand;

    const cFbName = normText(cand.fbName);
    if (dFbName.length >= 3 && cFbName && dFbName === cFbName) return cand;
  }
  return null;
};

const normUrlKey = (v?: string): string =>
  String(v || '')
    .trim()
    .toLowerCase()
    .replace(/^https?:\/\/(?:www\.|m\.|mbasic\.)?/, '')
    .replace(/[?#].*$/, '')
    .replace(/\/+$/, '');

const normPhoneKey = (v?: string): string => String(v || '').replace(/\D/g, '');

const normNameKey = (v?: string): string =>
  String(v || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();

const isRealSocialUrl = (v?: string): boolean => {
  const s = String(v || '').trim();
  if (!s) return false;
  if (/^(?:-|—|なし|無|不明|未定|n\/a|none)$/i.test(s)) return false;
  return /^https?:\/\//i.test(s) || /(?:facebook\.com|fb\.com|m\.me|messenger\.com)/i.test(s);
};

const isFbNameSameOrSubsetOfName = (fbName?: string, legalName?: string): boolean => {
  const f = normNameKey(fbName);
  const n = normNameKey(legalName);
  if (!f) return true;
  if (!n) return false;
  if (f === n) return true;
  const fTokens = f.split(/[^a-z0-9\p{L}]+/u).filter((t) => t.length >= 2);
  const nTokens = new Set(n.split(/[^a-z0-9\p{L}]+/u).filter((t) => t.length >= 2));
  if (fTokens.length === 0 || nTokens.size === 0) return false;
  return fTokens.every((t) => nTokens.has(t));
};

const areCandidatesSamePerson = (a: Candidate, b: Candidate): boolean => {
  if (a.id && b.id && a.id === b.id) return true;

  const urlA = normUrlKey(a.fbProfileUrl);
  const urlB = normUrlKey(b.fbProfileUrl);
  if (urlA && urlB) {
    return urlA === urlB;
  }

  const phoneA = normPhoneKey(a.phone);
  const phoneB = normPhoneKey(b.phone);
  if (phoneA.length >= 9 && phoneB.length >= 9) {
    return phoneA === phoneB;
  }

  const nameA = normNameKey(a.name);
  const nameB = normNameKey(b.name);
  if (nameA.length >= 3 && nameB && nameA === nameB) {
    // Do not merge if birth dates explicitly conflict
    const bdA = String(a.birthDate || '').trim();
    const bdB = String(b.birthDate || '').trim();
    if (bdA && bdB && bdA !== bdB) return false;
    return true;
  }

  return false;
};

const mergeCandidateRecords = (primary: Candidate, secondary: Candidate): Candidate => {
  const pickStr = (a?: string, b?: string) => {
    const sa = String(a || '').trim();
    const sb = String(b || '').trim();
    if (!sa || sa === '要確認' || sa === '国籍不明' || sa === '未定') {
      return sb || sa;
    }
    return sa;
  };
  const mergeUniqueText = (a?: string, b?: string) => {
    const sa = String(a || '').trim();
    const sb = String(b || '').trim();
    if (!sa) return sb;
    if (!sb) return sa;
    if (sa === sb || sa.includes(sb)) return sa;
    if (sb.includes(sa)) return sb;
    return `${sa}\n${sb}`;
  };

  const isSch = Boolean(
    primary._isSchool ||
      primary.channel === 'school' ||
      secondary._isSchool ||
      secondary.channel === 'school' ||
      String(primary.schoolName || secondary.schoolName || '').trim()
  );

  const primaryStatus = String(primary.progressStatus || '').trim();
  const secondaryStatus = String(secondary.progressStatus || '').trim();
  const effectiveStatus =
    (!primaryStatus || primaryStatus === '新規相談') &&
    secondaryStatus &&
    secondaryStatus !== '新規相談'
      ? secondaryStatus
      : primaryStatus || secondaryStatus || '新規相談';

  return {
    ...secondary,
    ...primary,
    id: primary.id || secondary.id,
    fbSenderId: pickStr(primary.fbSenderId, secondary.fbSenderId),
    fbProfileUrl: pickStr(primary.fbProfileUrl, secondary.fbProfileUrl),
    name: pickStr(primary.name, secondary.name),
    fbName: pickStr(primary.fbName, secondary.fbName),
    furigana: pickStr(primary.furigana, secondary.furigana),
    channel: isSch
      ? 'school'
      : primary.channel === 'personal' && isRealSocialUrl(primary.fbProfileUrl)
      ? 'personal'
      : secondary.channel === 'personal' && isRealSocialUrl(secondary.fbProfileUrl)
      ? 'personal'
      : primary.channel || secondary.channel || 'other',
    schoolName: pickStr(primary.schoolName, secondary.schoolName),
    graduationDate: pickStr(primary.graduationDate, secondary.graduationDate),
    nationality: pickStr(primary.nationality, secondary.nationality),
    gender: pickStr(primary.gender, secondary.gender),
    age: pickStr(primary.age, secondary.age),
    birthDate: pickStr(primary.birthDate, secondary.birthDate),
    postalCode: pickStr(primary.postalCode, secondary.postalCode),
    jlpt: pickStr(primary.jlpt, secondary.jlpt) || '要確認',
    desiredIndustry: pickStr(primary.desiredIndustry, secondary.desiredIndustry),
    desiredArea: pickStr(primary.desiredArea, secondary.desiredArea),
    conversionGoal: pickStr(primary.conversionGoal, secondary.conversionGoal),
    visaStatus: pickStr(primary.visaStatus, secondary.visaStatus),
    visaExpiry: pickStr(primary.visaExpiry, secondary.visaExpiry),
    passedExam: pickStr(primary.passedExam, secondary.passedExam),
    japanExperience: pickStr(primary.japanExperience, secondary.japanExperience),
    salaryExpectation: pickStr(primary.salaryExpectation, secondary.salaryExpectation),
    negotiationLog: mergeUniqueText(primary.negotiationLog, secondary.negotiationLog),
    startDate: pickStr(primary.startDate, secondary.startDate),
    phone: pickStr(primary.phone, secondary.phone),
    email: pickStr(primary.email, secondary.email),
    address: pickStr(primary.address, secondary.address),
    progressStatus: effectiveStatus,
    proposedCompany: pickStr(primary.proposedCompany, secondary.proposedCompany),
    proposalDate: pickStr(primary.proposalDate, secondary.proposalDate),
    naiteiCompany: pickStr(primary.naiteiCompany, secondary.naiteiCompany),
    expectedJoinDate: pickStr(primary.expectedJoinDate, secondary.expectedJoinDate),
    visaSubmissionDeadline: pickStr(primary.visaSubmissionDeadline, secondary.visaSubmissionDeadline),
    mendanDate: pickStr(primary.mendanDate, secondary.mendanDate),
    interviewDate: pickStr(primary.interviewDate, secondary.interviewDate),
    lastContactDate: pickStr(primary.lastContactDate, secondary.lastContactDate),
    declineReason: pickStr(primary.declineReason, secondary.declineReason),
    memo: mergeUniqueText(primary.memo, secondary.memo),
    nationalityManuallyEdited: Boolean(
      primary.nationalityManuallyEdited || secondary.nationalityManuallyEdited
    ),
    channelManuallyEdited: Boolean(
      primary.channelManuallyEdited || secondary.channelManuallyEdited
    ),
    googleCalendarEventId: pickStr(primary.googleCalendarEventId, secondary.googleCalendarEventId),
    googleCalendarEventLink: pickStr(
      primary.googleCalendarEventLink,
      secondary.googleCalendarEventLink
    ),
    savedRirekisho: primary.savedRirekisho || secondary.savedRirekisho,
    _isSchool: isSch,
  };
};

export const deduplicateCandidatesUnion = (
  candidatesList: Candidate[] = [],
  schoolStudentsList: Candidate[] = []
): Candidate[] => {
  const result: Candidate[] = [];
  const pushOrMerge = (item: Candidate, isSchool: boolean) => {
    const normalizedItem: Candidate = {
      ...item,
      channel: isSchool ? 'school' : item.channel || 'personal',
      _isSchool: isSchool,
    };
    const dupIdx = result.findIndex((existing) =>
      areCandidatesSamePerson(existing, normalizedItem)
    );
    if (dupIdx >= 0) {
      result[dupIdx] = mergeCandidateRecords(result[dupIdx], normalizedItem);
    } else {
      result.push(normalizedItem);
    }
  };

  schoolStudentsList.forEach((s) => pushOrMerge(s, true));
  candidatesList.forEach((c) => {
    const isSch = Boolean(c._isSchool || c.channel === 'school');
    pushOrMerge(c, isSch);
  });
  return result;
};

export interface CleanupCandidatesResult {
  candidates: Candidate[];
  schoolStudents: Candidate[];
  movedToOther: number;
  movedToCompany: number;
  movedToLine: number;
  movedToSchool: number;
  dedupedCount: number;
  changedCount: number;
}

const classifyLegacyNonSchoolChannel = (c: Candidate): ChannelType => {
  const rawCh = String(c.channel || '').trim().toLowerCase();
  if (c.channelManuallyEdited) {
    if (rawCh === 'company' || rawCh === 'line' || rawCh === 'other' || rawCh === 'personal') {
      return rawCh as ChannelType;
    }
  }
  if (rawCh === 'company' || rawCh === 'line' || rawCh === 'other') {
    return rawCh as ChannelType;
  }

  const fbUrl = String(c.fbProfileUrl || '').trim();
  const fbSender = String(c.fbSenderId || '').trim();
  const rawFbName = String(c.fbName || '').trim();
  const cleanFbName =
    rawFbName &&
    !/^(?:-|—|未定|不明|未設定|未記入|未確認|なし|無|none|n\/a|kintone)$/i.test(rawFbName) &&
    !/^https?:\/\//i.test(rawFbName)
      ? rawFbName
      : '';

  const notes = [c.memo, c.negotiationLog, c.conversionGoal, c.id].filter(Boolean).join(' ');
  const companyRx =
    /(?:fanpage|page\s*công\s*ty|fb\s*công\s*ty|会社公式|会社fb|公式fb|公式ページ|企業ページ|自社求人)/i;
  const lineRx =
    /(?:^|[\s:：・(（【\[「『])LINE(?:$|[\s:：・)）】\]」』]|で|から|にて|より|の|経由|相談|応募|登録|連絡|id)|ライン(?:経由|相談|応募)?/i;
  const kintoneOrReferralRx =
    /(?:kintone|キントーン|直接紹介|知人紹介|紹介経由|紹介者|zalo|tiktok|電話応募|メール応募|求職者登録|旧データ|cand_csv_|kintone_)/i;
  const personalFbChatRx =
    /(?:facebook|messenger|メッセンジャー|fb個人|fb\s*cá\s*nhân|fb経由|fbから|link\s*fb|inbox|tên\s*fb|Ghi chú nguyên văn|\bfb\b)/i;

  if (isRealSocialUrl(fbUrl) || Boolean(fbSender)) {
    if (companyRx.test(notes) && !/fb個人|fb\s*cá\s*nhân/i.test(notes)) {
      return 'company';
    }
    return 'personal';
  }

  if (companyRx.test(`${notes} ${cleanFbName}`)) {
    return 'company';
  }
  if (lineRx.test(`${notes} ${cleanFbName} ${c.phone || ''}`)) {
    return 'line';
  }
  if (kintoneOrReferralRx.test(notes)) {
    return 'other';
  }

  const hasExplicitPersonalFbEvidence =
    personalFbChatRx.test(notes) && !companyRx.test(notes);
  if (hasExplicitPersonalFbEvidence) {
    return 'personal';
  }

  if (!cleanFbName || isFbNameSameOrSubsetOfName(cleanFbName, c.name)) {
    return 'other';
  }

  return 'personal';
};

export const cleanupAndReclassifyCandidates = (
  candidatesList: Candidate[] = [],
  schoolStudentsList: Candidate[] = []
): CleanupCandidatesResult => {
  const initialTotal = candidatesList.length + schoolStudentsList.length;
  const schoolMerged: Candidate[] = [];
  const regularMerged: Candidate[] = [];
  let dedupedCount = 0;
  let movedToOther = 0;
  let movedToCompany = 0;
  let movedToLine = 0;
  let movedToSchool = 0;

  const isSchoolRec = (item: Candidate): boolean =>
    Boolean(item._isSchool || item.channel === 'school' || String(item.schoolName || '').trim());

  // 1. Seed school students first (deduplicating intra-school)
  schoolStudentsList.forEach((rawStu) => {
    const stu: Candidate = { ...rawStu, channel: 'school', _isSchool: true };
    const idx = schoolMerged.findIndex((existing) => areCandidatesSamePerson(existing, stu));
    if (idx >= 0) {
      schoolMerged[idx] = mergeCandidateRecords(schoolMerged[idx], stu);
      dedupedCount += 1;
    } else {
      schoolMerged.push(stu);
    }
  });

  // 2. Process candidates: eliminate overlap with schoolStudents, move school records, dedupe regular, reclassify channels
  candidatesList.forEach((rawCand) => {
    const cand: Candidate = { ...rawCand };
    const schIdx = schoolMerged.findIndex((stu) => areCandidatesSamePerson(stu, cand));
    if (schIdx >= 0) {
      schoolMerged[schIdx] = mergeCandidateRecords(schoolMerged[schIdx], cand);
      dedupedCount += 1;
      return;
    }

    if (isSchoolRec(cand)) {
      schoolMerged.push({ ...cand, channel: 'school', _isSchool: true });
      movedToSchool += 1;
      return;
    }

    const regIdx = regularMerged.findIndex((existing) => areCandidatesSamePerson(existing, cand));
    if (regIdx >= 0) {
      regularMerged[regIdx] = mergeCandidateRecords(regularMerged[regIdx], cand);
      dedupedCount += 1;
      return;
    }

    const prevChannel = String(cand.channel || 'personal').trim().toLowerCase() || 'personal';
    const nextChannel = classifyLegacyNonSchoolChannel(cand);
    if (nextChannel !== prevChannel) {
      if (nextChannel === 'other') movedToOther += 1;
      else if (nextChannel === 'company') movedToCompany += 1;
      else if (nextChannel === 'line') movedToLine += 1;
    }
    regularMerged.push({
      ...cand,
      channel: nextChannel,
      _isSchool: false,
    });
  });

  // 3. Re-check regularMerged after merges in case a merged record changed channel
  const finalRegular = regularMerged.map((cand) => {
    const prevChannel = String(cand.channel || 'personal').trim().toLowerCase() || 'personal';
    const nextChannel = classifyLegacyNonSchoolChannel(cand);
    if (nextChannel !== prevChannel) {
      if (nextChannel === 'other') movedToOther += 1;
      else if (nextChannel === 'company') movedToCompany += 1;
      else if (nextChannel === 'line') movedToLine += 1;
    }
    return {
      ...cand,
      channel: nextChannel,
      _isSchool: false,
    };
  });

  const finalTotal = finalRegular.length + schoolMerged.length;
  const changedCount =
    movedToOther +
    movedToCompany +
    movedToLine +
    movedToSchool +
    dedupedCount +
    (finalTotal !== initialTotal ? Math.abs(initialTotal - finalTotal) : 0);

  return {
    candidates: finalRegular,
    schoolStudents: schoolMerged,
    movedToOther,
    movedToCompany,
    movedToLine,
    movedToSchool,
    dedupedCount,
    changedCount,
  };
};

export const isStale7DayProposal = (
  person: Partial<Candidate> = {},
  todayIsoOverride?: string
): boolean => {
  const st = String(person.progressStatus || '').trim();
  if (st !== '企業提案中' && st !== '面接調整中') return false;

  const now = new Date();
  const todayIso =
    todayIsoOverride ||
    `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(
      now.getDate()
    ).padStart(2, '0')}`;
  if (String(person.lastContactDate || '').trim() === todayIso) {
    return false;
  }

  const todayTs = parseFlexibleDateToTimestamp(todayIso, 2026);
  const refDate =
    String(person.lastContactDate || '').trim() ||
    String(person.proposalDate || '').trim();
  const refTs = parseFlexibleDateToTimestamp(refDate, 2026);
  if (!refTs || !todayTs) return false;
  const diffDays = Math.floor((todayTs - refTs) / 86400000);
  return diffDays >= 7;
};

export const detectKyujinShiftType = (shiftText = '', fullText = ''): string => {
  const combined = `${shiftText} ${fullText}`.trim();
  if (!combined) return '';

  const hasNightKeyword =
    /夜勤あり|夜勤有り|夜勤必須|2交代|3交代|二交代|三交代|深夜勤務|22:00以降/.test(combined);
  const noNightExplicit = /夜勤なし|夜勤無|日勤のみ/.test(combined);

  const timeRanges: Array<{ startMin: number; endMin: number }> = [];
  const rangeRx = /(\d{1,2})\s*[:：]\s*(\d{2})\s*[〜～~\-]\s*(\d{1,2})\s*[:：]\s*(\d{2})/g;
  let m: RegExpExecArray | null;
  while ((m = rangeRx.exec(shiftText || combined)) !== null) {
    const startMin = Number(m[1]) * 60 + Number(m[2]);
    const endMin = Number(m[3]) * 60 + Number(m[4]);
    timeRanges.push({ startMin, endMin });
  }

  const hasOvernightShift = timeRanges.some(
    (tr) => tr.endMin < tr.startMin || tr.endMin > 22 * 60 || tr.startMin >= 22 * 60
  );

  if (hasOvernightShift || (hasNightKeyword && !noNightExplicit)) {
    return '夜勤あり';
  }

  const isMultiShift =
    timeRanges.length >= 2 || /シフト制|早番|遅番|中番|[①②③④]/.test(shiftText || combined);
  if (isMultiShift) {
    return 'シフト制（夜勤なし）';
  }
  if (/日勤のみ/.test(combined) || (timeRanges.length === 1 && !hasOvernightShift)) {
    return '日勤のみ';
  }
  return '';
};

export const migrateCandidateRecord = (rawInput: unknown): Candidate => {
  const raw =
    rawInput && typeof rawInput === 'object' && !Array.isArray(rawInput)
      ? (rawInput as Record<string, unknown>)
      : {};
  const schoolName = cleanLegacyStr(raw.schoolName ?? raw.school ?? raw.school_name);
  const rawChannel = cleanLegacyStr(raw.channel).toLowerCase();
  const isSchool =
    rawChannel === 'school' ||
    Boolean(raw._isSchool ?? raw.is_school) ||
    Boolean(schoolName);

  const fbSenderId = cleanLegacyStr(raw.fbSenderId ?? raw.fb_sender_id);
  const fbProfileUrl = cleanLegacyStr(raw.fbProfileUrl ?? raw.fb_profile_url);
  const rawFbName = cleanLegacyStr(raw.fbName ?? raw.fb_name ?? raw.facebookName ?? raw.facebook_name);
  const fbName = isPlaceholderExtractionVal(rawFbName) ? '' : rawFbName;
  const name = cleanLegacyStr(raw.name ?? raw.fullName ?? raw.full_name ?? raw.candidate_name);
  const phone = cleanLegacyStr(raw.phone ?? raw.tel);
  const memo = cleanLegacyStr(raw.memo ?? raw.notes ?? raw.note);
  const negotiationLog = cleanLegacyStr(
    raw.negotiationLog ?? raw.rawLog ?? raw.negotiation_log ?? raw.raw_log
  );
  const channelManuallyEdited = Boolean(
    raw.channelManuallyEdited ?? raw.channel_manually_edited
  );

  let channel: ChannelType | string = 'personal';
  if (isSchool) {
    channel = 'school';
  } else if (rawChannel === 'company' || rawChannel === 'line' || rawChannel === 'other') {
    channel = rawChannel;
  } else if (channelManuallyEdited && rawChannel === 'personal') {
    channel = 'personal';
  } else if (rawChannel === 'personal' || !rawChannel) {
    const distinctFbName = Boolean(fbName) && !isFbNameSameOrSubsetOfName(fbName, name);
    const hasFbSignal = Boolean(
      isRealSocialUrl(fbProfileUrl) || fbSenderId || distinctFbName
    );
    if (hasFbSignal) {
      channel = 'personal';
    } else {
      const hint = `${phone} ${memo} ${negotiationLog}`;
      channel =
        /(?:^|[\s:：・(（【\[「『])LINE(?:$|[\s:：・)）】\]」』]|で|から|にて|より|の|経由|相談|応募|登録|連絡|id)|ライン/i.test(
          hint
        )
          ? 'line'
          : 'other';
    }
  }

  const furigana = cleanLegacyStr(raw.furigana ?? raw.kana);
  const rawNationality = cleanLegacyStr(raw.nationality ?? raw.kokuseki ?? raw.country);
  const nationality = rawNationality ? normalizeCandidateNationality(rawNationality) : '';
  const gender = cleanLegacyStr(raw.gender ?? raw.sex);
  const age = cleanLegacyStr(raw.age);
  const birthDate = cleanLegacyStr(raw.birthDate ?? raw.birth_date ?? raw.dob);
  const postalCode = cleanLegacyStr(raw.postalCode ?? raw.postal_code ?? raw.zip);
  const visaStatus = cleanLegacyStr(raw.visaStatus ?? raw.visaType ?? raw.visa_status ?? raw.visa_type);
  const visaExpiry = cleanLegacyStr(raw.visaExpiry ?? raw.visa_expiry);
  const email = cleanLegacyStr(raw.email ?? raw.mail);
  const address = cleanLegacyStr(raw.address);
  const proposedCompany = cleanLegacyStr(raw.proposedCompany ?? raw.company ?? raw.proposed_company);
  const proposalDate = cleanLegacyStr(raw.proposalDate ?? raw.proposal_date);
  let naiteiCompany = cleanLegacyStr(raw.naiteiCompany ?? raw.naitei_company);
  const progressStatus =
    cleanLegacyStr(raw.progressStatus ?? raw.status ?? raw.progress_status) || '新規相談';
  let declineReason = cleanLegacyStr(
    raw.declineReason ?? raw.decline_reason ?? raw.jitaiReason ?? raw.jitai_reason
  );

  if (
    ['内定', '内定承諾', '入社手続き中', '入社済み'].includes(progressStatus) &&
    !naiteiCompany &&
    proposedCompany
  ) {
    naiteiCompany = proposedCompany;
  }

  if (
    ['内定辞退', '選考中辞退', '不合格', '辞退・終了'].includes(progressStatus) &&
    !declineReason &&
    memo
  ) {
    declineReason = extractDeclineReason({ progressStatus, memo, declineReason: '' });
  }

  const rawSavedCv =
    raw.savedRirekisho ?? raw.saved_rirekisho ?? raw.saved_cv ?? raw.savedCv;

  return {
    ...(raw as Partial<Candidate>),
    id:
      cleanLegacyStr(raw.id) ||
      `cand_mig_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
    fbSenderId,
    fbProfileUrl,
    name,
    fbName,
    furigana,
    channel,
    schoolName,
    graduationDate: cleanLegacyStr(raw.graduationDate ?? raw.graduation_date),
    nationality,
    gender,
    age,
    birthDate,
    postalCode,
    jlpt:
      cleanLegacyStr(raw.jlpt ?? raw.japaneseLevel ?? raw.japanese_level ?? raw.jlpt_level) ||
      '要確認',
    desiredIndustry: cleanLegacyStr(
      raw.desiredIndustry ?? raw.desiredField ?? raw.desired_industry ?? raw.desired_field ?? raw.industry
    ),
    desiredArea: cleanLegacyStr(raw.desiredArea ?? raw.desired_area ?? raw.area),
    conversionGoal: cleanLegacyStr(raw.conversionGoal ?? raw.goal ?? raw.conversion_goal),
    visaStatus,
    visaExpiry,
    passedExam: cleanLegacyStr(raw.passedExam ?? raw.certifications ?? raw.passed_exam),
    japanExperience: cleanLegacyStr(raw.japanExperience ?? raw.japan_experience),
    salaryExpectation: cleanLegacyStr(raw.salaryExpectation ?? raw.salary_expectation),
    negotiationLog,
    startDate: cleanLegacyStr(raw.startDate ?? raw.start_date),
    phone,
    email,
    address,
    progressStatus,
    proposedCompany,
    proposalDate,
    naiteiCompany,
    expectedJoinDate: cleanLegacyStr(raw.expectedJoinDate ?? raw.expected_join_date),
    visaSubmissionDeadline: cleanLegacyStr(
      raw.visaSubmissionDeadline ?? raw.visa_submission_deadline ?? raw.visa_deadline
    ),
    mendanDate: cleanLegacyStr(raw.mendanDate ?? raw.mendan_date),
    interviewDate: cleanLegacyStr(raw.interviewDate ?? raw.interview_date),
    lastContactDate: cleanLegacyStr(raw.lastContactDate ?? raw.last_contact_date),
    declineReason,
    memo,
    nationalityManuallyEdited: Boolean(
      raw.nationalityManuallyEdited ?? raw.nationality_manually_edited
    ),
    channelManuallyEdited,
    googleCalendarEventId: cleanLegacyStr(
      raw.googleCalendarEventId ?? raw.google_calendar_event_id
    ),
    googleCalendarEventLink: cleanLegacyStr(
      raw.googleCalendarEventLink ?? raw.google_calendar_event_link
    ),
    savedRirekisho: sanitizeSavedRirekishoRecord(rawSavedCv, {
      name,
      furigana,
      birthDate,
      age,
      gender,
      nationality,
      visaStatus,
      visaExpiry,
      phone,
      email,
      address,
      postalCode,
    }),
    _isSchool: isSchool,
  };
};

export interface MigratedLegacyDataResult {
  candidates: Candidate[];
  schoolStudents: Candidate[];
  kyujins: KyujinCompany[];
  ptPartners: PtPartner[];
  resignations: Array<Partial<ResignationRecord> & Record<string, unknown>>;
  hasCandidatesSection: boolean;
  hasStudentsSection: boolean;
  hasKyujinsSection: boolean;
  hasPtSection: boolean;
  hasResignationsSection: boolean;
  recognizedKeys: string[];
  warnings: string[];
}

const isRecordObject = (v: unknown): v is Record<string, unknown> =>
  Boolean(v) && typeof v === 'object' && !Array.isArray(v);

const pickFirstArrayField = (
  obj: Record<string, unknown>,
  keys: string[]
): { key: string; arr: unknown[] } | null => {
  for (const key of keys) {
    if (Array.isArray(obj[key])) {
      return { key, arr: obj[key] as unknown[] };
    }
  }
  return null;
};

const looksLikeKyujinArray = (arr: unknown[]): boolean => {
  const sample = arr.find((item) => isRecordObject(item)) as Record<string, unknown> | undefined;
  if (!sample) return false;
  const hasKyujinSignal = Boolean(
    sample.companyName ||
      sample.company_name ||
      Array.isArray(sample.areaBreakdown) ||
      Array.isArray(sample.area_breakdown) ||
      sample.jobCode ||
      sample.job_code ||
      sample.hourlyWage ||
      sample.hourly_wage
  );
  const hasCandidateSignal = Boolean(
    sample.fbName ||
      sample.fb_name ||
      sample.progressStatus ||
      sample.progress_status ||
      sample.jlpt ||
      sample.visaStatus ||
      sample.visa_status
  );
  return hasKyujinSignal && !hasCandidateSignal;
};

export const migrateLegacyData = (rawData: unknown): MigratedLegacyDataResult => {
  if (rawData === null || rawData === undefined) {
    throw new Error('File backup rỗng (null/undefined), không có dữ liệu để khôi phục.');
  }

  // Direct array backup (legacy export containing only candidates/users or kyujins)
  if (Array.isArray(rawData)) {
    const validObjects = rawData.filter((item) => isRecordObject(item));
    if (looksLikeKyujinArray(validObjects)) {
      return {
        candidates: [],
        schoolStudents: [],
        kyujins: validObjects.map((item) => normalizeKyujinItem(item)),
        ptPartners: [],
        resignations: [],
        hasCandidatesSection: false,
        hasStudentsSection: false,
        hasKyujinsSection: true,
        hasPtSection: false,
        hasResignationsSection: false,
        recognizedKeys: ['rootArray(kyujins)'],
        warnings: [],
      };
    }

    const regularCandidates: Candidate[] = [];
    const schoolStudents: Candidate[] = [];
    validObjects.forEach((rawItem) => {
      const migrated = migrateCandidateRecord(rawItem);
      if (migrated.channel === 'school' || migrated.schoolName) {
        schoolStudents.push({ ...migrated, channel: 'school', _isSchool: true });
      } else {
        regularCandidates.push({ ...migrated, _isSchool: false });
      }
    });

    return {
      candidates: regularCandidates,
      schoolStudents,
      kyujins: [],
      ptPartners: [],
      resignations: [],
      hasCandidatesSection: true,
      hasStudentsSection: schoolStudents.length > 0,
      hasKyujinsSection: false,
      hasPtSection: false,
      hasResignationsSection: false,
      recognizedKeys: ['rootArray(candidates)'],
      warnings:
        validObjects.length < rawData.length
          ? [`Đã bỏ qua ${rawData.length - validObjects.length} dòng bị rỗng/null trong mảng gốc.`]
          : [],
    };
  }

  if (!isRecordObject(rawData)) {
    throw new Error(
      `Định dạng JSON không hợp lệ (nhận được kiểu "${typeof rawData}", cần Object hoặc Array).`
    );
  }

  // Unwrap nested { data: {...} } / { payload: {...} } / { backup: {...} } if top level doesn't have CRM arrays
  let rootObj: Record<string, unknown> = rawData;
  const candidateKeys = ['candidates', 'users', 'applicants', 'members', 'candidateList', 'candidate_list'];
  const studentKeys = ['schoolStudents', 'students', 'school_students', 'ryugakusei'];
  const kyujinKeys = ['kyujins', 'jobs', 'companies', 'kyujinList', 'kyujin_list', 'kyujin'];
  const ptKeys = ['ptPartners', 'partners', 'ptList', 'pt_partners', 'pt'];
  const resignationKeys = [
    'resignations',
    'resignationList',
    'resignation_list',
    'resignationsList',
    'taishokusha',
  ];

  const allKnownSectionKeys = [
    ...candidateKeys,
    ...studentKeys,
    ...kyujinKeys,
    ...ptKeys,
    ...resignationKeys,
  ];

  const hasTopLevelMatch = allKnownSectionKeys.some((k) => Array.isArray(rootObj[k]));
  if (!hasTopLevelMatch) {
    for (const wrapperKey of ['data', 'payload', 'backup', 'backupData', 'crmData']) {
      const inner = rootObj[wrapperKey];
      if (isRecordObject(inner) && allKnownSectionKeys.some((k) => Array.isArray(inner[k]))) {
        rootObj = inner;
        break;
      }
      if (Array.isArray(inner)) {
        return migrateLegacyData(inner);
      }
    }
  }

  const foundCandidates = pickFirstArrayField(rootObj, candidateKeys);
  const foundStudents = pickFirstArrayField(rootObj, studentKeys);
  const foundKyujins = pickFirstArrayField(rootObj, kyujinKeys);
  const foundPt = pickFirstArrayField(rootObj, ptKeys);
  const foundResignations = pickFirstArrayField(rootObj, resignationKeys);

  const recognizedKeys = [
    foundCandidates?.key,
    foundStudents?.key,
    foundKyujins?.key,
    foundPt?.key,
    foundResignations?.key,
  ].filter(Boolean) as string[];

  if (recognizedKeys.length === 0) {
    const existingKeys = Object.keys(rawData).slice(0, 10).join(', ') || '(trống)';
    throw new Error(
      `Không nhận diện được cấu trúc backup CRM hợp lệ. Các key tìm thấy trong file: [${existingKeys}]. Cần có ít nhất một mảng như candidates/users, schoolStudents/students, kyujins/jobs/companies, ptPartners hoặc resignations.`
    );
  }

  const warnings: string[] = [];
  const separatedStudents: Candidate[] = [];
  const regularCandidates: Candidate[] = [];

  if (foundStudents) {
    const validStudents = foundStudents.arr.filter((item) => isRecordObject(item));
    if (validStudents.length < foundStudents.arr.length) {
      warnings.push(
        `Đã bỏ qua ${foundStudents.arr.length - validStudents.length} bản ghi du học sinh null/lỗi.`
      );
    }
    validStudents.forEach((rawStudent) => {
      const migrated = migrateCandidateRecord(rawStudent);
      separatedStudents.push({ ...migrated, channel: 'school', _isSchool: true });
    });
  }

  if (foundCandidates) {
    const validCands = foundCandidates.arr.filter((item) => isRecordObject(item));
    if (validCands.length < foundCandidates.arr.length) {
      warnings.push(
        `Đã bỏ qua ${foundCandidates.arr.length - validCands.length} bản ghi ứng viên null/lỗi.`
      );
    }
    validCands.forEach((rawItem) => {
      const migrated = migrateCandidateRecord(rawItem);
      if (migrated.channel === 'school' || migrated.schoolName) {
        separatedStudents.push({ ...migrated, channel: 'school', _isSchool: true });
      } else {
        regularCandidates.push({ ...migrated, _isSchool: false });
      }
    });
  }

  const normalizedKyujins: KyujinCompany[] = foundKyujins
    ? foundKyujins.arr
        .filter((item) => isRecordObject(item))
        .map((item) => normalizeKyujinItem(item))
    : [];

  if (foundKyujins && normalizedKyujins.length < foundKyujins.arr.length) {
    warnings.push(
      `Đã bỏ qua ${foundKyujins.arr.length - normalizedKyujins.length} bản ghi tin tuyển dụng null/lỗi.`
    );
  }

  const normalizedPt: PtPartner[] = foundPt
    ? (foundPt.arr.filter((item) => isRecordObject(item)) as unknown as PtPartner[])
    : [];

  const normalizedResignations: Array<Partial<ResignationRecord> & Record<string, unknown>> =
    foundResignations
      ? foundResignations.arr
          .filter((item) => isRecordObject(item))
          .map((item, idx) => {
            const rec = item as Record<string, unknown>;
            return {
              ...rec,
              id: cleanLegacyStr(rec.id) || `res_mig_${Date.now()}_${idx}`,
              name: cleanLegacyStr(rec.name ?? rec.fullName ?? rec.full_name),
              fbName: cleanLegacyStr(rec.fbName ?? rec.fb_name),
              nationality: cleanLegacyStr(rec.nationality ?? rec.kokuseki),
              previousCompany: cleanLegacyStr(
                rec.previousCompany ?? rec.previous_company ?? rec.company ?? rec.company_name
              ),
              resignationDate: cleanLegacyStr(
                rec.resignationDate ?? rec.resignation_date ?? rec.taishokubi
              ),
              paperworkDeadline: cleanLegacyStr(
                rec.paperworkDeadline ?? rec.paperwork_deadline ?? rec.deadline
              ),
              handoverNote: cleanLegacyStr(
                rec.handoverNote ?? rec.handover_note ?? rec.memo ?? rec.notes
              ),
              paperworkStatus: cleanLegacyStr(
                rec.paperworkStatus ?? rec.paperwork_status ?? rec.status
              ),
            };
          })
      : [];

  return {
    candidates: regularCandidates,
    schoolStudents: separatedStudents,
    kyujins: normalizedKyujins,
    ptPartners: normalizedPt,
    resignations: normalizedResignations,
    hasCandidatesSection: Boolean(foundCandidates),
    hasStudentsSection: Boolean(foundStudents) || separatedStudents.length > 0,
    hasKyujinsSection: Boolean(foundKyujins),
    hasPtSection: Boolean(foundPt),
    hasResignationsSection: Boolean(foundResignations),
    recognizedKeys,
    warnings,
  };
};

export const parseBackupCsvToCandidates = (csvText: string): Candidate[] => {
  const cleanText = String(csvText || '').replace(/^\uFEFF/, '');
  if (!cleanText.trim()) return [];
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;

  for (let i = 0; i < cleanText.length; i++) {
    const ch = cleanText[i];
    const next = cleanText[i + 1];
    if (inQuotes) {
      if (ch === '"' && next === '"') {
        field += '"';
        i += 1;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        field += ch;
      }
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === ',') {
      row.push(field);
      field = '';
    } else if (ch === '\r' && next === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
      i += 1;
    } else if (ch === '\n' || ch === '\r') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else {
      field += ch;
    }
  }
  if (field.length > 0 || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  if (rows.length < 2) return [];

  const headerMap: Record<string, string> = {
    id: 'id',
    name: 'name',
    fullname: 'name',
    氏名: 'name',
    họtên: 'name',
    fbname: 'fbName',
    fb名: 'fbName',
    fbアカウント名: 'fbName',
    fbprofileurl: 'fbProfileUrl',
    fbsenderid: 'fbSenderId',
    channel: 'channel',
    チャネル: 'channel',
    schoolname: 'schoolName',
    学校名: 'schoolName',
    graduationdate: 'graduationDate',
    nationality: 'nationality',
    国籍: 'nationality',
    quốctịch: 'nationality',
    gender: 'gender',
    性別: 'gender',
    age: 'age',
    年齢: 'age',
    birthdate: 'birthDate',
    生年月日: 'birthDate',
    jlpt: 'jlpt',
    日本語: 'jlpt',
    日本語レベル: 'jlpt',
    desiredindustry: 'desiredIndustry',
    希望分野: 'desiredIndustry',
    ngànhnghề: 'desiredIndustry',
    desiredarea: 'desiredArea',
    希望勤務エリア: 'desiredArea',
    希望勤務地: 'desiredArea',
    visastatus: 'visaStatus',
    在留資格: 'visaStatus',
    visaexpiry: 'visaExpiry',
    在留期限: 'visaExpiry',
    passedexam: 'passedExam',
    合格済み試験: 'passedExam',
    japanexperience: 'japanExperience',
    conversiongoal: 'conversionGoal',
    salaryexpectation: 'salaryExpectation',
    negotiationlog: 'negotiationLog',
    startdate: 'startDate',
    phone: 'phone',
    電話番号: 'phone',
    email: 'email',
    address: 'address',
    progressstatus: 'progressStatus',
    進捗ステータス: 'progressStatus',
    ステータス: 'progressStatus',
    trạngthái: 'progressStatus',
    proposedcompany: 'proposedCompany',
    提案先企業: 'proposedCompany',
    côngtyđềxuất: 'proposedCompany',
    naiteicompany: 'naiteiCompany',
    内定先企業: 'naiteiCompany',
    expectedjoindate: 'expectedJoinDate',
    入社予定日: 'expectedJoinDate',
    visasubmissiondeadline: 'visaSubmissionDeadline',
    ビザ書類期限: 'visaSubmissionDeadline',
    mendandate: 'mendanDate',
    面談日: 'mendanDate',
    interviewdate: 'interviewDate',
    面接日: 'interviewDate',
    lastcontactdate: 'lastContactDate',
    最終連絡日: 'lastContactDate',
    declinereason: 'declineReason',
    辞退理由: 'declineReason',
    memo: 'memo',
    メモ: 'memo',
    ghichú: 'memo',
  };

  const headers = rows[0].map((h) => {
    const norm = String(h || '')
      .trim()
      .toLowerCase()
      .replace(/[\s_\-\/()（）]+/g, '');
    return headerMap[norm] || String(h || '').trim();
  });

  return rows
    .slice(1)
    .filter((r) => r.some((cell) => String(cell || '').trim().length > 0))
    .map((r, idx) => {
      const obj: Record<string, string> = {};
      headers.forEach((key, colIdx) => {
        if (key) {
          obj[key] = cleanLegacyStr(r[colIdx]);
        }
      });
      if (!obj.id) {
        obj.id = `cand_csv_${Date.now()}_${idx}`;
      }
      return migrateCandidateRecord(obj);
    });
};

export const isActualProposedRecord = (person: Partial<Candidate> = {}): boolean => {
  const st = cleanLegacyStr(person.progressStatus);
  if (
    [
      '企業提案中',
      '1次面接予定',
      '2次面接予定',
      '面接調整中',
      '結果待ち',
      '内定',
      '内定承諾',
      '入社手続き中',
      '入社済み',
      '内定辞退',
      '選考中辞退',
      '不合格',
    ].includes(st)
  ) {
    return true;
  }
  if (st === '辞退・終了' && Boolean(extractDeclineReason(person))) {
    return true;
  }
  return false;
};

export const extractDeclineReason = (person: Partial<Candidate> = {}): string => {
  const st = String(person.progressStatus || '').trim();
  const memo = String(person.memo || '').trim();
  const direct = String(person.declineReason || '').trim();
  if (direct) return direct;

  const isDeclined = st.includes('辞退') || /辞退|キャンセル|不参加|見送り|他社決定/.test(memo);
  if (!isDeclined) return '';

  const lines = memo
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  for (let i = lines.length - 1; i >= 0; i--) {
    const line = lines[i];
    const explicitMatch = line.match(/(?:辞退理由|理由|キャンセル理由)[：:\s]*(.+)/);
    if (explicitMatch && explicitMatch[1]) {
      return explicitMatch[1].trim();
    }
    if (/辞退|キャンセル|他社|見送り|条件合わず|連絡途絶|帰国/.test(line)) {
      return line.replace(/^\[[^\]]+\]\s*/, '').trim();
    }
  }
  return st.includes('辞退') ? 'ステータス：辞退・終了（詳細メモ未記入）' : '';
};

export const getUrgentFollowUpReason = (
  person: Partial<Candidate> = {},
  lang: 'vi' | 'jp' = 'vi'
): string => {
  const now = new Date();
  const todayIso = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(
    now.getDate()
  ).padStart(2, '0')}`;
  if (String(person.lastContactDate || '').trim() === todayIso) {
    return '';
  }

  if (isStale7DayProposal(person, todayIso)) {
    return lang === 'jp' ? '⏰ 7日経過・要確認' : '⏰ Quá 7 ngày chưa phản hồi';
  }

  const st = String(person.progressStatus || '').trim();
  if (['保留・連絡待ち', '面談調整中', '面接調整中', 'ヒアリング中'].includes(st)) {
    return lang === 'jp' ? `ステータス：${st}` : `Trạng thái: ${st}`;
  }
  const mendan = extractNormalizedMendanInfo(person);
  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  if (mendan.timestamp > 0) {
    const diffDays = Math.round((mendan.timestamp - todayStart) / 86400000);
    if (diffDays >= 0 && diffDays <= 3) {
      return lang === 'jp'
        ? `面談/PV直前（${mendan.dateLabel}）`
        : `Sắp Mendan/PV (${mendan.dateLabel})`;
    }
    if (diffDays < 0 && !['内定', '入社手続き中', '辞退・終了'].includes(st)) {
      return lang === 'jp'
        ? `予定経過・要状況確認（${mendan.dateLabel}）`
        : `Đã qua lịch hẹn (${mendan.dateLabel})`;
    }
  }
  if (!['内定', '入社手続き中', '辞退・終了'].includes(st)) {
    const lastTs = parseFlexibleDateToTimestamp(person.lastContactDate, 2026);
    if (!lastTs) return lang === 'jp' ? '最終連絡日未記録' : 'Chưa ghi ngày liên hệ cuối';
    const daysSince = Math.floor((todayStart - lastTs) / 86400000);
    if (daysSince >= 7) {
      return lang === 'jp' ? `${daysSince}日間連絡なし` : `${daysSince} ngày chưa liên hệ`;
    }
  }
  return '';
};

const jlptRank = (lvl = ''): number => {
  const s = String(lvl || '').trim().toUpperCase();
  if (!s || s === '要確認' || s === '不明') return -1;
  if (s.includes('N1')) return 5;
  if (s.includes('N2')) return 4;
  if (s.includes('N3')) return 3;
  if (s.includes('N4')) return 2;
  if (s.includes('N5') || s.includes('日常') || s.includes('初級')) return 1;
  if (s.includes('JFT')) return 2;
  if (s.includes('不問')) return 0;
  return -1;
};

interface PrefecturePattern {
  rx: RegExp;
  label: string;
  shortLabel: string;
  keys: string[];
}

const PREFECTURE_PATTERNS: PrefecturePattern[] = [
  {
    rx: /\btokyo\b|東京都|都内|(?:^|[^京都])東京|23区|新宿区|渋谷区|杉並区|立川市|足立区|八王子市|世田谷区|品川区|大田区|中野区|練馬区|板橋区|豊島区|港区|中央区|千代田区/i,
    label: '東京都（Tokyo）',
    shortLabel: '東京都',
    keys: ['東京都', '東京', '都内', 'tokyo'],
  },
  {
    rx: /\bkyoto\b|京都府|(?:^|[^東])京都/i,
    label: '京都府（Kyoto）',
    shortLabel: '京都府',
    keys: ['京都府', 'kyoto'],
  },
  { rx: /\bosaka\b|大阪府|大阪/i, label: '大阪府（Osaka）', shortLabel: '大阪府', keys: ['大阪', 'osaka'] },
  { rx: /\bhyogo\b|\bkobe\b|兵庫県|神戸/i, label: '兵庫県（Hyogo）', shortLabel: '兵庫県', keys: ['兵庫', '神戸', 'hyogo', 'kobe'] },
  { rx: /\bnara\b|奈良県|奈良/i, label: '奈良県（Nara）', shortLabel: '奈良県', keys: ['奈良', 'nara'] },
  { rx: /\bshiga\b|滋賀県|滋賀/i, label: '滋賀県（Shiga）', shortLabel: '滋賀県', keys: ['滋賀', 'shiga'] },
  { rx: /\bwakayama\b|和歌山県|和歌山/i, label: '和歌山県（Wakayama）', shortLabel: '和歌山県', keys: ['和歌山', 'wakayama'] },
  { rx: /\baichi\b|\bnagoya\b|愛知県|名古屋/i, label: '愛知県（Aichi）', shortLabel: '愛知県', keys: ['愛知', '名古屋', 'aichi', 'nagoya'] },
  { rx: /\bshizuoka\b|静岡県|静岡/i, label: '静岡県（Shizuoka）', shortLabel: '静岡県', keys: ['静岡', 'shizuoka'] },
  { rx: /\bfukuoka\b|福岡県|福岡/i, label: '福岡県（Fukuoka）', shortLabel: '福岡県', keys: ['福岡', 'fukuoka'] },
  { rx: /\bokinawa\b|沖縄県|沖縄/i, label: '沖縄県（Okinawa）', shortLabel: '沖縄県', keys: ['沖縄', 'okinawa'] },
  { rx: /\bhokkaido\b|\bsapporo\b|北海道|札幌/i, label: '北海道（Hokkaido）', shortLabel: '北海道', keys: ['北海道', '札幌', 'hokkaido', 'sapporo'] },
  { rx: /\bmiyagi\b|\bsendai\b|宮城県|仙台/i, label: '宮城県（Miyagi）', shortLabel: '宮城県', keys: ['宮城', '仙台', 'miyagi', 'sendai'] },
  { rx: /\bhiroshima\b|広島県|広島/i, label: '広島県（Hiroshima）', shortLabel: '広島県', keys: ['広島', 'hiroshima'] },
  { rx: /\bkanagawa\b|\byokohama\b|\bkawasaki\b|神奈川県|神奈川|横浜|川崎/i, label: '神奈川県（Kanagawa）', shortLabel: '神奈川県', keys: ['神奈川', '横浜', '川崎', 'kanagawa', 'yokohama', 'kawasaki'] },
  { rx: /\bsaitama\b|埼玉県|埼玉/i, label: '埼玉県（Saitama）', shortLabel: '埼玉県', keys: ['埼玉', 'saitama'] },
  { rx: /\bchiba\b|千葉県|千葉/i, label: '千葉県（Chiba）', shortLabel: '千葉県', keys: ['千葉', 'chiba'] },
  { rx: /\bibaraki\b|茨城県|茨城/i, label: '茨城県（Ibaraki）', shortLabel: '茨城県', keys: ['茨城', 'ibaraki'] },
  { rx: /\btochigi\b|栃木県|栃木/i, label: '栃木県（Tochigi）', shortLabel: '栃木県', keys: ['栃木', 'tochigi'] },
  { rx: /\bgunma\b|群馬県|群馬/i, label: '群馬県（Gunma）', shortLabel: '群馬県', keys: ['群馬', 'gunma'] },
  { rx: /\bnagano\b|長野県|長野/i, label: '長野県（Nagano）', shortLabel: '長野県', keys: ['長野', 'nagano'] },
  { rx: /\bgifu\b|岐阜県|岐阜/i, label: '岐阜県（Gifu）', shortLabel: '岐阜県', keys: ['岐阜', 'gifu'] },
  { rx: /\bmie\b|三重県|三重/i, label: '三重県（Mie）', shortLabel: '三重県', keys: ['三重', 'mie'] },
  { rx: /\bniigata\b|新潟県|新潟/i, label: '新潟県（Niigata）', shortLabel: '新潟県', keys: ['新潟', 'niigata'] },
  { rx: /\bokayama\b|岡山県|岡山/i, label: '岡山県（Okayama）', shortLabel: '岡山県', keys: ['岡山', 'okayama'] },
  { rx: /\bkumamoto\b|熊本県|熊本/i, label: '熊本県（Kumamoto）', shortLabel: '熊本県', keys: ['熊本', 'kumamoto'] },
  { rx: /\bkagoshima\b|鹿児島県|鹿児島/i, label: '鹿児島県（Kagoshima）', shortLabel: '鹿児島県', keys: ['鹿児島', 'kagoshima'] },
  { rx: /\baomori\b|青森県|青森/i, label: '青森県（Aomori）', shortLabel: '青森県', keys: ['青森', 'aomori'] },
  { rx: /\biwate\b|岩手県|岩手/i, label: '岩手県（Iwate）', shortLabel: '岩手県', keys: ['岩手', 'iwate'] },
  { rx: /\bakita\b|秋田県|秋田/i, label: '秋田県（Akita）', shortLabel: '秋田県', keys: ['秋田', 'akita'] },
  { rx: /\byamagata\b|山形県|山形/i, label: '山形県（Yamagata）', shortLabel: '山形県', keys: ['山形', 'yamagata'] },
  { rx: /\bfukushima\b|福島県|福島/i, label: '福島県（Fukushima）', shortLabel: '福島県', keys: ['福島', 'fukushima'] },
  { rx: /\byamanashi\b|山梨県|山梨/i, label: '山梨県（Yamanashi）', shortLabel: '山梨県', keys: ['山梨', 'yamanashi'] },
  { rx: /\btoyama\b|富山県|富山/i, label: '富山県（Toyama）', shortLabel: '富山県', keys: ['富山', 'toyama'] },
  { rx: /\bishikawa\b|石川県|石川/i, label: '石川県（Ishikawa）', shortLabel: '石川県', keys: ['石川', 'ishikawa'] },
  { rx: /\bfukui\b|福井県|福井/i, label: '福井県（Fukui）', shortLabel: '福井県', keys: ['福井', 'fukui'] },
  { rx: /\btottori\b|鳥取県|鳥取/i, label: '鳥取県（Tottori）', shortLabel: '鳥取県', keys: ['鳥取', 'tottori'] },
  { rx: /\bshimane\b|島根県|島根/i, label: '島根県（Shimane）', shortLabel: '島根県', keys: ['島根', 'shimane'] },
  { rx: /\byamaguchi\b|山口県|山口/i, label: '山口県（Yamaguchi）', shortLabel: '山口県', keys: ['山口', 'yamaguchi'] },
  { rx: /\btokushima\b|徳島県|徳島/i, label: '徳島県（Tokushima）', shortLabel: '徳島県', keys: ['徳島', 'tokushima'] },
  { rx: /\bkagawa\b|香川県|香川/i, label: '香川県（Kagawa）', shortLabel: '香川県', keys: ['香川', 'kagawa'] },
  { rx: /\behime\b|愛媛県|愛媛/i, label: '愛媛県（Ehime）', shortLabel: '愛媛県', keys: ['愛媛', 'ehime'] },
  { rx: /\bkochi\b|高知県|高知/i, label: '高知県（Kochi）', shortLabel: '高知県', keys: ['高知', 'kochi'] },
  { rx: /\bsaga\b|佐賀県|佐賀/i, label: '佐賀県（Saga）', shortLabel: '佐賀県', keys: ['佐賀', 'saga'] },
  { rx: /\bnagasaki\b|長崎県|長崎/i, label: '長崎県（Nagasaki）', shortLabel: '長崎県', keys: ['長崎', 'nagasaki'] },
  { rx: /\boita\b|大分県|大分/i, label: '大分県（Oita）', shortLabel: '大分県', keys: ['大分', 'oita'] },
  { rx: /\bmiyazaki\b|宮崎県|宮崎/i, label: '宮崎県（Miyazaki）', shortLabel: '宮崎県', keys: ['宮崎', 'miyazaki'] },
];

const REGION_PREFECTURE_GROUPS: Array<{ rx: RegExp; prefectures: string[] }> = [
  {
    rx: /kanto|関東/i,
    prefectures: ['東京都', '神奈川県', '埼玉県', '千葉県', '茨城県', '栃木県', '群馬県'],
  },
  {
    rx: /kansai|kinki|関西|近畿/i,
    prefectures: ['大阪府', '京都府', '兵庫県', '奈良県', '滋賀県', '和歌山県'],
  },
  {
    rx: /tokai|chubu|東海|中部/i,
    prefectures: ['愛知県', '静岡県', '岐阜県', '三重県', '長野県'],
  },
  {
    rx: /kyushu|九州/i,
    prefectures: ['福岡県', '熊本県', '鹿児島県', '佐賀県', '長崎県', '大分県', '宮崎県'],
  },
  {
    rx: /tohoku|東北/i,
    prefectures: ['宮城県', '福島県', '青森県', '岩手県', '秋田県', '山形県'],
  },
];

const doesTextMatchPrefecture = (text: string, pref: PrefecturePattern): boolean =>
  pref.rx.test(String(text || ''));

export const extractAllAreaTokens = (
  rawArea = ''
): Array<{ label: string; shortLabel: string; keys: string[]; rx: RegExp }> => {
  const s = String(rawArea || '').trim();
  if (!s) return [];
  const matched: Array<{ label: string; shortLabel: string; keys: string[]; rx: RegExp }> = [];
  for (const pref of PREFECTURE_PATTERNS) {
    if (doesTextMatchPrefecture(s, pref)) {
      matched.push({
        label: pref.label,
        shortLabel: pref.shortLabel,
        keys: pref.keys,
        rx: pref.rx,
      });
    }
  }
  return matched;
};

export const doesAreaMatchFilter = (itemArea = '', filterText = ''): boolean => {
  const f = String(filterText || '').trim().toLowerCase();
  if (!f) return true;
  const targetRaw = String(itemArea || '');
  const target = targetRaw.toLowerCase();
  if (!target) return false;

  const filterPrefs = extractAllAreaTokens(f);
  if (filterPrefs.length > 0) {
    return filterPrefs.some((fp) => fp.rx.test(targetRaw));
  }

  for (const reg of REGION_PREFECTURE_GROUPS) {
    if (reg.rx.test(f)) {
      const targetPrefs = extractAllAreaTokens(targetRaw);
      if (targetPrefs.some((tp) => reg.prefectures.includes(tp.shortLabel))) {
        return true;
      }
    }
  }

  return target.includes(f);
};

type CanonicalIndustryKey =
  | 'KAIGO'
  | 'FOOD_SERVICE'
  | 'FOOD_MFG'
  | 'TRANSPORT'
  | 'AUTO_MAINT'
  | 'ACCOMMODATION'
  | 'CONSTRUCTION'
  | 'MANUFACTURING'
  | 'CLEANING'
  | 'AGRICULTURE'
  | 'RETAIL_BAITO';

const extractCanonicalIndustryKeys = (rawIndustry = ''): CanonicalIndustryKey[] => {
  const s = String(rawIndustry || '').trim();
  if (!s || /^(?:未設定|未定|不問|要確認)$/i.test(s)) return [];

  const keys = new Set<CanonicalIndustryKey>();
  if (/介護|福祉|デイサービス|グループホーム|有料老人|特別養護|特養|老健|訪問介護|kaigo|điều\s*dưỡng|hộ\s*lý/i.test(s)) {
    keys.add('KAIGO');
  }
  if (/外食|レストラン|居酒屋|ホール|調理|キッチン|カフェ|ngoại\s*thực|nhà\s*hàng|quán\s*ăn|ẩm\s*thực/i.test(s)) {
    keys.add('FOOD_SERVICE');
  }
  if (/飲食料品|食品|弁当|惣菜|製パン|製菓|chế\s*biến\s*thực\s*phẩm|thực\s*phẩm|cơm\s*hộp|bento/i.test(s)) {
    keys.add('FOOD_MFG');
  }
  if (/自動車運送|運送|トラック|ドライバー|配送|運転|vận\s*tải|tài\s*xế|lái\s*xe|logistics/i.test(s)) {
    keys.add('TRANSPORT');
  }
  if (/自動車整備|整備士|bảo\s*dưỡng\s*ô\s*tô|sửa\s*chữa\s*ô\s*tô/i.test(s)) {
    keys.add('AUTO_MAINT');
  }
  if (/宿泊|ホテル|旅館|フロント|khách\s*sạn|lưu\s*trú/i.test(s)) {
    keys.add('ACCOMMODATION');
  }
  if (/建設|建築|土木|とび|型枠|鉄筋|内装|xây\s*dựng|giàn\s*giáo|cốp\s*pha/i.test(s)) {
    keys.add('CONSTRUCTION');
  }
  // Avoid matching generic "製造" inside "飲食料品製造業" / "食品製造"
  const strippedFoodMfg = s.replace(/飲食料品製造業?|食品製造業?/g, '');
  if (/素形材|産業機械|電気電子|機械|溶接|プレス|旋盤|cnc|cơ\s*khí|chế\s*tạo|gia\s*công\s*kim\s*loại|製造/i.test(strippedFoodMfg)) {
    keys.add('MANUFACTURING');
  }
  if (/ビルクリーニング|清掃|vệ\s*sinh\s*tòa\s*nhà|dọn\s*dẹp/i.test(s)) {
    keys.add('CLEANING');
  }
  if (/農業|耕種|畜産|nông\s*nghiệp|trồng\s*trọt|chăn\s*nuôi/i.test(s)) {
    keys.add('AGRICULTURE');
  }
  if (/コンビニ|小売|レジ|bán\s*lẻ|konbini|siêu\s*thị/i.test(s)) {
    keys.add('RETAIL_BAITO');
  }

  return Array.from(keys);
};

const TOKYO_23_WARDS_RX =
  /23区|新宿|渋谷|杉並|足立|世田谷|豊島|港|中央|千代田|品川|大田|中野|練馬|板橋|北|荒川|台東|墨田|江東|江戸川|葛飾|目黒|文京/;

const CANONICAL_NAT_ALIASES: Array<{ canonical: string; rx: RegExp; ngLabel: string }> = [
  { canonical: 'ベトナム', rx: /ベトナム|viet\s*nam|việt\s*nam|\bvn\b/i, ngLabel: 'ベトナム人NG' },
  { canonical: 'ネパール', rx: /ネパール|nepal/i, ngLabel: 'ネパール人NG' },
  { canonical: 'インドネシア', rx: /インドネシア|indonesia/i, ngLabel: 'インドネシア人NG' },
  { canonical: 'ミャンマー', rx: /ミャンマー|myanmar|burma/i, ngLabel: 'ミャンマー人NG' },
  { canonical: 'フィリピン', rx: /フィリピン|philippines?|filipino/i, ngLabel: 'フィリピン人NG' },
  { canonical: 'カンボジア', rx: /カンボジア|cambodia|campuchia/i, ngLabel: 'カンボジア人NG' },
  { canonical: '中国', rx: /中国|china|trung\s*quốc/i, ngLabel: '中国人NG' },
  { canonical: 'バングラデシュ', rx: /バングラ|bangladesh/i, ngLabel: 'バングラデシュ人NG' },
  { canonical: 'スリランカ', rx: /スリランカ|sri\s*lanka/i, ngLabel: 'スリランカ人NG' },
  { canonical: 'タイ', rx: /タイ人|タイ国籍|thailand|thái\s*lan/i, ngLabel: 'タイ人NG' },
  { canonical: 'インド', rx: /インド人|インド国籍|\bindia\b|ấn\s*độ/i, ngLabel: 'インド人NG' },
  { canonical: 'モンゴル', rx: /モンゴル|mongolia|mông\s*cổ/i, ngLabel: 'モンゴル人NG' },
  { canonical: 'ウズベキスタン', rx: /ウズベキスタン|uzbekistan/i, ngLabel: 'ウズベキスタン人NG' },
];

export const normalizeCanonicalNationality = (raw = ''): string => {
  const s = String(raw || '').trim();
  if (!s) return '';
  for (const entry of CANONICAL_NAT_ALIASES) {
    if (entry.rx.test(s)) return entry.canonical;
  }
  return s.replace(/人$/, '').trim();
};

export const normalizeCanonicalGender = (raw = ''): '男性' | '女性' | '' => {
  const s = String(raw || '').trim();
  if (!s) return '';
  if (/女|女性|female|\bnữ\b|\bnu\b/i.test(s)) return '女性';
  if (/男|男性|male|\bnam\b/i.test(s)) return '男性';
  return '';
};

export const isCandidateOverseas = (cand: Partial<Candidate> = {}): boolean => {
  const visa = String(cand.visaStatus || '').trim();
  const addr = String(cand.address || '').trim();
  const notes = `${cand.conversionGoal || ''} ${cand.memo || ''} ${cand.negotiationLog || ''}`;
  const hasDomesticJapanVisa =
    /特定技能|留学|技能実習|技術|人文|国際|定住|永住|家族滞在|特定活動|配偶者/.test(visa);
  const hasDomesticJapanAddr =
    /東京都|北海道|大阪府|京都府|.{2,3}県|区|市/.test(addr);
  if (hasDomesticJapanVisa || hasDomesticJapanAddr) return false;
  if (/海外|国外|ベトナム在住|ミャンマー在住|ネパール在住|インドネシア在住|chưa\s*sang\s*nhật|đang\s*ở\s*(?:việt\s*nam|vn|nước\s*ngoài)|từ\s*việt\s*nam/i.test(`${visa} ${addr} ${notes}`)) {
    return true;
  }
  return false;
};

export const extractNormalizedKyujinConditions = (params: {
  text?: string;
  minJlpt?: string;
  employmentType?: string;
  existingExcludedNationalities?: string[];
  existingExcludedGenders?: string[];
  existingRequiresInJapan?: boolean;
  existingRequiresTokutei?: boolean;
  existingRequiredConditions?: string[];
  existingNgConditions?: string[];
}) => {
  const text = String(params.text || '');
  const excludedNats = new Set<string>(
    (params.existingExcludedNationalities || [])
      .map((n) => normalizeCanonicalNationality(n))
      .filter(Boolean)
  );
  const excludedGenders = new Set<string>(
    (params.existingExcludedGenders || [])
      .map((g) => normalizeCanonicalGender(g))
      .filter(Boolean)
  );
  const ngLabels = new Set<string>(
    (params.existingNgConditions || []).map((s) => String(s || '').trim()).filter(Boolean)
  );
  const reqLabels = new Set<string>(
    (params.existingRequiredConditions || []).map((s) => String(s || '').trim()).filter(Boolean)
  );

  // Detect NG nationalities in Japanese / Vietnamese text
  const natNgChunks: string[] = [];
  const ngClauseRx =
    /([^\n。、()（）]{1,40})\s*(?:人|国籍)?\s*(?:NG|不可|採用不可|応募不可|対象外|受け入れ不可)/gi;
  let m: RegExpExecArray | null;
  while ((m = ngClauseRx.exec(text)) !== null) {
    natNgChunks.push(m[1]);
  }
  const reverseNgRx =
    /(?:NG|不可|không\s*nhận|từ\s*chối|không\s*tuyển)\s*[：:\s]*([^\n。()（）]{1,50})/gi;
  while ((m = reverseNgRx.exec(text)) !== null) {
    natNgChunks.push(m[1]);
  }

  for (const chunk of natNgChunks) {
    for (const entry of CANONICAL_NAT_ALIASES) {
      if (entry.rx.test(chunk)) {
        excludedNats.add(entry.canonical);
        ngLabels.add(entry.ngLabel);
      }
    }
  }
  // Direct individual check e.g. "ネパール人NG" / "ベトナムNG"
  for (const entry of CANONICAL_NAT_ALIASES) {
    const directRx = new RegExp(
      `(?:${entry.rx.source})[人国籍]?\\s*(?:NG|不可|対象外)`,
      'i'
    );
    if (directRx.test(text)) {
      excludedNats.add(entry.canonical);
      ngLabels.add(entry.ngLabel);
    }
    if (excludedNats.has(entry.canonical)) {
      ngLabels.add(entry.ngLabel);
    }
  }

  // Detect NG genders
  if (
    /男性\s*(?:NG|不可|対象外)|男子\s*(?:NG|不可)|女性(?:のみ|限定|専用)|nam\s*NG|không\s*nhận\s*nam|chỉ\s*tuyển\s*nữ/i.test(
      text
    )
  ) {
    excludedGenders.add('男性');
    ngLabels.add('男性NG');
  }
  if (
    /女性\s*(?:NG|不可|対象外)|女子\s*(?:NG|不可)|男性(?:のみ|限定|専用)|nữ\s*NG|không\s*nhận\s*nữ|chỉ\s*tuyển\s*nam/i.test(
      text
    )
  ) {
    excludedGenders.add('女性');
    ngLabels.add('女性NG');
  }

  // Detect Domestic / Overseas NG
  const requiresInJapan =
    Boolean(params.existingRequiresInJapan) ||
    /国内在住|日本在住|国内限定|在日限定|日本国内のみ|海外招へい\s*(?:不可|NG|×)|海外\s*(?:NG|不可)|国外\s*(?:NG|不可)|chỉ\s*tuyển\s*(?:tại|trong\s*nước)\s*Nhật|ở\s*trong\s*nước\s*Nhật/i.test(
      text
    );
  if (requiresInJapan) {
    reqLabels.add('日本国内在住');
    if (/海外招へい\s*(?:不可|NG|×)|海外\s*(?:NG|不可)|国外\s*(?:NG|不可)/i.test(text)) {
      ngLabels.add('海外NG');
    }
  }

  // Detect extra general NG tags (tattoo, smoking, age, driver license, etc.)
  if (/タトゥー\s*(?:NG|不可)|刺青\s*(?:NG|不可)|hình\s*xăm\s*NG/i.test(text)) {
    ngLabels.add('タトゥーNG');
  }

  // Detect Tokutei requirement
  const empType = String(params.employmentType || '');
  const isBaitoOnly = /baito|アルバイト|パート/i.test(empType) && !/特定技能|tokutei/i.test(text);
  const requiresTokutei =
    Boolean(params.existingRequiresTokutei) ||
    (!isBaitoOnly &&
      (/特定技能|tokutei/i.test(empType) || /特定技能|技能測定試験|評価試験合格/.test(text)));
  if (requiresTokutei) {
    reqLabels.add('特定技能');
  }

  // Normalize JLPT requirement into requiredConditions
  const jlptStr = String(params.minJlpt || '').trim();
  if (/^N[1-5]$/i.test(jlptStr)) {
    reqLabels.add(`JLPT ${jlptStr.toUpperCase()}以上`);
  } else if (jlptStr === 'JFT') {
    reqLabels.add('JFT-Basic合格');
  } else if (jlptStr === '日常会話') {
    reqLabels.add('日常会話レベル');
  }

  return {
    excludedNationalities: Array.from(excludedNats),
    excludedGenders: Array.from(excludedGenders),
    requiresInJapan,
    requiresTokutei,
    requiredConditions: Array.from(reqLabels),
    ngConditions: Array.from(ngLabels),
  };
};

export const evaluateCandidateKyujinMatches = (
  cand: Partial<Candidate> = {},
  kyujinsList: KyujinCompany[] = [],
  lang: 'vi' | 'jp' = 'vi'
): KyujinMatchEvaluation => {
  const candInd = String(cand.desiredIndustry || '').trim();
  const candArea = String(cand.desiredArea || '').trim();
  const candGoal = String(cand.conversionGoal || '').trim();
  const candExp = String(cand.japanExperience || '').trim();
  const candPassedExam = String(cand.passedExam || '').trim();
  const candNat = normalizeCanonicalNationality(cand.nationality || '');
  const candGender = normalizeCanonicalGender(cand.gender || '');
  const candOverseas = isCandidateOverseas(cand);
  const candSalaryExp = String(cand.salaryExpectation || '').trim();
  const candJlptScore = jlptRank(cand.jlpt);

  const candIndKeys = extractCanonicalIndustryKeys(candInd);
  const hasCandidateIndustry =
    Boolean(candInd) && !/^(?:未設定|未定|不問|要確認)$/i.test(candInd);
  const hasCandidateArea =
    Boolean(candArea) && !/^(?:未設定|未定|不問|要確認)$/i.test(candArea);

  const candAreaPrefs = extractAllAreaTokens(candArea);
  const candRegions = REGION_PREFECTURE_GROUPS.filter((r) => r.rx.test(candArea));

  const explicitBaitoTargetRx =
    /baito|アルバイト|パート|làm\s*thêm|bán\s*thời\s*gian|part[- ]?time/i;
  const seekingBaitoInNotesRx =
    /(?:tìm|muốn|xin|hỏi|cần|đang\s*tìm|chuyển|kiếm|nguyện\s*vọng)[^.\n]{0,30}(?:baito|làm\s*thêm|アルバイト|パート)/i;
  const mentionsBaito =
    explicitBaitoTargetRx.test(`${candGoal} ${candInd}`) ||
    seekingBaitoInNotesRx.test(`${cand.memo || ''} ${cand.negotiationLog || ''}`);
  const mentionsTokuteiOrFullTime =
    /tokutei|特定技能|正社員|chính\s*thức|chuyển\s*(?:đổi|visa|việc)|特定/i.test(
      `${candGoal} ${candInd} ${cand.visaStatus || ''} ${candPassedExam}`
    ) || !mentionsBaito;

  const needsHousing = /寮|社宅|nhà\s*ở|ký\s*túc/i.test(
    `${candGoal} ${candSalaryExp} ${cand.memo || ''}`
  );

  const results: KyujinMatchResult[] = [];
  (kyujinsList || []).forEach((comp) => {
    if (comp.recruitmentStatus === '募集終了') return;

    const compConds = extractNormalizedKyujinConditions({
      text: `${comp.companyName || ''}`,
      employmentType: comp.employmentType,
      existingExcludedNationalities: comp.excludedNationalities,
      existingExcludedGenders: comp.excludedGenders,
      existingRequiresInJapan: comp.requiresInJapan,
      existingRequiresTokutei: comp.requiresTokutei,
      existingRequiredConditions: comp.requiredConditions,
      existingNgConditions: comp.ngConditions,
    });

    (comp.areaBreakdown || []).forEach((sl) => {
      if (sl.recruitmentStatus === '募集終了') return;

      const slotConds = extractNormalizedKyujinConditions({
        text: `${sl.jobDescription || ''} ${sl.experienceDetail || ''} ${sl.careQualification || ''} ${sl.facilityType || ''}`,
        minJlpt: sl.minJlpt,
        employmentType: sl.employmentType || comp.employmentType,
        existingExcludedNationalities: [
          ...compConds.excludedNationalities,
          ...(sl.excludedNationalities || []),
        ],
        existingExcludedGenders: [
          ...compConds.excludedGenders,
          ...(sl.excludedGenders || []),
        ],
        existingRequiresInJapan: Boolean(compConds.requiresInJapan || sl.requiresInJapan),
        existingRequiresTokutei: Boolean(compConds.requiresTokutei || sl.requiresTokutei),
        existingRequiredConditions: sl.requiredConditions,
        existingNgConditions: [...compConds.ngConditions, ...(sl.ngConditions || [])],
      });

      // Strict NG filtering: hide any job whose NG exclusions match the candidate
      if (candNat && slotConds.excludedNationalities.includes(candNat)) {
        return;
      }
      if (candGender && slotConds.excludedGenders.includes(candGender)) {
        return;
      }
      if (slotConds.requiresInJapan && candOverseas) {
        return;
      }

      // 1. Strict Industry Filter (NEVER suggest mismatched industries!)
      const slotInd = String(sl.industry || comp.industry || '').trim();
      const slotIndSource = `${slotInd} ${sl.facilityType || ''}`.trim();
      const slotIndKeys = extractCanonicalIndustryKeys(slotIndSource);

      let industryMatched = false;
      if (hasCandidateIndustry) {
        if (candIndKeys.length > 0 && slotIndKeys.length > 0) {
          industryMatched = candIndKeys.some((k) => slotIndKeys.includes(k));
        } else {
          const normCand = candInd.toLowerCase();
          const normSlot = slotInd.toLowerCase();
          industryMatched =
            Boolean(normSlot) &&
            (normSlot.includes(normCand) || normCand.includes(normSlot));
        }
        if (!industryMatched) {
          return;
        }
      }

      // 2. Strict Employment Type Filter (NEVER suggest Baito to Tokutei-only candidates)
      const empType = String(
        sl.employmentType || comp.employmentType || 'Tokutei Ginou (特定技能)'
      );
      const slotIsBaito = /baito|アルバイト|パート/i.test(empType);
      if (!mentionsBaito && slotIsBaito) {
        return;
      }
      if (mentionsBaito && !mentionsTokuteiOrFullTime && !slotIsBaito) {
        return;
      }

      let rawScore = 0;
      const reasons: string[] = [];

      if (industryMatched) {
        rawScore += 40;
        reasons.push(
          lang === 'jp'
            ? `✅ 希望分野一致 (${slotInd || candInd})`
            : `✅ Đúng ngành ${slotInd || candInd}`
        );
      } else {
        rawScore += 15;
      }

      // Employment Type bonus
      if (mentionsBaito && slotIsBaito) {
        rawScore += 5;
        reasons.push(lang === 'jp' ? '✅ バイト希望一致' : '✅ Đúng hình thức Baito');
      } else if (!slotIsBaito && mentionsTokuteiOrFullTime) {
        rawScore += 5;
      }

      // 3. Exact Multi-prefecture Area Matching (Tokyo != Kyoto!)
      const slotArea = String(sl.area || '').trim();
      let areaMatched = false;
      if (hasCandidateArea && slotArea) {
        const slotAreaPrefs = extractAllAreaTokens(slotArea);
        // Match concrete prefecture in slot against concrete prefectures desired by candidate
        const matchedPref = slotAreaPrefs.find((sp) =>
          candAreaPrefs.some((cp) => cp.shortLabel === sp.shortLabel)
        );
        if (matchedPref) {
          areaMatched = true;
          rawScore += 25;
          // Extra precision bonus if candidate wants Tokyo 23 wards and slot is inside 23 wards
          if (
            matchedPref.shortLabel === '東京都' &&
            /23区/.test(candArea) &&
            TOKYO_23_WARDS_RX.test(slotArea)
          ) {
            rawScore += 4;
          }
          reasons.push(
            lang === 'jp'
              ? `✅ エリア一致 (${matchedPref.shortLabel})`
              : `✅ Khớp ${matchedPref.shortLabel}`
          );
        } else if (candRegions.length > 0 && slotAreaPrefs.length > 0) {
          const matchedViaRegion = slotAreaPrefs.find((sp) =>
            candRegions.some((reg) => reg.prefectures.includes(sp.shortLabel))
          );
          if (matchedViaRegion) {
            areaMatched = true;
            rawScore += 22;
            reasons.push(
              lang === 'jp'
                ? `✅ エリア一致 (${matchedViaRegion.shortLabel})`
                : `✅ Khớp ${matchedViaRegion.shortLabel}`
            );
          }
        } else {
          // Only fallback to explicit ward/city names >= 2 chars when no prefecture tokens collide
          const fallbackTokens = candArea
            .split(/[・、,\s\/（）()]+/)
            .map((t) => t.replace(/[都道府県市区町村]/g, '').trim())
            .filter((t) => t.length >= 2 && !/^(?:エリア|area|toàn|nhật|bản)$/i.test(t));
          const hit = fallbackTokens.find(
            (tok) => tok !== '京都' && slotArea.includes(tok)
          );
          if (hit) {
            areaMatched = true;
            rawScore += 22;
            reasons.push(lang === 'jp' ? `✅ エリア一致 (${hit})` : `✅ Khớp ${hit}`);
          }
        }
      } else if (!hasCandidateArea) {
        rawScore += 10;
      }

      // 4. JLPT Requirement (honest check — no false ✅ when JLPT is 要確認 / empty!)
      const reqJlpt = String(sl.minJlpt || '要確認').trim();
      const reqJlptScore = jlptRank(reqJlpt);
      if (reqJlpt === '不問') {
        rawScore += 12;
        reasons.push(lang === 'jp' ? '✅ 日本語不問' : '✅ Không yêu cầu JLPT');
      } else if (reqJlpt === '要確認' || reqJlptScore < 0) {
        rawScore += 5;
      } else if (candJlptScore > 0 && candJlptScore >= reqJlptScore) {
        rawScore += 15;
        reasons.push(
          lang === 'jp' ? `✅ 条件達成 (${cand.jlpt})` : `✅ Đạt ${cand.jlpt}`
        );
      } else if (candJlptScore > 0 && candJlptScore < reqJlptScore) {
        rawScore -= 20;
      }

      // 5. Experience Requirement
      const expReq = String(sl.experienceRequirement || '');
      const hasCandExp = Boolean(candExp || candPassedExam);
      if (/未経験|không yêu cầu/i.test(expReq)) {
        rawScore += 10;
        reasons.push(lang === 'jp' ? '✅ 未経験OK' : '✅ Không yêu cầu kinh nghiệm');
      } else if (/経験者|có kinh nghiệm/i.test(expReq)) {
        if (hasCandExp) {
          rawScore += 10;
          reasons.push(lang === 'jp' ? '✅ 経験条件一致' : '✅ Đáp ứng kinh nghiệm');
        } else {
          rawScore -= 10;
        }
      } else if (hasCandExp) {
        rawScore += 5;
      }

      // 6. Salary & Housing Fit
      let salaryFitAdded = false;
      if (slotIsBaito) {
        const slotHourlyM = `${sl.hourlyWage || ''} ${sl.salary || ''}`
          .replace(/,/g, '')
          .match(/(\d{3,4})\s*円/);
        const expHourlyM = candSalaryExp
          .replace(/,/g, '')
          .match(/(\d{3,4})\s*(?:円|yên|yen)/i);
        if (slotHourlyM && expHourlyM) {
          if (Number(slotHourlyM[1]) >= Number(expHourlyM[1])) {
            rawScore += 10;
            reasons.push(lang === 'jp' ? '✅ 希望時給以上' : '✅ Đạt mức lương kỳ vọng');
            salaryFitAdded = true;
          }
        }
      } else if (candSalaryExp) {
        const wantsNet = /thực\s*nhận|手取り|sau\s*trừ/i.test(candSalaryExp);
        const offeredSource =
          wantsNet && sl.estimatedNetSalary ? sl.estimatedNetSalary : sl.salary || '';
        const numSalaryM = offeredSource.replace(/,/g, '').match(/(\d{5,6})/);
        const expSalaryM = candSalaryExp.match(/(\d{2})\s*(?:man|vạn|万)/i);
        if (numSalaryM && expSalaryM) {
          const offeredYen = Number(numSalaryM[1]);
          const desiredYen = Number(expSalaryM[1]) * 10000;
          if (offeredYen >= desiredYen) {
            rawScore += 10;
            reasons.push(lang === 'jp' ? '✅ 希望給与以上' : '✅ Đạt mức lương kỳ vọng');
            salaryFitAdded = true;
          }
        }
      }

      if (!salaryFitAdded) {
        if (
          needsHousing &&
          /自社寮|借上げ|借り上げ|寮・社宅あり|寮あり/.test(sl.housingSupport || '')
        ) {
          rawScore += 8;
          reasons.push(lang === 'jp' ? '✅ 寮・社宅あり' : '✅ Có hỗ trợ nhà ở / KTX');
        } else if (sl.bonus === '賞与あり') {
          rawScore += 5;
          reasons.push(lang === 'jp' ? '🎁 賞与あり' : '🎁 Có thưởng');
        } else if (`${sl.salary || ''} ${sl.hourlyWage || ''}`.trim()) {
          rawScore += 3;
        }
      }

      const clampedScore = Math.max(15, Math.min(100, rawScore));
      results.push({
        score: clampedScore,
        reasons: reasons.slice(0, 5),
        company: comp,
        slot: sl,
        industryMatched,
        areaMatched,
      });
    });
  });

  // Sort strictly by: areaMatched (when candidate specified area) -> score desc
  results.sort((a, b) => {
    if (hasCandidateArea && Boolean(a.areaMatched) !== Boolean(b.areaMatched)) {
      return a.areaMatched ? -1 : 1;
    }
    return b.score - a.score;
  });

  // Deduplicate jobs sharing the same company + location + salary (keep highest-scoring match)
  const seenJobSignatures = new Set<string>();
  const dedupedResults: KyujinMatchResult[] = [];
  for (const item of results) {
    const compKey = String(item.company.companyName || '').trim().toLowerCase();
    const areaKey = String(formatPublicArea(item.slot.area) || item.slot.area || '')
      .replace(/\s+/g, '')
      .toLowerCase();
    const salaryMeta = getKyujinSlotSalaryDisplayMeta(
      item.slot,
      item.slot.employmentType || item.company.employmentType || ''
    );
    const salaryKey = String(
      salaryMeta.displayValue || item.slot.salary || item.slot.hourlyWage || ''
    )
      .replace(/[,.\s]/g, '')
      .toLowerCase();
    const dedupeSig = `${compKey}::${areaKey}::${salaryKey}`;
    if (seenJobSignatures.has(dedupeSig)) continue;
    seenJobSignatures.add(dedupeSig);
    dedupedResults.push(item);
  }

  const hasExactIndustryAndAreaMatch =
    hasCandidateArea && dedupedResults.some((r) => Boolean(r.areaMatched));

  let noticeMessage = '';
  if (hasCandidateIndustry) {
    if (dedupedResults.length === 0) {
      noticeMessage =
        lang === 'jp'
          ? `希望分野「${candInd}」${
              hasCandidateArea ? `（${candArea}）` : ''
            }に該当する求人が現在ありません。`
          : `Chưa có đơn đúng ngành ${candInd}${
              hasCandidateArea ? ` tại ${candArea}` : ''
            } trong kho Tin tuyển dụng.`;
    } else if (hasCandidateArea && !hasExactIndustryAndAreaMatch) {
      noticeMessage =
        lang === 'jp'
          ? `${candArea}の「${candInd}」求人がないため、他エリアの同分野求人（${dedupedResults.length}件）を表示しています。`
          : `Chưa có đơn đúng ngành ${candInd} tại ${candArea} — hiển thị ${dedupedResults.length} đơn cùng ngành ở khu vực khác để tham khảo.`;
    }
  }

  return {
    matches: dedupedResults.slice(0, 8),
    hasCandidateIndustry,
    hasCandidateArea,
    hasExactIndustryAndAreaMatch,
    noticeMessage,
  };
};

export const findMatchingKyujinsForCandidate = (
  cand: Partial<Candidate> = {},
  kyujinsList: KyujinCompany[] = []
): KyujinMatchResult[] => evaluateCandidateKyujinMatches(cand, kyujinsList).matches;

export const buildAnonymousJobPostText = (
  comp: KyujinCompany,
  slot: KyujinSlot,
  lang: 'vi' | 'jp' = 'vi'
): string => {
  const empType = slot.employmentType || comp.employmentType || 'Tokutei Ginou (特定技能)';
  const sector = slot.sector || comp.sector || detectSectorFromIndustryOrText(slot.industry || comp.industry);
  const industry = slot.industry || comp.industry || '特定技能';
  const publicArea = formatPublicArea(slot.area) || slot.area || 'Nhật Bản';
  const isBaito = /baito|アルバイト|パート/i.test(empType);
  const cleanJobDesc = sanitizeJobDescriptionForCandidate(slot.jobDescription || '');
  const expText =
    slot.experienceRequirement ||
    (slot.careQualification && slot.careQualification !== '不問'
      ? lang === 'jp'
        ? `資格・経験優遇（${slot.careQualification}）`
        : `Ưu tiên có chứng chỉ/kinh nghiệm (${slot.careQualification})`
      : lang === 'jp'
      ? '未経験OK'
      : 'Không yêu cầu kinh nghiệm (未経験OK)');

  const shiftLineParts = [
    slot.shiftTime || '',
    slot.shiftType ? `(${slot.shiftType})` : '',
    slot.holidays ? (lang === 'jp' ? `休日：${slot.holidays}` : `Ngày nghỉ: ${slot.holidays}`) : '',
    slot.overtimeHours ? (lang === 'jp' ? `残業：${slot.overtimeHours}` : `Tăng ca: ${slot.overtimeHours}`) : '',
  ].filter(Boolean);

  if (lang === 'jp') {
    const compPay = isBaito
      ? [
          slot.hourlyWage ? `時給：${slot.hourlyWage}` : `給与：${slot.salary || '応相談'}`,
          slot.commuteAllowance ? `交通費：${slot.commuteAllowance}` : '',
          slot.shiftsPerWeek ? `シフト：${slot.shiftsPerWeek}` : '',
        ]
          .filter(Boolean)
          .join(' ／ ')
      : [
          `基本給・総支給：${slot.salary || '応相談'}`,
          slot.nightShiftAllowance ? `夜勤手当：${slot.nightShiftAllowance}` : '',
          slot.estimatedNetSalary ? `手取り目安：${slot.estimatedNetSalary}` : '',
          slot.annualSalary ? `想定年収：${slot.annualSalary}` : '',
          slot.bonus === '賞与あり'
            ? `賞与あり${slot.bonusDetail ? `（${slot.bonusDetail}）` : ''}`
            : slot.bonus === '賞与なし'
            ? `賞与なし${slot.bonusDetail ? `（${slot.bonusDetail}）` : ''}`
            : '',
        ]
          .filter(Boolean)
          .join(' ／ ');

    return [
      `【求人案内（匿名）】`,
      `・雇用形態：${empType}`,
      `・分野 / 職種：${sector} - ${industry}${slot.facilityType ? `（${slot.facilityType}）` : ''}`,
      `・勤務地：${publicArea}`,
      slot.recruitCount ? `・募集人数：${slot.recruitCount}` : '',
      `・仕事内容：${cleanJobDesc || '面接時に詳しくご説明いたします。'}`,
      `・経験要件：${expText}${slot.experienceDetail ? `（${slot.experienceDetail}）` : ''}`,
      `・日本語レベル：${slot.minJlpt || '応相談'}`,
      `・給与・待遇：${compPay}`,
      `・住宅・寮：${slot.housingSupport || '要確認'}`,
      shiftLineParts.length > 0 ? `・勤務時間・休日：${shiftLineParts.join(' / ')}` : '',
      slot.interviewProcess ? `・選考方法：${slot.interviewProcess}` : '',
      slot.specialBenefits ? `・特別手当・福利厚生：${slot.specialBenefits}` : '',
    ]
      .filter(Boolean)
      .join('\n');
  }

  const paySummaryVi = isBaito
    ? [
        slot.hourlyWage ? `Lương giờ: ${slot.hourlyWage}` : `Lương: ${slot.salary || 'Thỏa thuận'}`,
        slot.commuteAllowance ? `Hỗ trợ đi lại: ${slot.commuteAllowance}` : '',
        slot.shiftsPerWeek ? `Số ca/tuần: ${slot.shiftsPerWeek}` : '',
      ]
        .filter(Boolean)
        .join(' | ')
    : [
        `Lương cơ bản / Tổng nhận: ${slot.salary || 'Trao đổi khi phỏng vấn'}`,
        slot.nightShiftAllowance ? `Phụ cấp ca đêm: ${slot.nightShiftAllowance}` : '',
        slot.estimatedNetSalary ? `Thực nhận ước tính (手取り): ${slot.estimatedNetSalary}` : '',
        slot.annualSalary ? `Thu nhập năm (年収): ${slot.annualSalary}` : '',
        slot.bonus === '賞与あり'
          ? `Thưởng: Có${slot.bonusDetail ? ` (${slot.bonusDetail})` : ''}`
          : slot.bonus === '賞与なし'
          ? `Thưởng: Không${slot.bonusDetail ? ` (${slot.bonusDetail})` : ''}`
          : '',
      ]
        .filter(Boolean)
        .join(' | ');

  return [
    `📢 [ĐƠN TUYỂN DỤNG ẨN DANH]`,
    `• Hình thức: ${empType}`,
    `• Lĩnh vực - Ngành nghề: ${sector} - ${industry}${slot.facilityType ? ` (${slot.facilityType})` : ''}`,
    `• Khu vực làm việc: ${publicArea}`,
    slot.recruitCount ? `• Số lượng tuyển: ${slot.recruitCount}` : '',
    `• Nội dung công việc: ${cleanJobDesc || 'Trao đổi chi tiết khi Mendan'}`,
    `• Yêu cầu kinh nghiệm: ${expText}${slot.experienceDetail ? ` — ${slot.experienceDetail}` : ''}`,
    `• Trình độ tiếng Nhật: ${slot.minJlpt || 'Trao đổi thêm'}`,
    `• Chế độ lương thưởng: ${paySummaryVi}`,
    `• Nhà ở / KTX: ${slot.housingSupport || 'Xác nhận khi phỏng vấn'}`,
    shiftLineParts.length > 0 ? `• Thời gian & Ngày nghỉ: ${shiftLineParts.join(' | ')}` : '',
    slot.interviewProcess ? `• Quy trình phỏng vấn: ${slot.interviewProcess}` : '',
    slot.specialBenefits ? `• Phúc lợi đặc biệt: ${slot.specialBenefits}` : '',
  ]
    .filter(Boolean)
    .join('\n');
};

// ============================================================================
// EXCEL / PDF / DOCX / SECTOR / INDUSTRY EXTRACTION ENGINE
// ============================================================================
export const detectSectorFromIndustryOrText = (industry = '', text = ''): string => {
  const combined = `${industry} ${text}`;
  if (/介護|デイサービス|グループホーム|特養|老健|訪問介護|điều dưỡng|kaigo/i.test(combined)) {
    return 'Kaigo / Điều dưỡng';
  }
  if (/外食|飲食料品製造|レストラン|居酒屋|キッチン|弁当|食品|thực phẩm|nhà hàng|bếp/i.test(combined)) {
    return 'Thực phẩm & Dịch vụ ăn uống';
  }
  if (/宿泊|ホテル|旅館|khách sạn|du lịch|lễ tân/i.test(combined)) {
    return 'Khách sạn & Du lịch';
  }
  if (/自動車運送|トラック|ドライバー|配送|vận tải|tài xế|logistics/i.test(combined)) {
    return 'Vận tải & Logistics';
  }
  if (/建設|自動車整備|溶接|機械|xây dựng|cơ khí|bảo dưỡng/i.test(combined)) {
    return 'Xây dựng & Cơ khí';
  }
  if (/素形材|産業機械|電気電子|工場|製造|nhà máy|sản xuất|lắp ráp/i.test(combined)) {
    return 'Sản xuất & Nhà máy';
  }
  if (/ビルクリーニング|清掃|農業|vệ sinh|nông nghiệp/i.test(combined)) {
    return 'Vệ sinh tòa nhà & Nông nghiệp';
  }
  if (/コンビニ|スーパー|小売|アルバイト|パート|baito|làm thêm|combini/i.test(combined)) {
    return 'Bán lẻ / Dịch vụ Baito';
  }
  return 'Sản xuất & Nhà máy';
};

export const detectEmploymentTypeFromText = (text = '', hint = ''): string => {
  const combined = `${hint} ${text}`;
  if (/アルバイト|パート|baito|làm thêm|時給/i.test(combined) && !/特定技能/i.test(combined)) {
    return 'Baito / Làm thêm (アルバイト・パート)';
  }
  if (/正社員|登用|chuyển đổi visa|nhân viên chính thức/i.test(combined) && !/特定技能/i.test(combined)) {
    return 'Nhân viên chính thức / Chuyển đổi visa (正社員・登用)';
  }
  if (/アルバイト・パート|baito/i.test(combined)) {
    return 'Baito / Làm thêm (アルバイト・パート)';
  }
  return 'Tokutei Ginou (特定技能)';
};

export const detectIndustryFromText = (text = '', hint = ''): string => {
  const cleanHint = String(hint || '').trim();
  const cleanText = String(text || '').trim();

  const matchFromSource = (src: string): string => {
    if (!src) return '';
    if (/自動車運送|トラック|ドライバー|運行|配送|運転手|貨物/.test(src)) return '自動車運送業';
    if (/外食|飲食業|レストラン|ホール|キッチン|調理|居酒屋|カフェ/.test(src)) return '外食業';
    if (/飲食料品製造|食品加工|弁当|惣菜|製パン|水産加工/.test(src)) return '飲食料品製造業';
    if (/介護|デイサービス|グループホーム|老人ホーム|特養|老健|訪問介護|ヘルパー/.test(src))
      return '介護';
    if (/宿泊|ホテル|旅館|フロント|客室/.test(src)) return '宿泊';
    if (/ビルクリーニング|清掃/.test(src)) return 'ビルクリーニング';
    if (/建設|土木|足場|型枠|内装/.test(src)) return '建設';
    if (/自動車整備|メカニック/.test(src)) return '自動車整備';
    if (/農業|耕種|畜産/.test(src)) return '農業';
    if (/コンビニ|スーパー|小売|販売/.test(src)) return 'コンビニ・小売（アルバイト）';
    if (/製造|溶接|機械加工|組立|工場/.test(src))
      return '素形材・産業機械・電気電子情報関連製造業';
    return '';
  };

  const fromHint = matchFromSource(cleanHint);
  if (fromHint) return fromHint;

  const headerSegment = cleanText.slice(0, 260);
  const fromHeader = matchFromSource(headerSegment);
  if (fromHeader) return fromHeader;

  const fromBody = matchFromSource(cleanText);
  return fromBody || '特定技能';
};

export const extractExcelCheckedAnchors = async (
  zip: JSZip
): Promise<Record<string, ExcelAnchor[]>> => {
  const sheetAnchors: Record<string, ExcelAnchor[]> = {};
  if (!zip) return sheetAnchors;

  const ctrlCheckedMap: Record<string, boolean> = {};
  const ctrlFiles = Object.keys(zip.files).filter((f) =>
    /^xl\/ctrlProps\/ctrlProp\d+\.xml$/i.test(f)
  );
  for (const cf of ctrlFiles) {
    try {
      const xml = await zip.files[cf].async('string');
      if (/checked\s*=\s*["'](Checked|1|true)["']/i.test(xml)) {
        ctrlCheckedMap['/' + cf] = true;
        ctrlCheckedMap['../' + cf.replace(/^xl\//, '')] = true;
      }
    } catch {
      // ignore invalid xml
    }
  }

  const sheetFiles = Object.keys(zip.files).filter((f) =>
    /^xl\/worksheets\/sheet\d+\.xml$/i.test(f)
  );
  for (const sf of sheetFiles) {
    try {
      const sheetNameKey = sf.replace(/^xl\/worksheets\//, '').replace(/\.xml$/i, '');
      const relPath = sf.replace(
        /^xl\/worksheets\/(sheet\d+\.xml)$/i,
        'xl/worksheets/_rels/$1.rels'
      );
      if (!zip.files[relPath]) continue;

      const relXml = await zip.files[relPath].async('string');
      const idToTarget: Record<string, string> = {};
      const relRegex = /<Relationship\b[^>]*>/gi;
      let rm: RegExpExecArray | null;
      while ((rm = relRegex.exec(relXml)) !== null) {
        const tag = rm[0];
        const idMatch = tag.match(/\bId\s*=\s*["']([^"']+)["']/i);
        const targetMatch = tag.match(/\bTarget\s*=\s*["']([^"']+)["']/i);
        if (idMatch && targetMatch) {
          idToTarget[idMatch[1]] = targetMatch[1];
        }
      }

      const sheetXml = await zip.files[sf].async('string');
      const checkedShapeIds = new Set<string>();
      const ctrlTagRegex = /<control\b[^>]*>/gi;
      let cm: RegExpExecArray | null;
      while ((cm = ctrlTagRegex.exec(sheetXml)) !== null) {
        const cTag = cm[0];
        const shapeMatch = cTag.match(/\bshapeId\s*=\s*["']([^"']+)["']/i);
        const rIdMatch = cTag.match(/\br:id\s*=\s*["']([^"']+)["']/i);
        if (shapeMatch && rIdMatch) {
          const target = idToTarget[rIdMatch[1]] || '';
          if (ctrlCheckedMap[target] || ctrlCheckedMap['/' + target.replace(/^\.\.\//, 'xl/')]) {
            checkedShapeIds.add(shapeMatch[1]);
          }
        }
      }

      const vmlTargets = Object.values(idToTarget).filter((t) => /vmlDrawing/i.test(t));
      const anchors: ExcelAnchor[] = [];
      for (const vt of vmlTargets) {
        const vmlZipPath = vt.startsWith('/') ? vt.slice(1) : vt.replace(/^\.\.\//, 'xl/');
        if (!zip.files[vmlZipPath]) continue;
        const vmlXml = await zip.files[vmlZipPath].async('string');
        const shapeBlockRegex = /<v:shape\b([^>]*)>([\s\S]*?)<\/v:shape>/gi;
        let sm: RegExpExecArray | null;
        while ((sm = shapeBlockRegex.exec(vmlXml)) !== null) {
          const attrs = sm[1] || '';
          const body = sm[2] || '';
          const idAttr = attrs.match(/\bid\s*=\s*["']([^"']+)["']/i);
          const spidAttr = attrs.match(/\bo:spid\s*=\s*["']([^"']+)["']/i);
          const numId =
            (idAttr && idAttr[1].match(/(\d+)$/)?.[1]) ||
            (spidAttr && spidAttr[1].match(/(\d+)$/)?.[1]) ||
            '';
          const isExplicitChecked =
            /<x:Checked>\s*(?:1|true|Checked)?\s*<\/x:Checked>/i.test(body) ||
            /<x:Checked\s*\/>/i.test(body);
          if ((numId && checkedShapeIds.has(numId)) || isExplicitChecked) {
            const anchorMatch = body.match(/<x:Anchor>([\s\S]*?)<\/x:Anchor>/i);
            if (anchorMatch) {
              const nums = anchorMatch[1]
                .split(',')
                .map((n) => parseInt(n.trim(), 10))
                .filter((n) => !Number.isNaN(n));
              if (nums.length >= 8) {
                anchors.push({
                  c1: nums[0],
                  r1: nums[2],
                  c2: nums[4],
                  r2: nums[6],
                });
              }
            }
          }
        }
      }
      sheetAnchors[sheetNameKey] = anchors;
    } catch {
      // ignore sheet error
    }
  }
  return sheetAnchors;
};

const extractSingleKaigoFacilityTypeForSlot = (slot: Partial<KyujinSlot> = {}, industry = '') => {
  const combined = `${slot.jobCode || ''} ${slot.area || ''} ${slot.jobDescription || ''}`;
  if (
    industry !== '介護' &&
    !/介護|デイサービス|グループホーム|有料老人ホーム|特別養護老人ホーム|老健|訪問介護/.test(
      combined
    )
  ) {
    return '';
  }
  if (/デイサービス|#CP|CP[^\s]/i.test(combined)) return 'デイサービス';
  if (/グループホーム|#CG|CG[^\s]/i.test(combined)) return 'グループホーム';
  if (/有料老人ホーム/.test(combined)) return '有料老人ホーム';
  if (/特別養護老人ホーム|特養/.test(combined)) return '特別養護老人ホーム';
  if (/介護老人保健施設|老健/.test(combined)) return '老健';
  if (/訪問介護|#ソエルガーデン/.test(combined)) return '訪問介護';
  return '';
};

export const normalizeBonusInfo = (bonus = '', bonusDetail = '', fallbackText = '') => {
  const b = String(bonus || '').trim();
  const d = String(bonusDetail || '').trim();
  const fb = String(fallbackText || '').trim();
  const combined = `${b} ${d} ${fb}`.trim();

  const hasExplicitFrequency = /年\s*\d+\s*回/.test(`${b} ${d}`);
  const condMatch = combined.match(
    /([^\n。、（）()]{0,20}(?:店長|正社員|登用|業績)[^\n。、（）()]{0,25}賞与[^\n。、（）()]{0,12})/
  );

  // Check ticked 無 / 賞与なし / 退職金・賞与無 BEFORE matching 賞与あり
  if (
    b === '賞与なし' ||
    (!hasExplicitFrequency &&
      (/[☑✓✔■●レ\uF0FE\uF052\u00FE\u00FC]\s*(?:賞与\s*)?無(?:\s|$|・|\b)|賞与\s*[:：]?\s*(?:無|なし)|賞与なし|退職金[・･]?\s*賞与\s*無/.test(
        combined
      ) ||
        (Boolean(condMatch) && !/賞与あり\s*（\s*年\s*\d+\s*回/.test(combined))))
  ) {
    return {
      bonus: '賞与なし',
      bonusDetail: d || (condMatch ? condMatch[1].trim() : ''),
    };
  }
  if (b === '賞与あり' || /賞与あり|ボーナスあり|決算賞与/.test(combined)) {
    const detailMatch = `${d} ${fb}`.match(
      /年\s*\d+\s*回[^\s。、）)]*|\d+(?:,\d+)*\s*万?円[^\s。、）)]*/
    );
    return {
      bonus: '賞与あり',
      bonusDetail: d || (detailMatch ? detailMatch[0] : ''),
    };
  }
  return { bonus: '要確認', bonusDetail: d || '' };
};

export const normalizeHousingDisplay = (val = ''): string => {
  const s = String(val || '').trim();
  if (!s || s === '⚠️ 要確認' || s === '要確認') return '⚠️ 要確認';
  if (/自社寮/.test(s)) return '🏠 自社寮あり';
  if (/借上げ|借り上げ|宿舎/.test(s)) return '🏠 借上げ宿舎あり';
  if (/寮・社宅あり|寮あり|社宅あり/.test(s)) return '🏠 寮・社宅あり';
  if (/提供無|なし|無/.test(s)) return '❌ 提供無';
  return s;
};

export const sanitizeJobDescriptionForCandidate = (desc = ''): string => {
  const cleaned = stripKyujinFormBoilerplate(sanitizeExcelFloatText(String(desc || '')));
  return cleaned
    .replace(/(?:家賃控除額は実費を上回る[^\n。]*[。\n]?)/g, '')
    .replace(/(?:写真掲載許可(?:をもらう|取得)[^\n。]*[。\n]?)/g, '')
    .replace(/(?:[※＊]\s*(?:社内用|内部メモ|担当メモ|紹介手数料)[^\n。]*[。\n]?)/g, '')
    .replace(/(?:[（(](?:社内メモ|担当用|内部用)[^)）]*[)）])/g, '')
    .replace(
      /(?:受動喫煙|試用期間|選考方法|応募書類|担当者|連絡先|電話番号|FAX|メールアドレス|所在地|法人番号)[\s\S]*$/i,
      ''
    )
    .replace(/\s+/g, ' ')
    .trim();
};

const normalizeKyujinExperienceRequirement = (rawReq = '', contextText = ''): string => {
  const req = String(rawReq || '').trim();
  const combined = `${req} ${contextText}`.trim();
  if (/経験者必須|経験あり必須|要経験/.test(combined)) {
    return 'Yêu cầu có kinh nghiệm (経験者必須)';
  }
  if (
    /経験あれば尚可|経験者優遇|優遇|尚可/.test(combined) ||
    /未経験可・経験者優遇/.test(req)
  ) {
    return 'Ưu tiên có kinh nghiệm / Không yêu cầu bắt buộc (未経験可・経験者優遇)';
  }
  if (req) return req;
  if (/未経験OK|未経験歓迎|未経験可|不問/i.test(combined)) {
    return 'Không yêu cầu kinh nghiệm (未経験OK)';
  }
  return '';
};

export const normalizeKyujinItem = (
  compInput: (Partial<KyujinCompany> & Partial<KyujinSlot>) | unknown = {}
): KyujinCompany => {
  const rawCompObj =
    compInput && typeof compInput === 'object' && !Array.isArray(compInput)
      ? (compInput as Record<string, unknown>)
      : {};
  const comp = rawCompObj as Partial<KyujinCompany> & Partial<KyujinSlot> & Record<string, unknown>;

  const compId = cleanLegacyStr(comp.id) || `kj_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`;
  const compCompanyName = cleanLegacyStr(comp.companyName ?? comp.company_name ?? comp.name);
  const compSector = cleanLegacyStr(comp.sector);
  const compIndustry = cleanLegacyStr(comp.industry);
  const compEmploymentType = cleanLegacyStr(comp.employmentType ?? comp.employment_type);
  const compJobCode = cleanLegacyStr(comp.jobCode ?? comp.job_code);
  const compArea = cleanLegacyStr(comp.area);
  const compShiftTime = sanitizeExcelFloatText(
    stripKyujinFormBoilerplate(cleanLegacyStr(comp.shiftTime ?? comp.shift_time))
  );
  const compShiftType = cleanLegacyStr(comp.shiftType ?? comp.shift_type);
  const compHolidays = sanitizeExcelFloatText(
    stripKyujinFormBoilerplate(cleanLegacyStr(comp.holidays))
  );
  const compOvertimeHours = sanitizeExcelFloatText(
    stripKyujinFormBoilerplate(cleanLegacyStr(comp.overtimeHours ?? comp.overtime_hours))
  );
  const compSalary = sanitizeExcelFloatText(
    stripKyujinFormBoilerplate(cleanLegacyStr(comp.salary))
  );
  const compNightShiftAllowance = sanitizeExcelFloatText(
    cleanLegacyStr(comp.nightShiftAllowance ?? comp.night_shift_allowance)
  );
  const compEstimatedNetSalary = sanitizeExcelFloatText(
    stripKyujinFormBoilerplate(
      cleanLegacyStr(comp.estimatedNetSalary ?? comp.estimated_net_salary)
    )
  );
  const compAnnualSalary = sanitizeExcelFloatText(
    stripKyujinFormBoilerplate(cleanLegacyStr(comp.annualSalary ?? comp.annual_salary))
  );
  const compHourlyWage = sanitizeExcelFloatText(
    stripKyujinFormBoilerplate(cleanLegacyStr(comp.hourlyWage ?? comp.hourly_wage))
  );
  const compCommuteAllowance = cleanLegacyStr(
    comp.commuteAllowance ?? comp.commute_allowance
  );
  const compShiftsPerWeek = cleanLegacyStr(comp.shiftsPerWeek ?? comp.shifts_per_week);
  const compBonus = cleanLegacyStr(comp.bonus);
  const compBonusDetail = sanitizeExcelFloatText(
    stripKyujinFormBoilerplate(cleanLegacyStr(comp.bonusDetail ?? comp.bonus_detail))
  );
  const compMinJlpt = cleanLegacyStr(comp.minJlpt ?? comp.min_jlpt);
  const compHousingSupport = cleanLegacyStr(comp.housingSupport ?? comp.housing_support);
  const compRecruitmentStatus = cleanLegacyStr(
    comp.recruitmentStatus ?? comp.recruitment_status
  );
  const compRecruitCount = cleanLegacyStr(comp.recruitCount ?? comp.recruit_count);
  const compInterviewProcess = cleanLegacyStr(
    comp.interviewProcess ?? comp.interview_process
  );
  const compSpecialBenefits = sanitizeExcelFloatText(
    stripKyujinFormBoilerplate(cleanLegacyStr(comp.specialBenefits ?? comp.special_benefits))
  );
  const compCareQualification = cleanLegacyStr(
    comp.careQualification ?? comp.care_qualification
  );
  const compExperienceRequirement = cleanLegacyStr(
    comp.experienceRequirement ?? comp.experience_requirement
  );
  const compExperienceDetail = cleanLegacyStr(
    comp.experienceDetail ?? comp.experience_detail
  );
  const compJobDescription = cleanLegacyStr(
    comp.jobDescription ?? comp.job_description ?? comp.description
  );
  const compExcludedNats = toSafeStringArray(
    comp.excludedNationalities ?? comp.excluded_nationalities
  );
  const compExcludedGenders = toSafeStringArray(
    comp.excludedGenders ?? comp.excluded_genders
  );
  const compRequiredConditions = toSafeStringArray(
    comp.requiredConditions ?? comp.required_conditions
  );
  const compNgConditions = toSafeStringArray(comp.ngConditions ?? comp.ng_conditions);
  const compRequiresInJapan = Boolean(comp.requiresInJapan ?? comp.requires_in_japan);
  const compRequiresTokutei = Boolean(comp.requiresTokutei ?? comp.requires_tokutei);

  const candidateSlotArray = Array.isArray(comp.areaBreakdown)
    ? comp.areaBreakdown
    : Array.isArray(comp.area_breakdown)
    ? comp.area_breakdown
    : Array.isArray(comp.slots)
    ? comp.slots
    : Array.isArray(comp.jobs)
    ? comp.jobs
    : [];

  const validSlots = candidateSlotArray.filter(
    (s): s is Partial<KyujinSlot> & Record<string, unknown> =>
      Boolean(s) && typeof s === 'object' && !Array.isArray(s)
  );

  const rawSlots: Array<Partial<KyujinSlot> & Record<string, unknown>> =
    validSlots.length > 0
      ? validSlots
      : [
          {
            slotId: `${compId}_slot_1`,
            sector: compSector,
            industry: compIndustry,
            employmentType: compEmploymentType,
            jobCode: compJobCode,
            area: compArea,
            shiftTime: compShiftTime,
            shiftType: compShiftType,
            holidays: compHolidays,
            overtimeHours: compOvertimeHours,
            salary: compSalary,
            nightShiftAllowance: compNightShiftAllowance,
            estimatedNetSalary: compEstimatedNetSalary,
            annualSalary: compAnnualSalary,
            hourlyWage: compHourlyWage,
            commuteAllowance: compCommuteAllowance,
            shiftsPerWeek: compShiftsPerWeek,
            bonus: compBonus,
            bonusDetail: compBonusDetail,
            minJlpt: compMinJlpt,
            housingSupport: compHousingSupport,
            recruitmentStatus: compRecruitmentStatus || '募集中',
            recruitCount: compRecruitCount,
            interviewProcess: compInterviewProcess,
            specialBenefits: compSpecialBenefits,
            careQualification: compCareQualification || '',
            experienceRequirement: compExperienceRequirement,
            experienceDetail: compExperienceDetail,
            jobDescription: compJobDescription || '',
          },
        ];

  const normalizedSlots: KyujinSlot[] = rawSlots.map((slObj, idx) => {
    const sl = slObj as Partial<KyujinSlot> & Record<string, unknown>;
    const slJobDescRaw = cleanLegacyStr(sl.jobDescription ?? sl.job_description ?? sl.description);
    const cleanDesc = sanitizeJobDescriptionForCandidate(slJobDescRaw || compJobDescription || '');
    const slIndustry = cleanLegacyStr(sl.industry) || compIndustry || '特定技能';
    const slSector =
      cleanLegacyStr(sl.sector) ||
      compSector ||
      detectSectorFromIndustryOrText(slIndustry, cleanDesc);
    const slSalaryRaw = sanitizeExcelFloatText(
      stripKyujinFormBoilerplate(cleanLegacyStr(sl.salary))
    );
    const slHourlyRaw = sanitizeExcelFloatText(
      stripKyujinFormBoilerplate(cleanLegacyStr(sl.hourlyWage ?? sl.hourly_wage))
    );
    const slEmploymentType =
      cleanLegacyStr(sl.employmentType ?? sl.employment_type) ||
      compEmploymentType ||
      detectEmploymentTypeFromText(
        `${slSalaryRaw || compSalary} ${slHourlyRaw || compHourlyWage} ${cleanDesc}`
      );
    const facilityType =
      cleanLegacyStr(sl.facilityType ?? sl.facility_type) ||
      extractSingleKaigoFacilityTypeForSlot(
        { ...sl, jobDescription: cleanDesc },
        slIndustry
      );
    const bonusNorm = normalizeBonusInfo(
      cleanLegacyStr(sl.bonus) || compBonus,
      sanitizeExcelFloatText(
        stripKyujinFormBoilerplate(
          cleanLegacyStr(sl.bonusDetail ?? sl.bonus_detail) || compBonusDetail
        )
      ),
      `${slJobDescRaw} ${compJobDescription}`
    );
    const slCareQual =
      cleanLegacyStr(sl.careQualification ?? sl.care_qualification) || compCareQualification;
    const expReq = normalizeKyujinExperienceRequirement(
      cleanLegacyStr(sl.experienceRequirement ?? sl.experience_requirement) ||
        compExperienceRequirement,
      `${cleanDesc} ${slCareQual}`
    );

    const slotMinJlpt =
      cleanLegacyStr(sl.minJlpt ?? sl.min_jlpt) || compMinJlpt || '要確認';
    const slExpDetail =
      cleanLegacyStr(sl.experienceDetail ?? sl.experience_detail) || compExperienceDetail;
    const conds = extractNormalizedKyujinConditions({
      text: `${cleanDesc} ${slExpDetail} ${slCareQual}`,
      minJlpt: slotMinJlpt,
      employmentType: slEmploymentType,
      existingExcludedNationalities: [
        ...compExcludedNats,
        ...toSafeStringArray(sl.excludedNationalities ?? sl.excluded_nationalities),
      ],
      existingExcludedGenders: [
        ...compExcludedGenders,
        ...toSafeStringArray(sl.excludedGenders ?? sl.excluded_genders),
      ],
      existingRequiresInJapan: Boolean(
        sl.requiresInJapan ?? sl.requires_in_japan ?? compRequiresInJapan
      ),
      existingRequiresTokutei: Boolean(
        sl.requiresTokutei ?? sl.requires_tokutei ?? compRequiresTokutei
      ),
      existingRequiredConditions: [
        ...compRequiredConditions,
        ...toSafeStringArray(sl.requiredConditions ?? sl.required_conditions),
      ],
      existingNgConditions: [
        ...compNgConditions,
        ...toSafeStringArray(sl.ngConditions ?? sl.ng_conditions),
      ],
    });

    const slRecStatus =
      cleanLegacyStr(sl.recruitmentStatus ?? sl.recruitment_status) === '募集終了' ||
      compRecruitmentStatus === '募集終了'
        ? '募集終了'
        : '募集中';

    const effectiveHourly = slHourlyRaw || compHourlyWage;
    const effectiveSalary =
      slSalaryRaw ||
      compSalary ||
      (effectiveHourly ? `時給 ${effectiveHourly}` : '給与要確認');

    const slShiftTime =
      sanitizeExcelFloatText(
        stripKyujinFormBoilerplate(cleanLegacyStr(sl.shiftTime ?? sl.shift_time))
      ) || compShiftTime;
    const slShiftType =
      cleanLegacyStr(sl.shiftType ?? sl.shift_type) ||
      compShiftType ||
      detectKyujinShiftType(slShiftTime, `${cleanDesc} ${slJobDescRaw}`);
    const slHolidays =
      sanitizeExcelFloatText(
        stripKyujinFormBoilerplate(cleanLegacyStr(sl.holidays))
      ) || compHolidays;
    const slOvertimeHours =
      sanitizeExcelFloatText(
        stripKyujinFormBoilerplate(cleanLegacyStr(sl.overtimeHours ?? sl.overtime_hours))
      ) || compOvertimeHours;
    const slEstNet =
      sanitizeExcelFloatText(
        stripKyujinFormBoilerplate(
          cleanLegacyStr(sl.estimatedNetSalary ?? sl.estimated_net_salary)
        )
      ) || compEstimatedNetSalary;
    const slAnnualSal =
      sanitizeExcelFloatText(
        stripKyujinFormBoilerplate(cleanLegacyStr(sl.annualSalary ?? sl.annual_salary))
      ) || compAnnualSalary;

    return {
      ...sl,
      slotId: cleanLegacyStr(sl.slotId ?? sl.slot_id) || `${compId}_slot_${idx + 1}`,
      sector: slSector,
      industry: slIndustry,
      employmentType: slEmploymentType,
      jobCode: cleanLegacyStr(sl.jobCode ?? sl.job_code) || compJobCode,
      area: cleanLegacyStr(sl.area) || compArea,
      shiftTime: slShiftTime,
      shiftType: slShiftType,
      holidays: slHolidays,
      overtimeHours: slOvertimeHours,
      salary: effectiveSalary,
      nightShiftAllowance:
        sanitizeExcelFloatText(
          cleanLegacyStr(sl.nightShiftAllowance ?? sl.night_shift_allowance)
        ) || compNightShiftAllowance,
      estimatedNetSalary: slEstNet,
      annualSalary: slAnnualSal,
      hourlyWage: effectiveHourly,
      commuteAllowance:
        cleanLegacyStr(sl.commuteAllowance ?? sl.commute_allowance) || compCommuteAllowance,
      shiftsPerWeek:
        cleanLegacyStr(sl.shiftsPerWeek ?? sl.shifts_per_week) || compShiftsPerWeek,
      bonus: bonusNorm.bonus,
      bonusDetail: bonusNorm.bonusDetail,
      minJlpt: slotMinJlpt,
      housingSupport: normalizeHousingDisplay(
        cleanLegacyStr(sl.housingSupport ?? sl.housing_support) || compHousingSupport
      ),
      recruitmentStatus: slRecStatus,
      recruitCount:
        cleanLegacyStr(sl.recruitCount ?? sl.recruit_count) || compRecruitCount,
      interviewProcess:
        cleanLegacyStr(sl.interviewProcess ?? sl.interview_process) || compInterviewProcess,
      specialBenefits:
        sanitizeExcelFloatText(
          stripKyujinFormBoilerplate(cleanLegacyStr(sl.specialBenefits ?? sl.special_benefits))
        ) || compSpecialBenefits,
      careQualification: slCareQual,
      facilityType,
      experienceRequirement: expReq,
      experienceDetail: slExpDetail,
      jobDescription: cleanDesc,
      excludedNationalities: conds.excludedNationalities,
      excludedGenders: conds.excludedGenders,
      requiresInJapan: conds.requiresInJapan,
      requiresTokutei: conds.requiresTokutei,
      requiredConditions: conds.requiredConditions,
      ngConditions: conds.ngConditions,
      ...(sl.salaryManuallyEdited ? { salaryManuallyEdited: true } : {}),
      ...(Array.isArray(sl.manualProtectedFields) && sl.manualProtectedFields.length > 0
        ? { manualProtectedFields: toSafeStringArray(sl.manualProtectedFields) }
        : {}),
    };
  });

  const hasAnyOpen = normalizedSlots.some((s) => s.recruitmentStatus !== '募集終了');
  const primaryIndustry =
    normalizedSlots[0]?.industry || compIndustry || '特定技能';
  const primarySector =
    normalizedSlots[0]?.sector ||
    compSector ||
    detectSectorFromIndustryOrText(primaryIndustry);
  const primaryEmpType =
    normalizedSlots[0]?.employmentType ||
    compEmploymentType ||
    'Tokutei Ginou (特定技能)';
  const mergedExcludedNationalities = Array.from(
    new Set([
      ...compExcludedNats.map((n) => normalizeCanonicalNationality(n)).filter(Boolean),
      ...normalizedSlots.flatMap((s) => toSafeStringArray(s.excludedNationalities)),
    ])
  );
  const mergedExcludedGenders = Array.from(
    new Set([
      ...compExcludedGenders.map((g) => normalizeCanonicalGender(g)).filter(Boolean),
      ...normalizedSlots.flatMap((s) => toSafeStringArray(s.excludedGenders)),
    ])
  );
  const mergedRequiredConditions = Array.from(
    new Set([
      ...compRequiredConditions,
      ...normalizedSlots.flatMap((s) => toSafeStringArray(s.requiredConditions)),
    ])
  );
  const mergedNgConditions = Array.from(
    new Set([
      ...compNgConditions,
      ...normalizedSlots.flatMap((s) => toSafeStringArray(s.ngConditions)),
    ])
  );

  return {
    id: compId,
    companyName: (compCompanyName || '名称未設定企業')
      .replace(/\s*代表者.*$/, '')
      .trim(),
    sector: primarySector,
    industry: primaryIndustry,
    employmentType: primaryEmpType,
    recruitmentStatus: hasAnyOpen ? '募集中' : '募集終了',
    excludedNationalities: mergedExcludedNationalities,
    excludedGenders: mergedExcludedGenders,
    requiresInJapan: Boolean(
      compRequiresInJapan || normalizedSlots.some((s) => s.requiresInJapan)
    ),
    requiresTokutei: Boolean(
      compRequiresTokutei || normalizedSlots.some((s) => s.requiresTokutei)
    ),
    requiredConditions: mergedRequiredConditions,
    ngConditions: mergedNgConditions,
    areaBreakdown: normalizedSlots,
  };
};

const extractJlptLevel = (
  rows: unknown[][] = [],
  anchors: ExcelAnchor[] = [],
  rowOffset = 0
): string => {
  if (!Array.isArray(rows) || rows.length === 0) return '要確認';

  const labelColMap: Record<string, number> = {};
  let jpRowIdx = -1;

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    const rowStr = row.map((c) => String(c || '')).join(' ');
    if (/日本語レベル|日本語能力|求める日本語/.test(rowStr)) {
      jpRowIdx = r;
      row.forEach((cell, cIdx) => {
        const s = String(cell || '').trim();
        if (/^N4$/i.test(s)) labelColMap.N4 = cIdx;
        else if (/^N3$/i.test(s)) labelColMap.N3 = cIdx;
        else if (/^N2$/i.test(s)) labelColMap.N2 = cIdx;
        else if (/^N1$/i.test(s)) labelColMap.N1 = cIdx;
        else if (/^N5$/i.test(s)) labelColMap.N5 = cIdx;
        else if (/^不問$/.test(s)) labelColMap['不問'] = cIdx;
      });
      break;
    }
  }

  const absJpRow = jpRowIdx >= 0 ? jpRowIdx + rowOffset : -1;
  if (absJpRow >= 0 && Array.isArray(anchors) && anchors.length > 0) {
    const matchedLevels: string[] = [];
    anchors.forEach((anc) => {
      if (anc.r1 <= absJpRow && absJpRow <= anc.r2) {
        const targetCol = Math.max(anc.c2, Math.round((anc.c1 + anc.c2) / 2));
        if (anc.c1 === 5 && anc.c2 === 5) matchedLevels.push('N4');
        else if (targetCol === 6 || targetCol === labelColMap.N3) matchedLevels.push('N3');
        else if (targetCol === 7 || targetCol === labelColMap.N2) matchedLevels.push('N2');
        else if (targetCol === 8 || targetCol === labelColMap.N1) matchedLevels.push('N1');
        else if (targetCol >= 11 || (labelColMap['不問'] && targetCol >= labelColMap['不問'] - 1)) {
          matchedLevels.push('不問');
        }
      }
    });
    if (matchedLevels.length > 0) {
      const priority = ['N5', 'N4', 'N3', 'N2', 'N1', '不問'];
      matchedLevels.sort((a, b) => priority.indexOf(a) - priority.indexOf(b));
      return matchedLevels[0];
    }
  }

  if (jpRowIdx >= 0) {
    const row = rows[jpRowIdx] || [];
    for (let c = 0; c < row.length; c++) {
      const cell = String(row[c] || '').trim();
      const prevCell = c > 0 ? String(row[c - 1] || '').trim() : '';
      if (/[☑✓✔■●〇○レ\uF0FE\uF052\u00FE\u00FC]\s*(N[1-5]|日常会話|不問)/i.test(cell)) {
        const m = cell.match(/(N[1-5]|日常会話|不問)/i);
        if (m) return m[1].toUpperCase();
      }
      if (/^[①②③④⑤]$/.test(cell)) {
        const nextCell = c + 1 < row.length ? String(row[c + 1] || '').trim() : '';
        if (/^(N[1-5]|日常会話|不問)$/i.test(nextCell)) return nextCell.toUpperCase();
        if (cell === '④') return 'N4';
        if (cell === '③') return 'N3';
        if (cell === '②') return 'N2';
        if (cell === '①') return 'N1';
      }
      if (
        /^(N[1-5]|日常会話|不問)$/i.test(cell) &&
        /^[☑✓✔■●〇○レ\uF0FE\uF052\u00FE\u00FC]$/.test(prevCell)
      ) {
        return cell.toUpperCase();
      }
    }
  }

  const fullText = rows.map((r) => (r || []).join(' ')).join(' ');
  const circledMatch = fullText.match(/(?:日本語(?:能力|レベル)?|JLPT)[^\n]{0,30}[④③②①◯○◎●]\s*(N[1-5])/i);
  if (circledMatch) return circledMatch[1].toUpperCase();
  const condMatch = fullText.match(/(?:日本語(?:能力|レベル)?|JLPT)\s*[：:・\s]*[^\n]{0,20}\b(N[1-5])\s*以上/i);
  if (condMatch) return condMatch[1].toUpperCase();

  return '要確認';
};

const extractHousingFromRows = (
  rows: unknown[][] = [],
  anchors: ExcelAnchor[] = [],
  rowOffset = 0
): string => {
  let hRowIdx = -1;
  for (let r = 0; r < rows.length; r++) {
    const rowStr = (rows[r] || []).join(' ');
    if (/提供する住宅設備|住宅設備|寮・社宅|社宅/.test(rowStr)) {
      hRowIdx = r;
      break;
    }
  }
  if (hRowIdx >= 0) {
    const checkRows = [hRowIdx, hRowIdx + 1].filter((idx) => idx < rows.length);
    for (const rIdx of checkRows) {
      const absRow = rIdx + rowOffset;
      for (const anc of anchors || []) {
        if (anc.r1 <= absRow && absRow <= anc.r2) {
          const targetCol = Math.max(anc.c2, Math.round((anc.c1 + anc.c2) / 2));
          if (targetCol >= 10) return '❌ 提供無';
          if (targetCol >= 7 && targetCol <= 9) return '🏠 借上げ宿舎あり';
          if (targetCol <= 6) return '🏠 自社寮あり';
        }
      }
      const row = rows[rIdx] || [];
      for (let c = 0; c < row.length; c++) {
        const cell = String(row[c] || '').trim();
        const prev = c > 0 ? String(row[c - 1] || '').trim() : '';
        if (/[☑✓✔■●レ\uF0FE\uF052\u00FE\u00FC]\s*(自社寮|借上げ宿舎|提供無)/.test(cell)) {
          if (/自社寮/.test(cell)) return '🏠 自社寮あり';
          if (/借上げ/.test(cell)) return '🏠 借上げ宿舎あり';
          if (/提供無/.test(cell)) return '❌ 提供無';
        }
        if (/^[☑✓✔■●レ\uF0FE\uF052\u00FE\u00FC]$/.test(prev)) {
          if (/自社寮/.test(cell)) return '🏠 自社寮あり';
          if (/借上げ/.test(cell)) return '🏠 借上げ宿舎あり';
          if (/提供無/.test(cell)) return '❌ 提供無';
        }
      }
    }
  }
  return '⚠️ 要確認';
};

const extractSalaryFromRows = (rows: unknown[][] = []): string =>
  extractKyujinSalaryFromRows(rows).salary;

const splitSheetRowsIntoKyujinBlocks = (rows: unknown[][] = []) => {
  const jobCodeRowIndices: number[] = [];
  rows.forEach((r, idx) => {
    const joined = (r || []).join(' ');
    if (/案件番号/.test(joined)) {
      jobCodeRowIndices.push(idx);
    }
  });
  const validStarts: number[] = [];
  jobCodeRowIndices.forEach((idx) => {
    if (validStarts.length === 0 || idx - validStarts[validStarts.length - 1] >= 15) {
      validStarts.push(idx);
    }
  });
  if (validStarts.length <= 1) {
    return [{ startRow: 0, rows }];
  }
  const blocks: Array<{ startRow: number; rows: unknown[][] }> = [];
  for (let i = 0; i < validStarts.length; i++) {
    const s = Math.max(0, validStarts[i] - 5);
    const e =
      i + 1 < validStarts.length ? Math.max(s + 1, validStarts[i + 1] - 5) : rows.length;
    blocks.push({ startRow: s, rows: rows.slice(s, e) });
  }
  return blocks;
};

const extractDetailedKyujinSpecs = (
  rawText: string,
  rows: unknown[][] = []
) => {
  const cleanedText = sanitizeExcelFloatText(stripKyujinFormBoilerplate(rawText));
  let shiftTime = '';

  // 1) Inspect row-based table first for multi-line numbered shifts under 就業時間 / 勤務時間
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    const rowStr = row.map((v) => String(v || '').trim()).filter(Boolean).join(' ');
    if (/^(?:就業時間|勤務時間|シフト時間)/.test(rowStr) || /\b(?:就業時間|勤務時間)\b/.test(rowStr)) {
      const collected: string[] = [];
      const sameRowVal = rowStr
        .replace(/^.*?(?:就業時間|勤務時間|シフト時間)[：:\s]*/, '')
        .trim();
      if (sameRowVal && /\d{1,2}[:：]\d{2}|[①②③④]/.test(sameRowVal)) {
        collected.push(sameRowVal);
      }
      for (let nr = r + 1; nr < Math.min(rows.length, r + 6); nr++) {
        const nextStr = (rows[nr] || [])
          .map((v) => String(v || '').trim())
          .filter(Boolean)
          .join(' ');
        if (!nextStr) continue;
        if (/^(?:休憩時間|時間外|残業|休日|年間休日|給与|月給|時給|基本給|賞与|福利厚生|勤務地)/.test(nextStr)) {
          break;
        }
        if (/\d{1,2}[:：]\d{2}|[①②③④⑤]|シフト制|交代制|実働/.test(nextStr)) {
          collected.push(nextStr);
        }
      }
      if (collected.length > 0) {
        shiftTime = sanitizeExcelFloatText(stripKyujinFormBoilerplate(collected.join(' / ')));
        break;
      }
    }
  }

  // 2) Text block extraction for numbered shifts (①...②...③...④...)
  if (!shiftTime) {
    const shiftBlockM = cleanedText.match(
      /(?:就業時間|勤務時間|シフト)[：:\s]*([\s\S]{4,220}?)(?=\n(?:休憩|時間外|残業|休日|年間休日|給与|月給|時給|賞与|福利厚生|勤務地|応募|$))/i
    );
    if (shiftBlockM) {
      shiftTime = shiftBlockM[1]
        .split(/\r?\n/)
        .map((l) => l.trim())
        .filter(Boolean)
        .join(' / ');
    } else {
      const numberedShifts = cleanedText.match(/[①②③④⑤][^\n]{4,45}/g);
      if (numberedShifts && numberedShifts.length > 0) {
        shiftTime = numberedShifts.map((s) => s.trim()).join(' / ');
      }
    }
  }

  const holidayM =
    cleanedText.match(/(?:年間休日日数|年間休日|休日(?:・休暇)?)[：:\s]*([^\n]{2,60})/i);
  const holidays = holidayM ? holidayM[1].trim() : '';

  const overtimeM =
    cleanedText.match(/(?:残業(?:時間)?|時間外労働(?:時間)?)[：:\s]*([^\n]{2,45})/i);
  const overtimeHours = overtimeM ? overtimeM[1].trim() : '';

  const estimatedNetSalary = extractEstimatedNetSalaryText(cleanedText);
  const annualSalary = extractAnnualSalaryText(cleanedText);

  const recruitM = cleanedText.match(/(?:募集人数|採用予定人数|募集枠)[：:\s]*([^\n]{1,25})/i);
  const recruitCount = recruitM ? recruitM[1].trim() : '';

  const interviewM = cleanedText.match(/(?:選考方法|面接回数|選考プロセス)[：:\s]*([^\n]{2,60})/i);
  const interviewProcess = interviewM ? interviewM[1].trim() : '';

  const benefitsM = cleanedText.match(/(?:福利厚生|諸手当|特別待遇|資格取得支援)[：:\s]*([^\n]{2,120})/i);
  const specialBenefits = benefitsM ? benefitsM[1].trim() : '';

  const shiftType = detectKyujinShiftType(shiftTime, cleanedText);

  return {
    shiftTime,
    shiftType,
    holidays,
    overtimeHours,
    estimatedNetSalary,
    annualSalary,
    recruitCount,
    interviewProcess,
    specialBenefits,
  };
};

const parseExcelBlockToKyujin = (
  blockRows: unknown[][] = [],
  sheetName = '',
  anchors: ExcelAnchor[] = [],
  rowOffset = 0,
  fileName = ''
) => {
  let companyName = '';
  let jobCode = '';
  let workArea = '';
  let fallbackLocation = '';
  let jobDescription = '';
  let bonusText = '';

  for (let r = 0; r < blockRows.length; r++) {
    const row = blockRows[r] || [];
    for (let c = 0; c < row.length; c++) {
      const cell = String(row[c] || '').trim();
      if (!cell) continue;
      const rightValues = row
        .slice(c + 1)
        .map((v) => sanitizeExcelFloatText(v))
        .filter(Boolean);

      if (!companyName && /^(?:企業名|会社名|法人名|施設名|事業所名)$/.test(cell) && rightValues.length > 0) {
        companyName = rightValues[0].replace(/\s*代表者.*$/, '').trim();
      }
      if (!jobCode && /^案件番号$/.test(cell) && rightValues.length > 0) {
        jobCode = rightValues[0].trim();
      }
      if (!workArea && /^(?:勤務先|勤務地|就業場所)$/.test(cell) && rightValues.length > 0) {
        workArea = rightValues.join(' ').trim();
      }
      if (!fallbackLocation && /^所在地$/.test(cell) && rightValues.length > 0) {
        fallbackLocation = rightValues.join(' ').trim();
      }
      if (!jobDescription && /主な仕事内容|仕事内容|業務内容/.test(cell) && rightValues.length > 0) {
        jobDescription = rightValues.join(' ').trim();
      }
      if (/賞与/.test(cell) && rightValues.length > 0) {
        bonusText += ' ' + rightValues.join(' ');
      }
    }
  }

  const fullText = sanitizeExcelFloatText(
    blockRows.map((r) => (r || []).map((v) => sanitizeExcelFloatText(v)).join(' ')).join('\n')
  );
  if (!companyName && !jobCode && !workArea && !/求人票|特定技能|給与|月給|時給|賃金/.test(fullText)) {
    return null;
  }

  const finalCompany =
    companyName ||
    fileName
      .replace(/^求人票[_・\-\s]*/i, '')
      .replace(/様.*$/, '')
      .replace(/\.[^.]+$/, '')
      .trim() ||
    '新規取込企業';

  const cleanArea = formatPublicArea(workArea || fallbackLocation);
  const industry = detectIndustryFromText(
    `${fullText.slice(0, 260)} ${jobDescription}`,
    `${fileName} ${sheetName}`
  );
  const sector = detectSectorFromIndustryOrText(industry, `${fullText} ${jobDescription}`);
  const employmentType = detectEmploymentTypeFromText(fullText, `${fileName} ${sheetName}`);
  const minJlpt = extractJlptLevel(blockRows, anchors, rowOffset);
  const housingSupport = extractHousingFromRows(blockRows, anchors, rowOffset);
  const salaryParsed = extractKyujinSalaryFromRows(blockRows);
  const salary = salaryParsed.salary;
  const bonusNorm = normalizeBonusInfo('', '', `${bonusText} ${fullText}`);
  const hourlyM = fullText.match(/時給[：:\s]*([0-9,]+\s*円[^\n]*)/);
  const resolvedHourlyWage =
    salaryParsed.hourlyWage || (hourlyM ? extractKyujinSalaryInfo(`時給 ${hourlyM[1]}`).hourlyWage || hourlyM[1].trim() : '');
  const finalEmploymentType =
    salaryParsed.salaryType === 'hourly' && !/tokutei|特定技能|正社員/i.test(fullText)
      ? 'Baito / Làm thêm (アルバイト・パート)'
      : employmentType;
  const specs = extractDetailedKyujinSpecs(fullText, blockRows);
  const netM = fullText.match(/(?:手取り|実収入|ước tính thực nhận)[：:\s]*([^\n]{3,30})/i);
  const nightM = fullText.match(/(?:夜勤手当)[：:\s]*([^\n]{2,30})/);
  const expReq = normalizeKyujinExperienceRequirement('', fullText);

  const conds = extractNormalizedKyujinConditions({
    text: fullText,
    minJlpt,
    employmentType: finalEmploymentType,
  });

  return {
    companyName: finalCompany,
    sector,
    industry,
    employmentType: finalEmploymentType,
    recruitmentStatus: '募集中',
    excludedNationalities: conds.excludedNationalities,
    excludedGenders: conds.excludedGenders,
    requiresInJapan: conds.requiresInJapan,
    requiresTokutei: conds.requiresTokutei,
    requiredConditions: conds.requiredConditions,
    ngConditions: conds.ngConditions,
    slot: {
      slotId: `slot_xl_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      sector,
      industry,
      employmentType: finalEmploymentType,
      jobCode: jobCode || sheetName || cleanArea || '新規枠',
      area: cleanArea || '勤務地要確認',
      shiftTime: specs.shiftTime,
      shiftType: specs.shiftType,
      holidays: specs.holidays,
      overtimeHours: specs.overtimeHours,
      salary,
      nightShiftAllowance: nightM ? nightM[1].trim() : '',
      estimatedNetSalary: specs.estimatedNetSalary || (netM ? netM[1].trim() : ''),
      annualSalary: specs.annualSalary,
      hourlyWage: resolvedHourlyWage,
      bonus: bonusNorm.bonus,
      bonusDetail: bonusNorm.bonusDetail,
      minJlpt,
      housingSupport,
      recruitmentStatus: '募集中',
      recruitCount: specs.recruitCount,
      interviewProcess: specs.interviewProcess,
      specialBenefits: specs.specialBenefits,
      careQualification: '',
      experienceRequirement: expReq,
      jobDescription: sanitizeJobDescriptionForCandidate(jobDescription),
      excludedNationalities: conds.excludedNationalities,
      excludedGenders: conds.excludedGenders,
      requiresInJapan: conds.requiresInJapan,
      requiresTokutei: conds.requiresTokutei,
      requiredConditions: conds.requiredConditions,
      ngConditions: conds.ngConditions,
    },
  };
};

export const parseExcelFileToKyujinEntries = async (file: File) => {
  const lower = file.name.toLowerCase();
  const buf = await file.arrayBuffer();
  let sheetAnchors: Record<string, ExcelAnchor[]> = {};
  if (lower.endsWith('.xlsx')) {
    try {
      const zip = await JSZip.loadAsync(buf);
      sheetAnchors = await extractExcelCheckedAnchors(zip);
    } catch {
      // fallback without vml anchors
    }
  }
  const wb = XLSX.read(buf, { type: 'array' });
  const entries: Array<{
    companyName: string;
    sector?: string;
    industry: string;
    employmentType?: string;
    recruitmentStatus: string;
    slot: KyujinSlot;
  }> = [];

  wb.SheetNames.forEach((sName, sIdx) => {
    const ws = wb.Sheets[sName];
    const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: '' });
    const anchors = sheetAnchors[`sheet${sIdx + 1}`] || [];
    const blocks = splitSheetRowsIntoKyujinBlocks(rows);
    blocks.forEach((blk) => {
      const parsed = parseExcelBlockToKyujin(
        blk.rows,
        sName,
        anchors,
        blk.startRow,
        file.name
      );
      if (parsed) entries.push(parsed);
    });
  });

  return entries;
};

export const extractWordTextViaMammoth = async (file: File): Promise<string> => {
  const buf = await file.arrayBuffer();
  try {
    const res = await mammoth.extractRawText({ arrayBuffer: buf });
    const text = String(res.value || '').trim();
    if (text) return text;
  } catch {
    // Fallback to raw document.xml parsing if mammoth hits edge docx
  }
  const zip = await JSZip.loadAsync(buf);
  const docXml = zip.files['word/document.xml']
    ? await zip.files['word/document.xml'].async('string')
    : '';
  return docXml
    .replace(/<\/w:p>/g, '\n')
    .replace(/<\/w:tc>/g, ' | ')
    .replace(/<[^>]+>/g, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
};

export const extractRirekishoTextFromLocalFile = async (
  file: File
): Promise<Array<{ text?: string; inlineData?: { mimeType: string; data: string } }>> => {
  const lower = file.name.toLowerCase();
  if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const allSheetsText = wb.SheetNames.map((sName) => {
      const rows = XLSX.utils.sheet_to_json<unknown[]>(wb.Sheets[sName], {
        header: 1,
        defval: '',
      });
      const lines = rows
        .map((r) => r.map((c) => String(c || '').trim()).filter(Boolean).join(' | '))
        .filter(Boolean);
      return `【シート: ${sName}】\n${lines.join('\n')}`;
    }).join('\n\n');
    return [{ text: `以下はExcel形式の既存履歴書から抽出した全セル内容です:\n\n${allSheetsText}` }];
  }

  if (lower.endsWith('.docx')) {
    const plainWordText = await extractWordTextViaMammoth(file);
    if (!plainWordText) {
      throw new Error('Word文書 (.docx) からテキストを抽出できませんでした');
    }
    return [{ text: `以下はWord (.docx) 形式の履歴書内容です:\n\n${plainWordText}` }];
  }

  if (lower.endsWith('.csv') || lower.endsWith('.txt') || file.type.startsWith('text/')) {
    const rawText = await file.text();
    return [{ text: `以下はテキスト/CSV形式の履歴書データです:\n\n${rawText}` }];
  }

  if (lower.endsWith('.doc')) {
    throw new Error(
      '旧形式の .doc ファイルは直接読み込めません。.docx または PDF に変換してアップロードしてください。'
    );
  }

  if (file.type.startsWith('image/') || /\.(png|jpe?g|webp|gif)$/i.test(lower)) {
    const compressed = await compressImageFileToBase64(file, 1600, 0.78);
    return [{ inlineData: { mimeType: compressed.mimeType, data: compressed.base64Data } }];
  }

  if (file.size > 18 * 1024 * 1024) {
    throw new Error('File PDF vượt quá giới hạn 18MB. Vui lòng nén hoặc tách nhỏ PDF.');
  }

  const base64Data = await fileToBase64(file);
  const mimeType = file.type || (lower.endsWith('.pdf') ? 'application/pdf' : 'image/png');
  return [{ inlineData: { mimeType, data: base64Data } }];
};

export const fileToBase64 = (file: Blob): Promise<string> => fileToBase64Raw(file);

export const parseRawKyujinTextFallback = (rawText: string, sourceTitle = 'Text Nhập Tay') => {
  const cleanedRaw = sanitizeExcelFloatText(stripKyujinFormBoilerplate(rawText));
  const blocks = cleanedRaw
    .split(/(?:\n\s*[-=]{4,}\s*\n|\n{3,}(?=(?:企業名|会社名|【|■|Đơn|Job)))/i)
    .map((b) => b.trim())
    .filter((b) => b.length >= 8);
  const targetBlocks = blocks.length > 0 ? blocks : [cleanedRaw.trim()];

  return targetBlocks.map((block, idx) => {
    const compM = block.match(/(?:企業名|会社名|Công ty)[：:\s]*([^\n]+)/i);
    const areaM = block.match(/(?:勤務地|勤務先|エリア|Khu vực|Địa điểm)[：:\s]*([^\n]+)/i);
    const jlptM =
      block.match(/(?:日本語(?:能力|レベル)?|JLPT)[^\n]{0,20}\b(N[1-5])\s*以上/i) ||
      block.match(/\b(N[1-5])\b/i);
    const salaryInfo = extractKyujinSalaryInfo(block);
    const specs = extractDetailedKyujinSpecs(block);
    const bonusNorm = normalizeBonusInfo('', '', block);
    const industry = detectIndustryFromText(block, sourceTitle);
    const sector = detectSectorFromIndustryOrText(industry, block);
    let employmentType = detectEmploymentTypeFromText(block, sourceTitle);
    if (
      salaryInfo.salaryType === 'hourly' &&
      !/tokutei|特定技能|正社員/i.test(block)
    ) {
      employmentType = 'Baito / Làm thêm (アルバイト・パート)';
    }
    const inferredArea = areaM ? areaM[1].trim() : extractPreferredAreaFromChat(block) || 'Nhật Bản';

    const resolvedMinJlpt = jlptM ? jlptM[1].toUpperCase() : '要確認';
    const conds = extractNormalizedKyujinConditions({
      text: block,
      minJlpt: resolvedMinJlpt,
      employmentType,
    });

    return {
      companyName: compM ? compM[1].trim() : `${sourceTitle} #${idx + 1}`,
      sector,
      industry,
      employmentType,
      jobCode: `JOB-${idx + 1}`,
      area: inferredArea,
      shiftTime: specs.shiftTime,
      shiftType: specs.shiftType,
      holidays: specs.holidays,
      overtimeHours: specs.overtimeHours,
      salary: salaryInfo.salary,
      estimatedNetSalary: specs.estimatedNetSalary,
      annualSalary: specs.annualSalary,
      hourlyWage: salaryInfo.hourlyWage,
      minJlpt: resolvedMinJlpt,
      housingSupport: /自社寮/.test(block)
        ? '自社寮あり'
        : /借上げ|借り上げ/.test(block)
        ? '借上げ宿舎あり'
        : /寮あり|社宅/.test(block)
        ? '寮・社宅あり'
        : /提供無|寮なし|社宅なし/.test(block)
        ? '提供無'
        : '要確認',
      bonus: bonusNorm.bonus,
      bonusDetail: bonusNorm.bonusDetail,
      recruitCount: specs.recruitCount,
      interviewProcess: specs.interviewProcess,
      specialBenefits: specs.specialBenefits,
      experienceRequirement: normalizeKyujinExperienceRequirement('', block),
      jobDescription: sanitizeJobDescriptionForCandidate(block.slice(0, 800)),
      excludedNationalities: conds.excludedNationalities,
      excludedGenders: conds.excludedGenders,
      requiresInJapan: conds.requiresInJapan,
      requiresTokutei: conds.requiresTokutei,
      requiredConditions: conds.requiredConditions,
      ngConditions: conds.ngConditions,
    };
  });
};

export const extractKyujinViaServer = async (
  file: File,
  idToken?: string | null,
  onProgress?: (statusText: string) => void
) => {
  const lower = file.name.toLowerCase();
  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
  };
  if (idToken) {
    headers.Authorization = `Bearer ${idToken}`;
  }

  const callKyujinApi = async (bodyPayload: Record<string, unknown>) => {
    const response = await fetch('/api/gemini/extract-kyujin', {
      method: 'POST',
      headers,
      body: JSON.stringify(bodyPayload),
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(data?.error || `HTTP ${response.status}`);
    }
    return Array.isArray(data.items) ? data.items : [];
  };

  if (lower.endsWith('.pdf') || file.type === 'application/pdf') {
    if (file.size > 30 * 1024 * 1024) {
      throw new Error(`File PDF "${file.name}" vượt quá giới hạn 30MB.`);
    }

    const pdfResult = await extractPdfWithOcrFallback(file, onProgress);

    if (pdfResult.mode === 'text') {
      onProgress?.(`Đang phân tích nội dung tiếng Nhật (${pdfResult.numPages} trang)...`);
      try {
        const items = await callKyujinApi({
          rawText: pdfResult.text,
          fileName: file.name,
        });
        if (items.length > 0) return items;
      } catch {
        // Fallback to local Japanese Kyujin parser when Gemini key is absent/unavailable
      }
      return parseRawKyujinTextFallback(pdfResult.text, file.name);
    }

    // Scanned PDF -> OCR via Vision AI on rendered page images (same pipeline as dropped/pasted images)
    const rendered = pdfResult.renderedImages;
    if (rendered.length === 0) {
      throw new Error(`Không thể dựng ảnh trang từ file PDF scan "${file.name}".`);
    }

    const batchSize = 4;
    const allItems: Array<Record<string, string>> = [];
    for (let i = 0; i < rendered.length; i += batchSize) {
      const batch = rendered.slice(i, i + batchSize);
      const startPage = batch[0].pageNumber;
      const endPage = batch[batch.length - 1].pageNumber;
      onProgress?.(
        `Đang nhận diện OCR AI trang ${startPage}${endPage > startPage ? `–${endPage}` : ''}/${rendered.length}...`
      );
      const batchItems = await callKyujinApi({
        images: batch.map((img) => ({
          mimeType: img.mimeType,
          base64Data: img.base64Data,
        })),
        fileName: file.name,
      });
      allItems.push(...batchItems);
    }
    return allItems;
  }

  let payload: Record<string, unknown> = { fileName: file.name };

  if (lower.endsWith('.docx')) {
    onProgress?.('Đang trích xuất văn bản Word (.docx)...');
    const rawText = await extractWordTextViaMammoth(file);
    try {
      const items = await callKyujinApi({ rawText, fileName: file.name });
      if (items.length > 0) return items;
    } catch {
      // Fallback to local parser for .docx when Gemini key is absent
    }
    return parseRawKyujinTextFallback(rawText, file.name);
  } else if (lower.endsWith('.txt') || lower.endsWith('.csv') || file.type.startsWith('text/')) {
    onProgress?.('Đang đọc dữ liệu văn bản / CSV...');
    const rawText = await file.text();
    payload = { rawText, fileName: file.name };
  } else if (file.type.startsWith('image/') || /\.(png|jpe?g|webp|gif)$/i.test(lower)) {
    onProgress?.('Đang nén & gửi ảnh qua OCR AI...');
    const compressed = await compressImageFileToBase64(file, 1600, 0.78);
    payload = {
      mimeType: compressed.mimeType,
      base64Data: compressed.base64Data,
      fileName: file.name,
    };
  } else {
    if (file.size > 18 * 1024 * 1024) {
      throw new Error(`File "${file.name}" vượt quá 18MB.`);
    }
    const base64Data = await fileToBase64(file);
    const mimeType =
      file.type || (lower.endsWith('.pdf') ? 'application/pdf' : 'image/png');
    payload = { mimeType, base64Data, fileName: file.name };
  }

  return callKyujinApi(payload);
};

export const extractKyujinRawTextViaServer = async (
  rawText: string,
  idToken?: string | null
) => {
  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
  };
  if (idToken) {
    headers.Authorization = `Bearer ${idToken}`;
  }

  const response = await fetch('/api/gemini/extract-kyujin', {
    method: 'POST',
    headers,
    body: JSON.stringify({ rawText, fileName: 'PastedKyujinText' }),
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || `HTTP ${response.status}`);
  }
  return Array.isArray(data.items) ? data.items : [];
};

export const buildRirekishoFallbackFromText = (
  rawText = '',
  person: Partial<Candidate> = {}
): Partial<RirekishoDraft> => {
  const base = createInitialRirekishoDraft(person);
  const text = String(rawText || '').trim();
  if (!text) return base;

  const phoneM = text.match(/(0[789]0[-\s]?\d{4}[-\s]?\d{4})/);
  const emailM = text.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/);
  const postalM = text.match(/〒?\s*(\d{3}-\d{4})/);
  const jlptM = text.match(/\b(N[1-5])\b/i);
  const birthM = text.match(/(?:生年月日|sinh ngày|birth)[：:\s]*(\d{4}[年\/\-]\d{1,2}[月\/\-]\d{1,2})/i);
  const visaExpM = text.match(/(?:在留期限|hạn visa|expiry)[：:\s]*(\d{4}[年\/\-]\d{1,2}[月\/\-]\d{1,2})/i);

  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const workMatches = lines.filter((l) => /入社|退社|勤務|介護|レストラン|アルバイト|công ty|làm việc/i.test(l)).slice(0, 5);
  const eduMatches = lines.filter((l) => /大学|高校|専門学校|日本語学校|trường|tốt nghiệp/i.test(l)).slice(0, 4);

  return {
    ...base,
    phone: phoneM ? phoneM[1] : base.phone,
    email: emailM ? emailM[0] : base.email,
    postalCode: postalM ? postalM[1] : base.postalCode,
    birthDate: birthM ? birthM[1].replace(/[年月]/g, '-').replace(/日/g, '') : base.birthDate,
    visaExpiry: visaExpM ? visaExpM[1].replace(/[年月]/g, '-').replace(/日/g, '') : base.visaExpiry,
    availableLanguages: jlptM ? `日本語（${jlptM[1].toUpperCase()}）` : base.availableLanguages,
    educationRows:
      eduMatches.length > 0
        ? eduMatches.map((line) => ({ year: '', month: '', content: line }))
        : base.educationRows,
    workRows:
      workMatches.length > 0
        ? workMatches.map((line) => ({ year: '', month: '', content: line }))
        : base.workRows,
  };
};

export const extractRirekishoViaServer = async (
  parts: Array<{ text?: string; inlineData?: { mimeType: string; data: string } }>,
  idToken?: string | null
) => {
  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
  };
  if (idToken) {
    headers.Authorization = `Bearer ${idToken}`;
  }

  const response = await fetch('/api/gemini/extract-rirekisho', {
    method: 'POST',
    headers,
    body: JSON.stringify({ parts }),
  });
  const data = await response.json();
  if (!response.ok) {
    const err = new Error(data?.error || `HTTP ${response.status}`) as Error & { code?: string };
    err.code = data?.code;
    throw err;
  }
  return data.extracted || {};
};

const normalizeExtractedDateYmd = (raw = ''): string => {
  const s = String(raw || '').trim();
  if (!s) return '';
  const ymd = s.match(/^(\d{4})[年\/\-\.](\d{1,2})[月\/\-\.](\d{1,2})日?$/);
  if (ymd) {
    return `${ymd[1]}-${ymd[2].padStart(2, '0')}-${ymd[3].padStart(2, '0')}`;
  }
  const dmy = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
  if (dmy) {
    const p1 = Number(dmy[1]);
    const p2 = Number(dmy[2]);
    // Assume DD/MM/YYYY unless second segment > 12 (MM/DD/YYYY)
    if (p2 > 12 && p1 <= 12) {
      return `${dmy[3]}-${String(p1).padStart(2, '0')}-${String(p2).padStart(2, '0')}`;
    }
    return `${dmy[3]}-${String(p2).padStart(2, '0')}-${String(p1).padStart(2, '0')}`;
  }
  return s;
};

export interface ServerCandidateExtractionResponse extends Partial<Candidate> {
  fullName?: string;
  japaneseLevel?: string;
  desiredField?: string;
  visaType?: string;
  certifications?: string;
  goal?: string;
  rawLog?: string;
  fbUrl?: string;
}

const normalizeCandidateJlptValue = (rawVal?: string): string => {
  const s = String(rawVal || '').trim();
  if (!s) return '';
  if (/^JFT/i.test(s)) return 'JFT';
  const nMatch = s.match(/^(N[1-5])$/i);
  if (nMatch) return nMatch[1].toUpperCase();
  return s;
};

export const extractCandidateViaServer = async (
  payload: {
    text?: string;
    parts?: Array<{ text?: string; inlineData?: { mimeType: string; data: string } }>;
  },
  idToken?: string | null
): Promise<Partial<Candidate>> => {
  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
  };
  if (idToken) {
    headers.Authorization = `Bearer ${idToken}`;
  }

  const response = await fetch('/api/gemini/extract-candidate', {
    method: 'POST',
    headers,
    body: JSON.stringify(payload),
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    const err = new Error(data?.error || `HTTP ${response.status}`) as Error & { code?: string };
    err.code = data?.code;
    throw err;
  }
  const rawCand = (data.candidate || {}) as ServerCandidateExtractionResponse;
  const rawArea = String(rawCand.desiredArea || '')
    .replace(/都内/g, '東京都')
    .replace(/^関東$/g, '関東エリア')
    .trim();

  return {
    ...rawCand,
    name: String(rawCand.fullName || rawCand.name || '').trim(),
    fbName: String(rawCand.fbName || '').trim(),
    fbProfileUrl: String(rawCand.fbProfileUrl || rawCand.fbUrl || '').trim(),
    nationality: String(rawCand.nationality || '').trim(),
    jlpt: normalizeCandidateJlptValue(rawCand.japaneseLevel || rawCand.jlpt) || '要確認',
    desiredIndustry: String(rawCand.desiredField || rawCand.desiredIndustry || '').trim(),
    desiredArea: rawArea,
    visaStatus: String(rawCand.visaType || rawCand.visaStatus || '').trim(),
    visaExpiry: rawCand.visaExpiry ? normalizeExtractedDateYmd(rawCand.visaExpiry) : '',
    birthDate: rawCand.birthDate ? normalizeExtractedDateYmd(rawCand.birthDate) : '',
    passedExam: String(rawCand.certifications || rawCand.passedExam || '').trim(),
    conversionGoal: String(rawCand.goal || rawCand.conversionGoal || '').trim(),
    negotiationLog: String(rawCand.rawLog || rawCand.negotiationLog || '').trim(),
  };
};

export const recognizeCandidateImageWithTesseract = async (
  imageSource: string | Blob,
  onProgress?: (progressText: string) => void
): Promise<string> => {
  const Tesseract = await import('tesseract.js');
  onProgress?.('Đang khởi tạo bộ đọc OCR đa ngữ (vie+jpn+eng)...');
  const result = await Tesseract.recognize(imageSource, 'vie+jpn+eng', {
    logger: (m: { status?: string; progress?: number }) => {
      if (m && typeof m.progress === 'number' && m.status === 'recognizing text') {
        onProgress?.(`Đang nhận diện chữ trong ảnh (${Math.round(m.progress * 100)}%)...`);
      }
    },
  });
  return String(result?.data?.text || '').trim();
};

// ============================================================================
// CHAT PARSING & SMART REPLY BUILDER
// ============================================================================
const MESSENGER_UI_NOISE_RX =
  /^(?:đã gửi|bạn đã gửi|you sent|đã xem|seen|active now|đang hoạt động|nhắn tin|trả lời|thích|like|\d{1,2}:\d{2}(?:\s*(?:am|pm|sa|ch))?|(?:hôm nay|hôm qua|thứ\s*\w+|chủ nhật|today|yesterday)\s*(?:lúc|at)?\s*\d{0,2}:?\d{0,2}|[👍❤️😂😮😢😡🙏✨]+)$/i;

const CONSULTANT_SPEAKER_LABEL_RX =
  /^(?:bạn|you|admin|staff|tư vấn(?: viên)?|nhân viên|hỗ trợ|cty|công ty|tokutei|page|ad)\s*[:：]/i;

const CONSULTANT_CONTENT_RX =
  /(?:chào bạn[,.!]?\s*bên (?:mình|em|anh|chị)|bên (?:mình|anh|chị) (?:hỗ trợ|có đơn|đang tuyển|gửi bạn|xin thêm)|bạn có (?:n|jlpt|bằng) mấy|thẻ ngoại kiều (?:của )?bạn|bạn gửi (?:mình|ad) xem|mình gửi (?:bạn )?(?:\d+ )?(?:đơn|job|lịch mendan)|hẹn bạn mendan|chúng tôi xin phép)/i;

const CANDIDATE_CONTENT_RX =
  /(?:bên (?:anh|chị|ad|mình) có đơn|em (?:đang|muốn|cần) tìm|mình (?:đang|muốn|cần) tìm|em có (?:n[1-5]|jft)|mình có (?:n[1-5]|jft)|em đang (?:ở|làm)|hạn thẻ (?:của )?(?:em|mình)|lương thực nhận|em tên là|mình tên là|cho em hỏi|xin việc|chuyển visa|chuyển việc)/i;

const extractCandidateLinesOnly = (rawText: string): {
  cleanedLines: string[];
  candidateLines: string[];
  detectedSpeakerFbName: string;
} => {
  const rawLines = rawText
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l && !MESSENGER_UI_NOISE_RX.test(l));

  const candidateLines: string[] = [];
  let currentSpeaker: 'candidate' | 'consultant' | 'unknown' = 'unknown';
  const speakerCounts = new Map<string, { cand: number; staff: number }>();

  for (const line of rawLines) {
    if (CONSULTANT_SPEAKER_LABEL_RX.test(line)) {
      currentSpeaker = 'consultant';
      continue;
    }

    const isRawUrlLine = /^(?:https?|ftp):\/\//i.test(line);
    const turnM = !isRawUrlLine ? line.match(/^([^\n:：]{2,40})[:：]\s*(.+)$/) : null;
    if (turnM) {
      const label = turnM[1].trim();
      const body = turnM[2].trim();
      const isFieldLabel =
        /^(?:https?|ftp|mailto|họ\s*(?:và\s*)?tên|tên\s*(?:khai\s*sinh|fb|facebook)|fb|facebook|sđt|sdt|điện\s*thoại|tel|phone|email|mail|visa|hạn\s*(?:thẻ|visa)|trình\s*độ|jlpt|tiếng\s*nhật|khu\s*vực|tỉnh|ngành|kinh\s*nghiệm|chứng\s*chỉ|lương|mục\s*tiêu|nhu\s*cầu|ghi\s*chú|ngày\s*sinh|tuổi|quốc\s*tịch|địa\s*chỉ|link|địa\s*điểm|yêu\s*cầu|lương\s*cơ\s*bản|氏名|名前|本名|国籍|生年月日|性別|住所|在留資格|在留期限|在留期間(?:\s*[\(（]満了日[\)）])?|希望.*|エリア|給料|給与|月給|時給|経験|資格|合格.*|免許|備考|メモ|勤務.*|応募.*|連絡先)$/i.test(
          label
        );
      if (!isFieldLabel) {
        const stats = speakerCounts.get(label) || { cand: 0, staff: 0 };
        if (CONSULTANT_CONTENT_RX.test(body)) {
          stats.staff += 2;
          currentSpeaker = 'consultant';
        } else if (CANDIDATE_CONTENT_RX.test(body)) {
          stats.cand += 2;
          currentSpeaker = 'candidate';
        } else if (currentSpeaker !== 'consultant') {
          stats.cand += 1;
          currentSpeaker = 'candidate';
        }
        speakerCounts.set(label, stats);
        if (currentSpeaker !== 'consultant') {
          candidateLines.push(body);
        }
        continue;
      }
    }

    if (CONSULTANT_CONTENT_RX.test(line)) {
      currentSpeaker = 'consultant';
      continue;
    }
    if (CANDIDATE_CONTENT_RX.test(line)) {
      currentSpeaker = 'candidate';
    }
    if (currentSpeaker !== 'consultant') {
      candidateLines.push(line);
    }
  }

  let detectedSpeakerFbName = '';
  let bestScore = 0;
  for (const [label, counts] of speakerCounts.entries()) {
    const score = counts.cand - counts.staff;
    if (counts.cand > 0 && score >= bestScore) {
      bestScore = score;
      detectedSpeakerFbName = label;
    }
  }

  return {
    cleanedLines: rawLines,
    candidateLines: candidateLines.length > 0 ? candidateLines : rawLines,
    detectedSpeakerFbName,
  };
};

export const extractPreferredAreaFromChat = (chatText = ''): string => {
  const text = String(chatText || '');
  if (!text.trim()) return '';

  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const wishLineRx =
    /(?:chuyển\s*(?:sang|qua|về|tới|đến)|muốn\s*(?:tìm|làm|về|sang|ở)|tìm\s*(?:đơn|việc|job).*?(?:ở|tại)|khu\s*vực(?:\s*mong\s*muốn)?|tỉnh\s*mong\s*muốn|nguyện\s*vọng|希望|エリア|勤務地|方面|あたり|近く|ないですか|ありますか)/i;

  const wishMatches = new Set<string>();
  for (const line of lines) {
    if (wishLineRx.test(line)) {
      // If line expresses "đang ở A, muốn sang B, C", inspect the target clause first
      const switchParts = line.split(/(?:muốn\s*(?:chuyển|tìm|sang|về|qua)|chuyển\s*(?:sang|qua|về|tới))/i);
      const targetSegment =
        switchParts.length > 1 ? switchParts.slice(1).join(' ') : line;
      let foundInTarget = false;
      for (const item of PREFECTURE_PATTERNS) {
        if (item.rx.test(targetSegment)) {
          wishMatches.add(item.label);
          foundInTarget = true;
        }
      }
      if (!foundInTarget) {
        for (const item of PREFECTURE_PATTERNS) {
          if (item.rx.test(line)) {
            wishMatches.add(item.label);
          }
        }
      }
    }
  }
  if (wishMatches.size > 0) {
    return Array.from(wishMatches).join('・');
  }

  const allMatches: Array<{ label: string; index: number }> = [];
  for (const item of PREFECTURE_PATTERNS) {
    const m = text.match(item.rx);
    if (m) {
      allMatches.push({ label: item.label, index: m.index ?? 9999 });
    }
  }
  if (allMatches.length > 0) {
    allMatches.sort((a, b) => a.index - b.index);
    const uniqueLabels = Array.from(new Set(allMatches.map((m) => m.label)));
    return uniqueLabels.join('・');
  }

  const regions: string[] = [];
  if (/kanto|関東/i.test(text)) regions.push('関東エリア');
  if (/kansai|関西/i.test(text)) regions.push('関西エリア');
  if (/chubu|中部|tokai|東海/i.test(text)) regions.push('中部・東海エリア');
  if (/kyushu|九州/i.test(text)) regions.push('九州エリア');
  return regions.join('・');
};

const isPlaceholderExtractionVal = (val: string): boolean => {
  const s = String(val || '').trim();
  if (!s) return true;
  return /^(?:\(?\s*(?:chưa\s*(?:có|rõ|cập\s*nhật|biết)|không\s*có|trống)\b[^)]*\)?|未定|不明|未設定|未記入|未確認|なし|無|none|n\/a|-+|—+)$/i.test(
    s
  );
};

export const parseChatIntoUpdates = (chatText = '') => {
  const text = String(chatText || '').trim();
  const todayIso = new Date(Date.now() + 9 * 60 * 60 * 1000).toISOString().slice(0, 10);
  const updates: Partial<Candidate> & { memoLine: string } = {
    lastContactDate: todayIso,
    memoLine: '',
  };

  if (!text) return updates;

  const { cleanedLines, candidateLines, detectedSpeakerFbName } =
    extractCandidateLinesOnly(text);
  const candText = candidateLines.join('\n');

  // 1. Facebook Profile / Messenger URL
  const fbUrlM = text.match(
    /https?:\/\/(?:www\.|m\.|mbasic\.)?(?:facebook\.com|fb\.com|m\.me)\/[^\s<>"')\]}]+/i
  );
  if (fbUrlM) {
    updates.fbProfileUrl = fbUrlM[0].replace(/[.,;!?]+$/, '');
  }

  // 2. Legal Name vs Facebook Display Name — strict line-by-line label isolation
  for (const rawLine of text.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line) continue;

    if (!updates.name && !updates.fbName) {
      const parenM = line.match(
        /^([A-ZÀ-Ỹ][^\n()]{3,45})\s*\((?:fb|facebook)[：:\s]*([^)\n]+)\)/i
      );
      if (parenM) {
        const nm = parenM[1].trim();
        const fbn = parenM[2].trim();
        if (!isPlaceholderExtractionVal(nm) && !CONSULTANT_CONTENT_RX.test(nm)) {
          updates.name = nm;
        }
        if (
          !isPlaceholderExtractionVal(fbn) &&
          !/https?:\/\//i.test(fbn) &&
          !CONSULTANT_CONTENT_RX.test(fbn)
        ) {
          updates.fbName = fbn;
        }
      }
    }

    if (!updates.name) {
      const legalLineM = line.match(
        /^(?:[-・*•\s]*)?(?:họ\s*(?:và\s*)?tên(?:\s*khai\s*sinh)?|tên\s*khai\s*sinh|full\s*name|氏名|本名|名前)\s*[：:]\s*(.+)$/i
      );
      if (legalLineM) {
        const candidateName = legalLineM[1].split(/[;,]/)[0].trim();
        if (
          !isPlaceholderExtractionVal(candidateName) &&
          !/https?:\/\//i.test(candidateName) &&
          !CONSULTANT_CONTENT_RX.test(candidateName)
        ) {
          updates.name = candidateName;
        }
      }
    }

    if (!updates.fbName) {
      // Never match lines labeled Link / URL / リンク
      const isUrlLabelLine = /(?:link|url|リンク)/i.test(line.split(/[：:]/)[0] || '');
      if (!isUrlLabelLine) {
        const fbLineM = line.match(
          /^(?:[-・*•\s]*)?(?:tên\s*(?:nick\s*)?(?:fb|facebook|messenger)|nick\s*(?:fb|facebook)|fbアカウント名|fb名|fb|facebook)\s*[：:]\s*(.+)$/i
        );
        if (fbLineM) {
          const candidateFb = fbLineM[1].split(/[;,]/)[0].trim();
          if (
            !isPlaceholderExtractionVal(candidateFb) &&
            !/https?:\/\//i.test(candidateFb) &&
            !CONSULTANT_CONTENT_RX.test(candidateFb)
          ) {
            updates.fbName = candidateFb;
          }
        }
      }
    }
  }

  // Check self-introduction patterns in candidate turns only: "em tên là ..." / "mình tên ..."
  if (!updates.name) {
    const selfIntroM = candText.match(
      /(?:em|mình|tôi)\s*(?:tên\s*(?:là|thật\s*là)?|là)\s+([A-ZÀ-Ỹ][a-zà-ỹ]+(?:\s+[A-ZÀ-Ỹ][a-zà-ỹ]+){1,4}|[A-Z]{2,}(?:\s+[A-Z]{2,}){1,4})/
    );
    if (selfIntroM) {
      const introName = selfIntroM[1].trim();
      if (!/^(?:Em|Mình|Tôi|Bạn|Anh|Chị)\b/i.test(introName) && !isPlaceholderExtractionVal(introName)) {
        updates.name = introName;
      }
    }
  }

  // Fallback for FB Nickname ONLY (do NOT promote to legal name)
  if (
    !updates.fbName &&
    detectedSpeakerFbName &&
    !/^(?:https?|ftp|mailto)$/i.test(detectedSpeakerFbName) &&
    !/https?:\/\//i.test(detectedSpeakerFbName) &&
    !isPlaceholderExtractionVal(detectedSpeakerFbName)
  ) {
    updates.fbName = detectedSpeakerFbName;
  } else if (
    !updates.fbName &&
    !updates.name &&
    cleanedLines.length > 0 &&
    cleanedLines[0].length <= 40 &&
    !/[:：。?!？！]/.test(cleanedLines[0]) &&
    !/https?:\/\//i.test(cleanedLines[0]) &&
    !isPlaceholderExtractionVal(cleanedLines[0]) &&
    !CONSULTANT_CONTENT_RX.test(cleanedLines[0]) &&
    !CANDIDATE_CONTENT_RX.test(cleanedLines[0]) &&
    !/^(?:em|anh|chị|bạn|chào|alo|hello|hi|xin chào)\b/i.test(cleanedLines[0])
  ) {
    updates.fbName = cleanedLines[0];
  }

  // 2b. Nationality (explicit label / country keywords BEFORE Vietnamese heuristic!)
  const explicitNatM = text.match(/(?:quốc\s*tịch|国籍|nationality)[：:\s]+([^\n,;]+)/i);
  const explicitNatCleaned = explicitNatM
    ? normalizeCandidateNationality(explicitNatM[1])
    : '';
  if (explicitNatCleaned && explicitNatCleaned !== '国籍不明') {
    updates.nationality = explicitNatCleaned;
  } else if (/indonesia|インドネシア/i.test(text)) {
    updates.nationality = 'インドネシア';
  } else if (/myanmar|ミャンマー|miến\s*điện/i.test(text)) {
    updates.nationality = 'ミャンマー';
  } else if (/nepal|ネパール/i.test(text)) {
    updates.nationality = 'ネパール';
  } else if (/philippines?|フィリピン/i.test(text)) {
    updates.nationality = 'フィリピン';
  } else if (/cambodia|カンボジア|campuchia/i.test(text)) {
    updates.nationality = 'カンボジア';
  } else if (/china|中国|trung\s*quốc/i.test(text)) {
    updates.nationality = '中国';
  } else if (/bangladesh|バングラデシュ/i.test(text)) {
    updates.nationality = 'バングラデシュ';
  } else if (/sri\s*lanka|スリランカ/i.test(text)) {
    updates.nationality = 'スリランカ';
  } else if (/viet\s*nam|việt\s*nam|ベトナム/i.test(text)) {
    updates.nationality = 'ベトナム';
  } else {
    const hasVietnameseSignal =
      /[àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]|\b(?:dhs|em|mình|tìm\s+việc|chuyển\s+việc|tốt\s+nghiệp|năm\s+sau|kinh\s+nghiệm|nguyen|duong|tran|pham|hoang)\b/i.test(
        candText || text
      );
    if (hasVietnameseSignal) {
      updates.nationality = 'ベトナム';
    }
  }

  // Channel detection across personal / company / line / other
  if (/fanpage|page\s*công\s*ty|fb\s*công\s*ty|公式/i.test(text)) {
    updates.channel = 'company';
  } else if (/\bline\b|ライン/i.test(text) && !updates.fbProfileUrl && !updates.fbName) {
    updates.channel = 'line';
  } else if (
    /(?:zalo|tiktok|trực\s*tiếp|giới\s*thiệu|直接紹介|知人紹介|その他)/i.test(text) &&
    !updates.fbProfileUrl
  ) {
    updates.channel = 'other';
  } else if (updates.fbProfileUrl || updates.fbName || /fb|facebook|messenger/i.test(text)) {
    updates.channel = 'personal';
  } else {
    updates.channel = 'other';
  }
  updates.progressStatus = '新規相談';

  // 3. Language (JLPT / JFT / 初級) — NEVER promote upcoming target exam OR job-post requirement (e.g. N3/N4 in job card) as candidate level!
  const upcomingExamM = candText.match(
    /(?:sắp\s*(?:tới)?|chuẩn\s*bị|đang\s*ôn|dự\s*định|tháng\s*\d{1,2}(?:\/\d{4})?)\s*(?:sẽ\s*)?thi\s*(?:jlpt\s*)?(N[1-5])/i
  );
  const upcomingExamNote = upcomingExamM
    ? `Sắp tới thi JLPT ${upcomingExamM[1].toUpperCase()}`
    : '';

  const nonJobReqLines = candidateLines
    .filter(
      (l) =>
        !/(?:募集|応募条件|求める|資格要件|対象者|時給|月給|基本給|年収|仕事内容|勤務時間|シフト|N[1-5]\s*[〜～\/・]\s*N[1-5]|N[1-5]\s*(?:以上|程度|相当)|yêu\s*cầu\s*N[1-5]|tuyển\s*N[1-5])/i.test(
          l
        )
    )
    .join('\n');

  const candTextCurrentJlpt = nonJobReqLines.replace(
    /(?:sắp\s*(?:tới)?|chuẩn\s*bị|đang\s*ôn|dự\s*định|tháng\s*\d{1,2}(?:\/\d{4})?)\s*(?:sẽ\s*)?thi\s*(?:jlpt\s*)?N[1-5]/gi,
    ''
  );
  const jlptM = candTextCurrentJlpt.match(/\b(N[1-5]|JFT(?:\s*Basic)?)\b/i);
  if (jlptM) {
    const rawLevel = jlptM[1].toUpperCase().replace(/\s+/g, ' ');
    updates.jlpt = rawLevel.startsWith('JFT') ? 'JFT' : rawLevel;
  } else if (/初級|sơ\s*cấp/i.test(nonJobReqLines)) {
    updates.jlpt = '初級';
  } else if (/giao tiếp cơ bản|日常会話/i.test(nonJobReqLines)) {
    updates.jlpt = '日常会話';
  } else {
    updates.jlpt = '要確認';
  }

  // 4. Phone / LINE & Email (prefer candidate turns so consultant signature isn't grabbed)
  const phoneSource = /(?:0[789]0|0\d{1,3})[-\s]?\d{3,4}[-\s]?\d{4}/.test(candText)
    ? candText
    : text;
  const phoneM = phoneSource.match(/(0[789]0[-\s]?\d{4}[-\s]?\d{4})/);
  if (phoneM) {
    updates.phone = phoneM[1].replace(/\s+/g, '-');
  } else {
    const lineM = candText.match(/(?:line(?:\s*id)?)[：:\s]+([A-Za-z0-9._-]{3,40})/i);
    if (lineM) updates.phone = `LINE: ${lineM[1]}`;
  }

  const emailSource = /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/.test(candText)
    ? candText
    : text;
  const emailM = emailSource.match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/);
  if (emailM) updates.email = emailM[0];

  // 5. Birth Date & Age
  const birthYmdM = candText.match(
    /(?:生年月日|ngày\s*sinh|sinh\s*ngày|dob)[：:\s]*(\d{4}[年\/\-]\d{1,2}[月\/\-]\d{1,2}日?|\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{4})/i
  );
  if (birthYmdM) {
    updates.birthDate = normalizeExtractedDateYmd(birthYmdM[1]);
  }
  const ageM = candText.match(/(\d{2})\s*(?:tuổi|歳)/i);
  if (ageM) updates.age = ageM[1];

  // 6. Visa Expiry (careful DD/MM/YYYY vs YYYY/MM/DD & Zairyu Card 在留期間(満了日))
  const expPrefix =
    '(?:在留期限|在留期間(?:\\s*[\\(（]満了日[\\)）])?|ビザ期限|hạn\\s*(?:thẻ\\s*ngoại\\s*kiều|thẻ|visa|ngoại\\s*kiều)?|thẻ\\s*(?:ngoại\\s*kiều\\s*)?(?:hết\\s*hạn|đến)|hết\\s*hạn)(?:\\s*ngày)?';
  const expYmdM =
    candText.match(
      new RegExp(`${expPrefix}[：:\\s]*(\\d{4}[年\\/\\-]\\d{1,2}[月\\/\\-]\\d{1,2}日?)`, 'i')
    ) ||
    text.match(
      new RegExp(`${expPrefix}[：:\\s]*(\\d{4}[年\\/\\-]\\d{1,2}[月\\/\\-]\\d{1,2}日?)`, 'i')
    );
  const expDmyM =
    candText.match(
      new RegExp(`${expPrefix}[：:\\s]*(\\d{1,2}[\\/\\-\\.]\\d{1,2}[\\/\\-\\.]\\d{4})`, 'i')
    ) ||
    text.match(
      new RegExp(`${expPrefix}[：:\\s]*(\\d{1,2}[\\/\\-\\.]\\d{1,2}[\\/\\-\\.]\\d{4})`, 'i')
    );
  if (expYmdM) {
    updates.visaExpiry = normalizeExtractedDateYmd(expYmdM[1]);
  } else if (expDmyM) {
    updates.visaExpiry = normalizeExtractedDateYmd(expDmyM[1]);
  }

  // 7. Current Visa Status (includes slang "dhs" -> 留学)
  const isDhs = /\bdhs\b|du\s*học(?:\s*sinh)?|留学/i.test(candText);
  if (/thực\s*tập\s*sinh|tts|技能実習/i.test(candText)) {
    updates.visaStatus = '技能実習（Thực tập sinh）';
  } else if (isDhs) {
    updates.visaStatus = '留学（Du học sinh）';
  } else if (/gia\s*đình|phụ\s*thuộc|家族滞在/i.test(candText)) {
    updates.visaStatus = '家族滞在（Gia đình）';
  } else if (/kỹ\s*sư|giáo\s*dục|nhân\s*văn|技人国|技術・人文/i.test(candText)) {
    updates.visaStatus = '技術・人文知識・国際業務（Kỹ sư）';
  } else if (/định\s*trú|vĩnh\s*trú|定住者|永住者/i.test(candText)) {
    updates.visaStatus = '定住者 / 永住者';
  } else {
    const remPeriodM = candText.match(
      /(?:còn|tokutei\s*(?:còn)?|thời\s*hạn\s*(?:còn)?|残り)?\s*(\d+(?:[.,]\d+)?)\s*(?:n\b|năm|年)/i
    );
    const remYearsNum = remPeriodM ? Number(remPeriodM[1].replace(',', '.')) : 0;
    if (
      remPeriodM &&
      remYearsNum > 0 &&
      remYearsNum <= 10 &&
      /(?:tokutei|特定|còn|残り|visa|ビザ|\d+[.,]\d+\s*n\b)/i.test(candText)
    ) {
      const yrs = remPeriodM[1].replace(',', '.');
      updates.visaStatus = `特定技能1号（残り約${yrs}年）`;
    } else if (/tokutei|特定技能/i.test(candText)) {
      updates.visaStatus = '特定技能1号';
    }
  }

  // 8. Desired Industry (evaluate candidate text first; map nhà hàng/quán ăn/bếp strictly -> 外食業)
  if (/介護|特養|デイサービス|グループホーム|有料老人|shoninsha|初任者|viện de|viện có phí|kaigo|điều\s*dưỡng|hộ\s*lý/i.test(candText)) {
    updates.desiredIndustry = '介護';
  } else if (/外食|レストラン|居酒屋|キッチン|ngoại\s*thực|nhà\s*hàng|quán\s*ăn|ẩm\s*thực|phục\s*vụ\s*bàn|bếp/i.test(candText)) {
    updates.desiredIndustry = '外食業';
  } else if (/chế\s*biến\s*thực\s*phẩm|飲食料品|食品|bento|弁当|cơm\s*hộp|bánh\s*kẹo|thực\s*phẩm/i.test(candText)) {
    updates.desiredIndustry = '飲食料品製造業';
  } else if (/xây\s*dựng|建設|giàn\s*giáo|cốp\s*pha|sắt\s*thép\s*xây|nội\s*thất/i.test(candText)) {
    updates.desiredIndustry = '建設';
  } else if (/cơ\s*khí|gia\s*công\s*kim\s*loại|hàn|tiện|phay|cnc|素形材|産業機械|機械/i.test(candText)) {
    updates.desiredIndustry = '素形材・産業機械・電気電子情報関連製造業';
  } else if (/bảo\s*dưỡng\s*ô\s*tô|sửa\s*chữa\s*ô\s*tô|自動車整備/i.test(candText)) {
    updates.desiredIndustry = '自動車整備';
  } else if (/tài\s*xế|lái\s*xe|vận\s*tải|ドライバー|トラック|運送|運転/i.test(candText)) {
    updates.desiredIndustry = '自動車運送業';
  } else if (/khách\s*sạn|lưu\s*trú|ホテル|宿泊|フロント|buồng\s*phòng/i.test(candText)) {
    updates.desiredIndustry = '宿泊';
  } else if (/nông\s*nghiệp|農業|trồng\s*trọt|chăn\s*nuôi/i.test(candText)) {
    updates.desiredIndustry = '農業';
  } else if (/vệ\s*sinh\s*tòa\s*nhà|building\s*clean|ビルクリーニング|dọn\s*dẹp/i.test(candText)) {
    updates.desiredIndustry = 'ビルクリーニング';
  }

  // 9. Desired Areas (multi-prefecture)
  const extractedArea = extractPreferredAreaFromChat(candText);
  if (extractedArea) {
    updates.desiredArea = extractedArea;
  }

  // 10. Certificates / Passed Exams & Experience
  const examList: string[] = [];
  if (/shoninsha|初任者/i.test(candText)) examList.push('介護職員初任者研修修了');
  if (/jitsumusha|実務者/i.test(candText)) examList.push('介護福祉士実務者研修修了');
  if (/kaigo fukushishi|介護福祉士/i.test(candText)) examList.push('介護福祉士');
  if (
    /gino|技能評価試験|技能測定試験|特定技能[12]?号?(?:技能)?(?:評価|測定)?試験|chứng\s*chỉ\s*tokutei|đỗ\s*tay\s*nghề|senmon/i.test(
      candText
    )
  ) {
    examList.push('特定技能評価試験合格');
  }
  if (/bằng\s*lái|gplx|免許/i.test(candText)) examList.push('運転免許');
  if (updates.jlpt && /^N[1-5]$/.test(updates.jlpt)) {
    examList.push(`JLPT ${updates.jlpt}`);
  } else if (updates.jlpt === 'JFT') {
    examList.push('JFT-Basic');
  }
  if (examList.length > 0) {
    updates.passedExam = examList.join('・');
  }

  const expParts: string[] = [];
  if (/(?:chưa\s*(?:có\s*)?(?:kn|kinh\s*nghiệm)|chưa\s*từng\s*làm|mới\s*sang\s*chưa\s*làm|未経験)/i.test(candText)) {
    expParts.push('Chưa có kinh nghiệm');
  } else {
    const knMatch =
      candText.match(/(?:kn|kinh\s*nghiệm|経験)\s*(\d+(?:[.,]\d+)?)\s*(?:n\b|năm|年)/i) ||
      candText.match(/(\d+(?:[.,]\d+)?)\s*(?:năm|年)\s*(?:kn|kinh\s*nghiệm|経験)/i);
    if (knMatch) {
      expParts.push(`Kinh nghiệm ${knMatch[1].replace(',', '.')} năm`);
    }
  }
  if (/kaigo năm cuối|năm cuối/i.test(candText)) {
    expParts.push('Đang là Kaigo năm cuối');
  }
  const facilities: string[] = [];
  if (/viện\s*de\b|デイサービス|デイ\b/i.test(candText)) facilities.push('デイサービス');
  if (/viện\s*có\s*phí|有料老人ホーム|有料/i.test(candText)) facilities.push('有料老人ホーム');
  if (/tokuyou|特別養護老人ホーム|特養/i.test(candText)) facilities.push('特別養護老人ホーム');
  if (/group\s*home|グループホーム|GH\b/i.test(candText)) facilities.push('グループホーム');
  if (/rouken|老健/i.test(candText)) facilities.push('老健');
  if (facilities.length > 0) {
    expParts.push(`施設経験：${facilities.join('・')}`);
  }
  if (upcomingExamNote) {
    expParts.push(upcomingExamNote);
  }
  if (expParts.length > 0) {
    updates.japanExperience = expParts.join('／');
  }

  // 10b. Graduation date & Available startDate (current year = 2026)
  const CURRENT_CRM_YEAR = 2026;
  const gradNextYearM =
    candText.match(
      /(?:k[ìi]\s*)?tháng\s*(\d{1,2})\s*(?:[\/\-]\s*(20\d{2})|(?:năm\s*sau|sang\s*năm)|(năm\s*nay))?\s*(?:là\s*)?(?:tốt\s*nghiệp|ra\s*trường)/i
    ) ||
    candText.match(
      /(?:tốt\s*nghiệp|ra\s*trường)\s*(?:vào\s*)?(?:k[ìi]\s*)?tháng\s*(\d{1,2})\s*(?:[\/\-]\s*(20\d{2})|(?:năm\s*sau|sang\s*năm)|(năm\s*nay))?/i
    );
  const gradJpM = candText.match(/(20\d{2})年(\d{1,2})月(?:卒業|卒予定)/);
  if (gradNextYearM) {
    const gMonth = Math.min(12, Math.max(1, Number(gradNextYearM[1])));
    const explicitYr = gradNextYearM[2] ? Number(gradNextYearM[2]) : 0;
    const isNextYr = /năm\s*sau|sang\s*năm/i.test(gradNextYearM[0]);
    const isThisYr = /năm\s*nay/i.test(gradNextYearM[0]);
    const gYear = explicitYr
      ? explicitYr
      : isNextYr
      ? CURRENT_CRM_YEAR + 1
      : isThisYr
      ? CURRENT_CRM_YEAR
      : gMonth <= 11
      ? CURRENT_CRM_YEAR + 1
      : CURRENT_CRM_YEAR;
    updates.graduationDate = `${gYear}-${String(gMonth).padStart(2, '0')}`;
    const startMonth = (gMonth % 12) + 1;
    const startYear = gMonth === 12 ? gYear + 1 : gYear;
    updates.startDate = `${startYear}年${startMonth}月以降（sau tốt nghiệp tháng ${gMonth}/${gYear}）`;
  } else if (gradJpM) {
    const gYear = Number(gradJpM[1]);
    const gMonth = Math.min(12, Math.max(1, Number(gradJpM[2])));
    updates.graduationDate = `${gYear}-${String(gMonth).padStart(2, '0')}`;
    const startMonth = (gMonth % 12) + 1;
    const startYear = gMonth === 12 ? gYear + 1 : gYear;
    updates.startDate = `${startYear}年${startMonth}月以降（sau tốt nghiệp tháng ${gMonth}/${gYear}）`;
  } else if (/即日|làm\s*ngay|đi\s*làm\s*luôn/i.test(candText)) {
    updates.startDate = '即日可能';
  }

  // 11. Salary expectation & housing / night-shift lines
  const salaryLines = candidateLines.filter((l) =>
    /(?:lương|thực\s*nhận|cơ\s*bản|trừ\s*nhà|tiền\s*nhà|ký\s*túc|ca\s*đêm|yakin|jikyu|時給|手取り|基本給|総支給|寮|夜勤|\d+\s*(?:man|vạn|yên|円|万円))/i.test(
      l
    )
  );
  if (salaryLines.length > 0) {
    updates.salaryExpectation = Array.from(new Set(salaryLines)).slice(0, 3).join(' | ');
  }

  // 12. Conversion Goal / Need
  const goalLines = candidateLines.filter((l) =>
    /(?:mục\s*tiêu|nhu\s*cầu|tìm\s*(?:công\s*ty|đơn|việc|baito)|chuyển\s*(?:đổi|visa|việc|công\s*ty|sang\s*tokutei)|nguyện\s*vọng|希望|転職)/i.test(
      l
    )
  );
  if (goalLines.length > 0) {
    updates.conversionGoal = Array.from(new Set(goalLines)).slice(0, 3).join(' | ');
  } else if (isDhs) {
    updates.conversionGoal = 'Du học sinh tìm đơn chuyển đổi sang visa Tokutei Ginou';
  }

  // 13. Mendan / Interview (ignore hypothetical / question / example lines!)
  const scheduledLines = text
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(
      (l) =>
        l &&
        !/[?？]|ví\s*dụ|\bvd\s*:|例\s*[:：]|nếu\s+|chưa\s*chốt|khi\s*nào\s*tiện|dự\s*kiến\s*hỏi/i.test(
          l
        )
    )
    .join('\n');

  const mendanDateM = scheduledLines.match(
    /(?:mendan|面談)[^\n\d]{0,15}(\d{4}[年\/\-]\d{1,2}[月\/\-]\d{1,2}|\d{1,2}[月\/]\d{1,2}日?)(?:[^\n\d]{0,12}(\d{1,2}[:h]\d{2}))?/i
  ) || scheduledLines.match(
    /(\d{4}[年\/\-]\d{1,2}[月\/\-]\d{1,2}|\d{1,2}[月\/]\d{1,2}日?)\s*(?:lúc\s*(\d{1,2}[:h]\d{2})\s*)?(?:に|の)?\s*(?:面談|mendan)/i
  );
  if (mendanDateM) {
    const ts = parseFlexibleDateToTimestamp(mendanDateM[1], CURRENT_CRM_YEAR);
    if (ts > 0) {
      const timePart = (mendanDateM[2] || '').replace('h', ':');
      updates.mendanDate = timePart
        ? `${formatTimestampToIso(ts)} ${timePart.padStart(5, '0')}`
        : formatTimestampToIso(ts);
    }
  }

  const interviewDateM = scheduledLines.match(
    /(?:phỏng\s*vấn|pv|面接)[^\n\d]{0,15}(\d{4}[年\/\-]\d{1,2}[月\/\-]\d{1,2}|\d{1,2}[月\/]\d{1,2}日?)(?:[^\n\d]{0,12}(\d{1,2}[:h]\d{2}))?/i
  );
  if (interviewDateM) {
    const ts = parseFlexibleDateToTimestamp(interviewDateM[1], CURRENT_CRM_YEAR);
    if (ts > 0) {
      const timePart = (interviewDateM[2] || '').replace('h', ':');
      updates.interviewDate = timePart
        ? `${formatTimestampToIso(ts)} ${timePart.padStart(5, '0')}`
        : formatTimestampToIso(ts);
    }
  }

  if (/内定/.test(scheduledLines)) updates.progressStatus = '内定';
  else if (/2次面接/.test(scheduledLines)) updates.progressStatus = '2次面接予定';
  else if (/1次面接|面接予定|面接決ま/.test(scheduledLines)) updates.progressStatus = '1次面接予定';
  else if (/面談した|面談済|面談終了/.test(scheduledLines)) updates.progressStatus = '面談済み';
  else if (/面談調整|日程調整/.test(scheduledLines)) updates.progressStatus = '面談調整中';

  const compM = text.match(/([^\s、。,\n]+(?:株式会社|有限会社|医療法人|社会福祉法人))/);
  if (compM) updates.proposedCompany = compM[1];

  // Preserve 100% verbatim in negotiationLog, and build a concise <=4-5 bullet summary in memoLine
  const missingLegalNote =
    !updates.name && updates.fbName
      ? `[Chưa có họ tên khai sinh — mới ghi nhận tên FB: ${updates.fbName}]\n`
      : '';
  updates.negotiationLog = `[${todayIso} Ghi chú nguyên văn]\n${missingLegalNote}${text}`;

  const summaryBullets: string[] = [];
  const labelPerson = updates.name || updates.fbName;
  if (labelPerson) summaryBullets.push(`• Ứng viên: ${labelPerson}${updates.nationality ? ` (${updates.nationality})` : ''}`);
  if (updates.desiredIndustry || updates.desiredArea) {
    summaryBullets.push(
      `• Nguyện vọng: ${[updates.desiredIndustry, updates.desiredArea].filter(Boolean).join(' · ')}`
    );
  }
  if (updates.visaStatus || updates.jlpt) {
    summaryBullets.push(
      `• Visa / Tiếng Nhật: ${[updates.visaStatus, updates.jlpt].filter(Boolean).join(' · ')}${
        updates.visaExpiry ? ` (Hạn ${updates.visaExpiry})` : ''
      }`
    );
  }
  if (updates.conversionGoal) {
    summaryBullets.push(`• Mục tiêu: ${updates.conversionGoal.slice(0, 120)}`);
  } else if (candidateLines.length > 0) {
    summaryBullets.push(`• Ghi chú chính: ${candidateLines[0].slice(0, 120)}`);
  }
  if (updates.mendanDate || updates.interviewDate) {
    summaryBullets.push(
      `• Lịch hẹn: ${updates.interviewDate || updates.mendanDate}`
    );
  }
  updates.memoLine =
    summaryBullets.length > 0
      ? `【要点 ${todayIso}】\n${summaryBullets.slice(0, 5).join('\n')}`
      : `【要点 ${todayIso}】\n${text.slice(0, 180)}`;
  return updates;
};

export const buildSmartReplyMessages = (
  person: Partial<Candidate> = {},
  scenario: ReplyScenario = 'first_mendan',
  kyujins: KyujinCompany[] = []
) => {
  const rawName = String(person.fbName || person.name || '').trim();
  const nameVi = rawName || 'bạn';
  const nameJp = rawName ? `${rawName}さん` : 'ご担当者様';
  const nameEn = rawName || 'there';

  const missing = getMissingInfoForCandidate(person);
  const matches = findMatchingKyujinsForCandidate(person, kyujins);
  const topMatch = matches[0];
  const rawComp = getEffectiveProposedCompany(person) || topMatch?.company?.companyName || '';
  const proposedCompVi = rawComp || 'quý công ty';
  const proposedCompJp = rawComp || '企業様';
  const proposedCompEn = rawComp || 'the company';

  const mendanInfo = extractNormalizedMendanInfo(person);
  const rawDateText = mendanInfo.dateLabel || person.mendanDate || '';
  const dateTextVi = rawDateText || 'sắp tới';
  const dateTextJp = rawDateText || '近日中';
  const dateTextEn = rawDateText || 'soon';

  const missingVi =
    missing.length > 0
      ? `\nĐể chuẩn bị hồ sơ (CV / Rirekisho) và chọn đơn tốt nhất cho bạn, bạn cho mình xin thêm các thông tin sau nhé:\n${missing
          .map((m) => `・${m.vi}`)
          .join('\n')}`
      : '';
  const missingJp =
    missing.length > 0
      ? `\n履歴書作成および最適な求人ご案内のため、以下の点も教えていただけますでしょうか。\n${missing
          .map((m) => `・${m.ja}`)
          .join('\n')}`
      : '';
  const missingEn =
    missing.length > 0
      ? `\nTo prepare your resume (Rirekisho) and match the best job offers, could you please share:\n${missing
          .map((m) => `・${m.en}`)
          .join('\n')}`
      : '';

  if (scenario === 'missing_info_request') {
    const listVi = missing.length > 0 ? missing.map((m) => `・${m.vi}`).join('\n') : '・SĐT / Trình độ JLPT / Hạn thẻ ngoại kiều';
    const listJp = missing.length > 0 ? missing.map((m) => `・${m.ja}`).join('\n') : '・電話番号 / 日本語レベル / 在留期限';
    const listEn = missing.length > 0 ? missing.map((m) => `・${m.en}`).join('\n') : '・Phone / JLPT / Residence card expiry';
    return {
      vi: `Chào ${nameVi}! Để mình lọc chính xác 3 đơn hàng lương tốt và hỗ trợ làm Rirekisho gửi công ty cho bạn nhanh nhất, bạn bổ sung giúp mình mấy thông tin còn thiếu sau nhé:\n${listVi}\n\nBạn nhắn ngắn gọn hoặc chụp ảnh mặt trước thẻ ngoại kiều (che số thẻ nếu muốn) gửi mình cũng được nha!`,
      jp: `${nameJp}、こんにちは！ご希望に合う求人のご提案と履歴書作成を進めるため、不足している以下の情報をお送りいただけますでしょうか。\n${listJp}\n\nメッセージまたは在留カードのお写真でお知らせいただければすぐに対応いたします！`,
      en: `Hello ${nameEn}! To match the best job openings and prepare your resume quickly, could you please share the missing details below:\n${listEn}`,
    };
  }

  if (scenario === 'post_interview_followup') {
    return {
      vi: `Chào ${nameVi}! Buổi phỏng vấn với ${proposedCompVi} (${dateTextVi}) vừa rồi thế nào bạn ơi?\nBạn thấy không khí công ty, nội dung công việc và mức lương phía công ty chia sẻ có phù hợp với kỳ vọng của bạn không? Có điều gì băn khoăn cứ nhắn thật lòng cho mình để mình trao đổi thêm với công ty hoặc tính phương án tốt nhất cho bạn nhé!`,
      jp: `${nameJp}、お疲れ様でした！${proposedCompJp}との面接（${dateTextJp}）はいかがでしたでしょうか？\n業務内容や給与条件・職場の雰囲気など、ご感想や気になる点がございましたら遠慮なくお聞かせください。結果連絡や次のステップに向けてしっかりサポートいたします！`,
      en: `Hello ${nameEn}! How did your interview with ${proposedCompEn} (${dateTextEn}) go? Please let me know your thoughts on the role, salary, and workplace so I can support your next steps!`,
    };
  }

  if (scenario === 'recommend_job') {
    const jobBlockVi =
      matches.length > 0
        ? matches
            .slice(0, 3)
            .map(({ company, slot, score }, i) => `--- 【Đơn gợi ý #${i + 1} | Độ khớp ${score}%】 ---\n${buildAnonymousJobPostText(company, slot, 'vi')}`)
            .join('\n\n')
        : `Ngành ${person.desiredIndustry || 'Tokutei'} lương tốt, hỗ trợ hồ sơ chuyển việc trọn gói.`;

    const jobBlockJp =
      matches.length > 0
        ? matches
            .slice(0, 3)
            .map(({ company, slot, score }, i) => `--- 【おすすめ求人 #${i + 1} | マッチ度 ${score}%】 ---\n${buildAnonymousJobPostText(company, slot, 'jp')}`)
            .join('\n\n')
        : `${person.desiredIndustry || '特定技能'}の好条件求人がございます。`;

    const jobBlockEn =
      matches.length > 0
        ? matches
            .slice(0, 3)
            .map(
              ({ slot, score }, i) =>
                `[Job ${i + 1} (${score}% match)] ${slot.employmentType || 'Tokutei'} / ${slot.industry} / Area: ${formatPublicArea(
                  slot.area
                )} / Salary: ${slot.salary} / Housing: ${slot.housingSupport}`
            )
            .join('\n')
        : `We have great Tokutei Ginou openings in ${person.desiredIndustry || 'your field'}.`;

    return {
      vi: `Chào ${nameVi}! Dựa theo hồ sơ của bạn, mình gửi 3 đơn phù hợp nhất hiện tại (ẩn danh tên công ty trước khi phỏng vấn):\n\n${jobBlockVi}\n${missingVi}\nNếu bạn ưng đơn số mấy, nhắn lại cho mình để mình gửi lịch phỏng vấn sớm nhất nhé!`,
      jp: `${nameJp}、こんにちは！ご希望条件に合う求人3件をご案内いたします（選考前のため企業名は非公開です）。\n\n${jobBlockJp}\n${missingJp}\n気になる求人がございましたら番号をお知らせください！`,
      en: `Hello ${nameEn}! Here are the top matching openings for your profile:\n\n${jobBlockEn}\n${missingEn}\nPlease let me know which one interests you!`,
    };
  }

  if (scenario === 'remind_interview') {
    return {
      vi: `Chào ${nameVi}! Mình nhắn nhắc bạn lịch hẹn Mendan / Phỏng vấn sắp tới:\n・Thời gian: ${dateTextVi}\n・Nội dung / Công ty: ${proposedCompVi}\n\nBạn nhớ kiểm tra đường truyền mạng, trang phục gọn gàng và vào trước giờ hẹn 5–10 phút nhé. Nếu cần hỗ trợ luyện tập câu hỏi phỏng vấn thì nhắn mình ngay nha!`,
      jp: `${nameJp}、こんにちは！面談・面接のリマインド連絡です。\n・日時：${dateTextJp}\n・対象企業／内容：${proposedCompJp}\n\n当日は開始5〜10分前にご準備をお願いいたします。事前練習やご不明点があればいつでもご連絡ください！`,
      en: `Hello ${nameEn}! This is a friendly reminder for your upcoming Mendan / Interview:\n・Date/Time: ${dateTextEn}\n・Company / Details: ${proposedCompEn}\n\nPlease be ready 5–10 minutes early in neat attire. Let me know if you need any interview prep!`,
    };
  }

  if (scenario === 'followup_silent') {
    return {
      vi: `Chào ${nameVi}! Lâu rồi chưa thấy bạn phản hồi, không biết tình hình công việc và tìm việc Tokutei của bạn dạo này thế nào rồi?\nHiện bên mình mới về thêm nhiều đơn lương tốt và có thưởng. Nếu bạn vẫn đang tìm việc hoặc muốn tham khảo thêm đơn mới thì nhắn lại cho mình nhé!`,
      jp: `${nameJp}、こんにちは！その後お仕事探しの状況はいかがでしょうか？\n現在、賞与あり・好条件の特定技能求人も新しく入っております。まだ転職・就職をご検討中でしたら、ぜひお気軽にご返信ください！`,
      en: `Hello ${nameEn}! How has your job search been going lately?\nWe have received several new Tokutei Ginou openings with good salary and bonuses. If you are still looking, feel free to reply anytime!`,
    };
  }

  return {
    vi: `Chào ${nameVi}! Cảm ơn bạn đã liên hệ tư vấn việc làm Tokutei Ginou.\nBên mình hỗ trợ miễn phí 100% từ khâu làm hồ sơ (CV / Rirekisho) đến luyện phỏng vấn và xin visa.\n${missingVi}\n\nHôm nay hoặc ngày mai khung giờ nào bạn tiện nghe điện thoại/gọi video Mendan khoảng 15 phút để mình tư vấn đơn phù hợp nhất cho bạn?`,
    jp: `${nameJp}、こんにちは！特定技能のお仕事相談にお問い合わせいただきありがとうございます。\n履歴書作成から面接対策・ビザ申請まで完全無料でサポートしております。\n${missingJp}\n\n本日または明日で、15分ほどお電話・オンライン面談が可能なお時間帯はございますでしょうか？`,
    en: `Hello ${nameEn}! Thank you for reaching out regarding Tokutei Ginou job support.\nWe provide 100% free support from resume preparation to interview coaching and visa procedures.\n${missingEn}\n\nWhat time today or tomorrow would be convenient for a quick 15-minute Mendan call?`,
  };
};

// ============================================================================
// RIREKISHO STANDALONE HTML BUILDER
// ============================================================================
const escapeHtml = (str = ''): string =>
  String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

export const buildRirekishoStandaloneHtml = (draft: RirekishoDraft): string => {
  const eduRowsHtml = (draft.educationRows || [])
    .map(
      (r) =>
        `<tr><td class="c">${escapeHtml(r.year)}</td><td class="c">${escapeHtml(
          r.month
        )}</td><td>${escapeHtml(r.content)}</td></tr>`
    )
    .join('');
  const workRowsHtml = (draft.workRows || [])
    .map(
      (r) =>
        `<tr><td class="c">${escapeHtml(r.year)}</td><td class="c">${escapeHtml(
          r.month
        )}</td><td>${escapeHtml(r.content)}</td></tr>`
    )
    .join('');
  const licRowsHtml = (draft.licenseRows || [])
    .map(
      (r) =>
        `<tr><td class="c">${escapeHtml(r.year)}</td><td class="c">${escapeHtml(
          r.month
        )}</td><td>${escapeHtml(r.content)}</td></tr>`
    )
    .join('');

  return `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8" />
<title>履歴書_${escapeHtml(draft.fullName || '応募者')}</title>
<style>
  @page { size: A4 landscape; margin: 10mm; }
  * { box-sizing: border-box; }
  body {
    font-family: "Hiragino Kaku Gothic ProN", "Yu Gothic", "Meiryo", sans-serif;
    font-size: 10.5pt;
    color: #111;
    margin: 0;
    padding: 12px;
    background: #fff;
  }
  .sheet {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 16px;
    width: 100%;
    max-width: 1120px;
    margin: 0 auto;
  }
  .col {
    border: 2px solid #111;
    padding: 10px 12px;
    display: flex;
    flex-direction: column;
    gap: 8px;
  }
  .header-row {
    display: flex;
    justify-content: space-between;
    align-items: baseline;
    border-bottom: 1.5px solid #111;
    padding-bottom: 4px;
  }
  .title { font-size: 18pt; font-weight: bold; letter-spacing: 0.35em; }
  .date { font-size: 9.5pt; }
  table { width: 100%; border-collapse: collapse; table-layout: fixed; }
  th, td {
    border: 1px solid #333;
    padding: 4px 6px;
    vertical-align: middle;
    word-break: break-word;
    font-size: 9.5pt;
  }
  th { background: #f3f4f6; font-weight: bold; text-align: center; }
  .subhead { background: #f8fafc; font-weight: bold; text-align: center; }
  .c { text-align: center; }
  .name-cell { font-size: 13pt; font-weight: bold; padding: 6px; }
  .box {
    border: 1px solid #333;
    padding: 6px 8px;
    font-size: 9.5pt;
  }
  .box-title { font-weight: bold; margin-bottom: 4px; font-size: 9pt; }
  .pre { white-space: pre-wrap; line-height: 1.45; min-height: 44px; }
  .photo-box {
    width: 84px;
    height: 112px;
    border: 1px dashed #666;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 8.5pt;
    color: #666;
    overflow: hidden;
    flex-shrink: 0;
    background: #fafafa;
  }
  .photo-box img {
    width: 100%;
    height: 100%;
    object-fit: cover;
  }
  @media print {
    body { padding: 0; }
    .no-print { display: none !important; }
  }
</style>
</head>
<body>
  <div class="sheet">
    <div class="col">
      <div class="header-row">
        <span class="title">履歴書</span>
        <span class="date">${escapeHtml(draft.dateStr)}</span>
      </div>
      <div style="display:flex; gap:8px; align-items:stretch;">
        <table style="flex:1;">
          <tr>
            <th style="width:26%">フリガナ</th>
            <td>${escapeHtml(draft.furigana)}</td>
          </tr>
          <tr>
            <th>氏名</th>
            <td class="name-cell">${escapeHtml(draft.fullName)}</td>
          </tr>
          <tr>
            <th>生年月日</th>
            <td>${escapeHtml(draft.birthDate)}（満 ${escapeHtml(draft.age)} 歳） ／ 性別：${escapeHtml(
    draft.gender
  )}</td>
          </tr>
        </table>
        <div class="photo-box">
          ${
            draft.photoDataUrl && draft.photoDataUrl.startsWith('data:image/')
              ? `<img src="${draft.photoDataUrl}" alt="証明写真" />`
              : '証明写真'
          }
        </div>
      </div>
      <table>
        <tr>
          <th style="width:22%">現住所</th>
          <td>〒${escapeHtml(draft.postalCode)}<br/>${escapeHtml(draft.address)}</td>
        </tr>
        <tr>
          <th>電話 / Email</th>
          <td>TEL：${escapeHtml(draft.phone)} ／ Mail：${escapeHtml(draft.email)}</td>
        </tr>
      </table>

      <table>
        <tr>
          <th>国籍</th>
          <th>在留期限</th>
          <th>在留資格</th>
        </tr>
        <tr>
          <td class="c">${escapeHtml(draft.nationality)}</td>
          <td class="c">${escapeHtml(draft.visaExpiry)}</td>
          <td class="c">${escapeHtml(draft.visaStatus)}</td>
        </tr>
      </table>

      <table>
        <tr>
          <th style="width:14%">年</th>
          <th style="width:10%">月</th>
          <th>学歴・職歴</th>
        </tr>
        <tr><td colspan="3" class="subhead">学 歴</td></tr>
        ${eduRowsHtml}
        <tr><td colspan="3" class="subhead">職 歴</td></tr>
        ${workRowsHtml}
      </table>

      <div class="box" style="flex:1;">
        <div class="box-title">本人希望記入欄（給料・職種・勤務時間・勤務地・その他についての希望）</div>
        <div class="pre">${escapeHtml(draft.personalRequests)}</div>
      </div>
    </div>

    <div class="col">
      <table>
        <tr>
          <th style="width:14%">年</th>
          <th style="width:10%">月</th>
          <th>免許・資格</th>
        </tr>
        ${licRowsHtml}
      </table>

      <table>
        <tr><th colspan="2" style="text-align:left;">話せる言語</th></tr>
        <tr>
          <th style="width:28%">母国語</th>
          <td>${escapeHtml(draft.nativeLanguage)}</td>
        </tr>
        <tr>
          <th>使用可能言語</th>
          <td>${escapeHtml(draft.availableLanguages)}</td>
        </tr>
      </table>

      <table>
        <tr><th colspan="2" style="text-align:left;">日本語能力CA評価</th></tr>
        <tr>
          <th style="width:28%">話す</th>
          <td>${escapeHtml(draft.caSpeaking)}</td>
        </tr>
        <tr>
          <th>聞く</th>
          <td>${escapeHtml(draft.caListening)}</td>
        </tr>
      </table>

      <div class="box" style="flex:1;">
        <div class="box-title">志望の動機、自己PRなど</div>
        <div style="font-weight:bold; font-size:9pt; margin-top:2px;">＜志望動機＞</div>
        <div class="pre">${escapeHtml(draft.motivation)}</div>
        <div style="font-weight:bold; font-size:9pt; margin-top:6px;">＜自己PR＞</div>
        <div class="pre">${escapeHtml(draft.selfPr)}</div>
      </div>

      <table>
        <tr>
          <th>アレルギー</th>
          <th>扶養家族数(配偶者を除く)</th>
          <th>配偶者 / 配偶者の扶養義務</th>
        </tr>
        <tr>
          <td class="c">${escapeHtml(draft.allergy)}</td>
          <td class="c">${escapeHtml(draft.dependentsCount)}</td>
          <td class="c">${escapeHtml(draft.spouse)} / ${escapeHtml(draft.spouseSupport)}</td>
        </tr>
      </table>
    </div>
  </div>
</body>
</html>`;
};

export const exportRirekishoToExcelFile = (draft: RirekishoDraft): string => {
  const safeName = String(draft.fullName || '応募者')
    .replace(/[\\/:*?"<>|\s]+/g, '_')
    .trim();
  const dateTag = new Date().toISOString().slice(0, 10);
  const fileName = `履歴書_${safeName}_${dateTag}.xlsx`;

  const aoa: Array<Array<string>> = [
    ['履 歴 書', '', draft.dateStr || ''],
    ['フリガナ', draft.furigana || '', ''],
    ['氏名', draft.fullName || '', ''],
    [
      '生年月日 / 年齢 / 性別',
      `${draft.birthDate || ''}（満 ${draft.age || ''} 歳）`,
      draft.gender || '',
    ],
    ['郵便番号 / 現住所', `〒${draft.postalCode || ''}`, draft.address || ''],
    ['電話番号 / Email', draft.phone || '', draft.email || ''],
    ['国籍 / 在留期限 / 在留資格', draft.nationality || '', `${draft.visaExpiry || ''} / ${draft.visaStatus || ''}`],
    [],
    ['年', '月', '学 歴'],
    ...(draft.educationRows || []).map((r) => [r.year || '', r.month || '', r.content || '']),
    ['年', '月', '職 歴'],
    ...(draft.workRows || []).map((r) => [r.year || '', r.month || '', r.content || '']),
    [],
    ['年', '月', '免許・資格'],
    ...(draft.licenseRows || []).map((r) => [r.year || '', r.month || '', r.content || '']),
    [],
    ['話せる言語（母国語）', draft.nativeLanguage || '', ''],
    ['使用可能言語', draft.availableLanguages || '', ''],
    ['日本語能力CA評価（話す）', draft.caSpeaking || '', ''],
    ['日本語能力CA評価（聞く）', draft.caListening || '', ''],
    [],
    ['志望動機', draft.motivation || '', ''],
    ['自己PR', draft.selfPr || '', ''],
    ['本人希望記入欄', draft.personalRequests || '', ''],
    [],
    ['アレルギー', draft.allergy || '', ''],
    ['扶養家族数(配偶者を除く)', draft.dependentsCount || '', ''],
    ['配偶者 / 配偶者の扶養義務', `${draft.spouse || ''} / ${draft.spouseSupport || ''}`, ''],
  ];

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{ wch: 24 }, { wch: 24 }, { wch: 52 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '履歴書');
  XLSX.writeFile(wb, fileName);
  return fileName;
};
