import * as XLSX from 'xlsx';
import type {
  Candidate,
  ResignationPaperworkStatus,
  ResignationProtectedField,
  ResignationRecord,
} from '../types/crm';

export const RESIGNATION_PROTECTED_FIELDS: ResignationProtectedField[] = [
  'nationality',
  'resignationDate',
  'paperworkDeadline',
];

export function normalizeResignationNationality(raw: unknown): string {
  const s = String(raw ?? '').trim();
  if (!s) return '国籍不明';
  const lower = s.toLowerCase();

  if (
    /^(?:国籍不明|不明|未設定|未確認|chưa\s*rõ(?:\s*quốc\s*tịch)?|unknown|null|undefined|n\/a|-+)$/i.test(
      lower
    )
  ) {
    return '国籍不明';
  }
  if (/^(?:ベトナム|viet\s*nam|việt\s*nam|vietnam|vn|vnm)$/i.test(lower) || /ベトナム|viet\s*nam|việt\s*nam|vietnam/i.test(lower)) {
    return 'ベトナム';
  }
  if (/^(?:インドネシア|indonesia|indo|idn)$/i.test(lower) || /インドネシア|indonesia/i.test(lower)) {
    return 'インドネシア';
  }
  if (/^(?:ミャンマー|myanmar|burma|mmr|miến\s*điện|mien\s*dien)$/i.test(lower) || /ミャンマー|myanmar|miến\s*điện/i.test(lower)) {
    return 'ミャンマー';
  }
  if (/^(?:フィリピン|philippines?|filipino|phl)$/i.test(lower) || /フィリピン|philippines?|filipino/i.test(lower)) {
    return 'フィリピン';
  }
  if (/^(?:カンボジア|cambodia|campuchia|khmer|khm)$/i.test(lower) || /カンボジア|cambodia|campuchia/i.test(lower)) {
    return 'カンボジア';
  }
  if (/^(?:ネパール|nepal|nepali|npl)$/i.test(lower) || /ネパール|nepal/i.test(lower)) {
    return 'ネパール';
  }
  if (/^(?:その他|other|khác|khac)$/i.test(lower)) {
    return 'その他';
  }
  return s;
}

const VIETNAMESE_DIACRITIC_RX =
  /[ăâđêôơưàáảãạằắẳẵặầấẩẫậèéẻẽẹềếểễệìíỉĩịòóỏõọồốổỗộờớởỡợùúủũụừứửữựỳýỷỹỵ]/i;

export function suggestNationalityFromName(fullName?: string): string | null {
  const raw = String(fullName || '').trim();
  if (!raw || /^chưa rõ tên/i.test(raw)) return null;

  if (VIETNAMESE_DIACRITIC_RX.test(raw)) {
    return 'ベトナム';
  }

  const tokens = raw
    .toUpperCase()
    .replace(/[^A-Z\s]/g, ' ')
    .split(/\s+/)
    .filter(Boolean);
  if (tokens.length === 0) return null;

  const hasToken = (set: Set<string>) => tokens.some((t) => set.has(t));

  const vietnameseTokens = new Set([
    'NGUYEN',
    'TRAN',
    'LE',
    'PHAM',
    'HOANG',
    'HUYNH',
    'PHAN',
    'VU',
    'VO',
    'DANG',
    'BUI',
    'NGO',
    'DUONG',
    'DINH',
    'TRUONG',
    'DOAN',
    'TRINH',
    'PHUONG',
    'THUY',
    'HUONG',
    'QUYNH',
    'TUYET',
    'NHUNG',
    'KHANH',
    'THANH',
    'THI',
    'VAN',
    'ANH',
    'HUNG',
    'DUNG',
    'MINH',
    'TUAN',
    'QUANG',
    'HIEU',
    'CUONG',
    'HAO',
    'NGOC',
    'TRANG',
    'LINH',
  ]);

  if (
    tokens[0] === 'NGUYEN' ||
    tokens[0] === 'TRAN' ||
    tokens[0] === 'PHAM' ||
    tokens[0] === 'HOANG' ||
    tokens[0] === 'HUYNH' ||
    tokens.filter((t) => vietnameseTokens.has(t)).length >= 2
  ) {
    return 'ベトナム';
  }

  const myanmarTokens = new Set([
    'AUNG',
    'KYAW',
    'THANT',
    'PHOME',
    'PHYO',
    'ZAW',
    'MYO',
    'HTET',
    'MAUNG',
    'HLA',
    'SWE',
    'NYEIN',
    'THURA',
    'WIN',
    'TUN',
    'SOE',
    'HTUN',
    'TIN',
    'MON',
    'OO',
    'KO',
    'HEIN',
    'WAI',
    'PAING',
    'KHANT',
    'THEIN',
    'NAING',
    'MYINT',
    'SAN',
    'THIDA',
    'YEE',
    'EI',
    'SU',
    'HNIN',
  ]);
  if (
    tokens.filter((t) => myanmarTokens.has(t)).length >= 2 ||
    hasToken(new Set(['AUNG', 'KYAW', 'THANT', 'PHOME', 'PHYO', 'HTET', 'MAUNG', 'NYEIN', 'THURA', 'KHANT']))
  ) {
    return 'ミャンマー';
  }

  const indonesianTokens = new Set([
    'PUTRA',
    'PUTRI',
    'DEWI',
    'SARI',
    'PRATAMA',
    'SAPUTRA',
    'SAPUTRI',
    'WATI',
    'ARDHANA',
    'NOVIA',
    'SUSANTO',
    'SETIAWAN',
    'HIDAYAT',
    'NUGROHO',
    'WIBOWO',
    'KUSUMA',
    'RAHMAWATI',
    'LESTARI',
    'MUHAMMAD',
    'AHMAD',
    'RIZKI',
    'RIZKY',
    'WAHYU',
    'AGUNG',
    'BUDI',
    'FITRI',
    'INDAH',
    'YULIANA',
    'PURNAMA',
    'SYAHPUTRA',
  ]);
  if (hasToken(indonesianTokens)) {
    return 'インドネシア';
  }

  const nepaliTokens = new Set([
    'TAMANG',
    'GURUNG',
    'THAPA',
    'MAGAR',
    'LIMBU',
    'SHRESTHA',
    'ADHIKARI',
    'KHADKA',
    'POUDEL',
    'PAUDEL',
    'BHANDARI',
    'KARKI',
    'BASNET',
    'LAMA',
    'GHALE',
    'BAHADUR',
    'BIKRAM',
    'KUMARI',
    'PRASAD',
    'SUBEDI',
    'ACHARYA',
  ]);
  if (hasToken(nepaliTokens)) {
    return 'ネパール';
  }

  const filipinoTokens = new Set([
    'DELA',
    'CRUZ',
    'SANTOS',
    'REYES',
    'MENDOZA',
    'BAUTISTA',
    'VILLANUEVA',
    'RAMOS',
    'AQUINO',
    'CASTILLO',
    'MAGSAYSAY',
    'DIZON',
    'SORIANO',
    'OCAMPO',
    'GONZALES',
    'HERNANDEZ',
  ]);
  if (
    tokens.filter((t) => filipinoTokens.has(t)).length >= 1 &&
    !hasToken(vietnameseTokens)
  ) {
    return 'フィリピン';
  }

  const cambodianTokens = new Set([
    'SOK',
    'SOPHEA',
    'PHEAKDEY',
    'VANNA',
    'CHANN',
    'SREY',
    'RATANA',
    'DARA',
    'BORA',
    'NHEM',
    'CHEA',
    'OUK',
    'CHANTHA',
    'SOVANN',
    'VANN',
    'RITH',
    'MAKARA',
    'KIMSAN',
  ]);
  if (hasToken(cambodianTokens)) {
    return 'カンボジア';
  }

  return null;
}

export function applyManualProtectionOnResignationSave(
  prevRecord: ResignationRecord | null,
  nextRecord: ResignationRecord,
  explicitFields: ResignationProtectedField[] = []
): ResignationRecord {
  const protectedSet = new Set<ResignationProtectedField>([
    ...(prevRecord?.manualProtectedFields || []),
    ...(nextRecord.manualProtectedFields || []),
    ...explicitFields,
  ]);

  if (prevRecord) {
    if ((prevRecord.nationality || '') !== (nextRecord.nationality || '')) {
      protectedSet.add('nationality');
    }
    if ((prevRecord.resignationDate || '') !== (nextRecord.resignationDate || '')) {
      protectedSet.add('resignationDate');
    }
    if ((prevRecord.paperworkDeadline || '') !== (nextRecord.paperworkDeadline || '')) {
      protectedSet.add('paperworkDeadline');
    }
  } else {
    if (nextRecord.nationality && nextRecord.nationality !== '国籍不明') {
      protectedSet.add('nationality');
    }
    if (nextRecord.resignationDate) {
      protectedSet.add('resignationDate');
    }
    if (nextRecord.paperworkDeadline) {
      protectedSet.add('paperworkDeadline');
    }
  }

  const manualProtectedFields = Array.from(protectedSet);
  return {
    ...nextRecord,
    nationality: normalizeResignationNationality(nextRecord.nationality),
    manualProtectedFields,
    gasProtected: Boolean(nextRecord.gasProtected || manualProtectedFields.length > 0),
  };
}

export function mergeResignationWithGasProtection(
  existing: ResignationRecord | undefined,
  incomingRaw: Partial<ResignationRecord> & Record<string, unknown>,
  gasWriteLocked: boolean
): ResignationRecord {
  const incomingRawNat =
    incomingRaw.nationality ??
    incomingRaw['国籍'] ??
    incomingRaw['Quốc tịch'] ??
    incomingRaw['Nationality'] ??
    incomingRaw['country'];
  const incomingNationality = normalizeResignationNationality(incomingRawNat);
  const incomingResigDate = normalizeFlexibleDateToYmd(
    incomingRaw.resignationDate ?? incomingRaw['退職日']
  );
  const incomingDeadline = normalizeFlexibleDateToYmd(
    incomingRaw.paperworkDeadline ?? incomingRaw['書類提出期限'] ?? incomingRaw['提出期限']
  );

  if (!existing) {
    return {
      id: String(incomingRaw.id || `resig_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`),
      fullName: String(incomingRaw.fullName || incomingRaw['氏名'] || 'Chưa rõ tên').trim(),
      nationality: incomingNationality,
      previousCompany: String(
        incomingRaw.previousCompany || incomingRaw['前職'] || 'Chưa rõ công ty cũ'
      ).trim(),
      resignationDate: incomingResigDate,
      paperworkDeadline: incomingDeadline,
      paperworkStatus: normalizePaperworkStatus(incomingRaw.paperworkStatus),
      handoverNotes: String(incomingRaw.handoverNotes || '').trim() || undefined,
      newJobNeed: String(incomingRaw.newJobNeed || '').trim() || undefined,
      phone: String(incomingRaw.phone || '').trim() || undefined,
      fbName: String(incomingRaw.fbName || '').trim() || undefined,
      fbProfileUrl: String(incomingRaw.fbProfileUrl || '').trim() || undefined,
      industry: String(incomingRaw.industry || '').trim() || undefined,
      jlpt: String(incomingRaw.jlpt || '').trim() || undefined,
      convertedCandidateId: incomingRaw.convertedCandidateId
        ? String(incomingRaw.convertedCandidateId)
        : undefined,
      googleCalendarEventId: incomingRaw.googleCalendarEventId
        ? String(incomingRaw.googleCalendarEventId)
        : undefined,
      googleCalendarEventLink: incomingRaw.googleCalendarEventLink
        ? String(incomingRaw.googleCalendarEventLink)
        : undefined,
      manualProtectedFields: Array.isArray(incomingRaw.manualProtectedFields)
        ? (incomingRaw.manualProtectedFields as ResignationProtectedField[])
        : undefined,
      gasProtected: Boolean(incomingRaw.gasProtected),
      updatedAt: incomingRaw.updatedAt ? String(incomingRaw.updatedAt) : new Date().toISOString(),
    };
  }

  const protectedFields = new Set<ResignationProtectedField>(
    existing.manualProtectedFields || []
  );
  const isProtected = (field: ResignationProtectedField) =>
    gasWriteLocked || Boolean(existing.gasProtected) || protectedFields.has(field);

  const nextNationality = isProtected('nationality')
    ? normalizeResignationNationality(existing.nationality)
    : incomingNationality !== '国籍不明'
    ? incomingNationality
    : normalizeResignationNationality(existing.nationality);

  const nextResigDate = isProtected('resignationDate')
    ? existing.resignationDate
    : incomingResigDate || existing.resignationDate;

  const nextDeadline = isProtected('paperworkDeadline')
    ? existing.paperworkDeadline
    : incomingDeadline || existing.paperworkDeadline;

  return {
    ...existing,
    id: existing.id,
    fullName: String(incomingRaw.fullName || existing.fullName).trim(),
    previousCompany: String(incomingRaw.previousCompany || existing.previousCompany).trim(),
    nationality: nextNationality,
    resignationDate: nextResigDate,
    paperworkDeadline: nextDeadline,
    paperworkStatus: incomingRaw.paperworkStatus
      ? normalizePaperworkStatus(incomingRaw.paperworkStatus)
      : existing.paperworkStatus,
    handoverNotes: String(incomingRaw.handoverNotes || existing.handoverNotes || '').trim() || undefined,
    newJobNeed: String(incomingRaw.newJobNeed || existing.newJobNeed || '').trim() || undefined,
    phone: String(incomingRaw.phone || existing.phone || '').trim() || undefined,
    fbName: String(incomingRaw.fbName || existing.fbName || '').trim() || undefined,
    fbProfileUrl: String(incomingRaw.fbProfileUrl || existing.fbProfileUrl || '').trim() || undefined,
    industry: String(incomingRaw.industry || existing.industry || '').trim() || undefined,
    jlpt: String(incomingRaw.jlpt || existing.jlpt || '').trim() || undefined,
    convertedCandidateId:
      (incomingRaw.convertedCandidateId ? String(incomingRaw.convertedCandidateId) : undefined) ||
      existing.convertedCandidateId,
    googleCalendarEventId:
      (incomingRaw.googleCalendarEventId ? String(incomingRaw.googleCalendarEventId) : undefined) ||
      existing.googleCalendarEventId,
    googleCalendarEventLink:
      (incomingRaw.googleCalendarEventLink ? String(incomingRaw.googleCalendarEventLink) : undefined) ||
      existing.googleCalendarEventLink,
    manualProtectedFields:
      protectedFields.size > 0 ? Array.from(protectedFields) : existing.manualProtectedFields,
    gasProtected: Boolean(existing.gasProtected || protectedFields.size > 0),
    updatedAt: existing.updatedAt || new Date().toISOString(),
  };
}

export function getJstTodayYmd(): string {
  return new Date(Date.now() + 9 * 60 * 60 * 1000).toISOString().slice(0, 10);
}

export function normalizeFlexibleDateToYmd(raw: unknown): string {
  if (raw === null || raw === undefined || raw === '') return '';
  if (typeof raw === 'number' && raw > 20000 && raw < 80000) {
    // Excel serial date
    const utcDays = Math.floor(raw - 25569);
    const dt = new Date(utcDays * 86400 * 1000);
    if (!Number.isNaN(dt.getTime())) {
      return dt.toISOString().slice(0, 10);
    }
  }
  const text = String(raw).trim();
  if (!text) return '';

  // YYYY-MM-DD or YYYY/MM/DD or YYYY年MM月DD日
  const ymd = text.match(/(\d{4})\s*[年\/.\-]\s*(\d{1,2})\s*[月\/.\-]\s*(\d{1,2})/);
  if (ymd) {
    const y = ymd[1];
    const m = String(Number(ymd[2])).padStart(2, '0');
    const d = String(Number(ymd[3])).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  // DD/MM/YYYY or DD-MM-YYYY
  const dmy = text.match(/^(\d{1,2})\s*[\/.\-]\s*(\d{1,2})\s*[\/.\-]\s*(\d{4})$/);
  if (dmy) {
    const d = String(Number(dmy[1])).padStart(2, '0');
    const m = String(Number(dmy[2])).padStart(2, '0');
    const y = dmy[3];
    return `${y}-${m}-${d}`;
  }

  return '';
}

export function normalizePaperworkStatus(raw: unknown): ResignationPaperworkStatus {
  const s = String(raw || '').trim().toLowerCase();
  if (!s) return 'Chưa làm';
  if (/hoàn\s*tất|done|complete|完了|済/.test(s)) return 'Hoàn tất';
  if (/đã\s*gửi|sent|提出済|送付/.test(s)) return 'Đã gửi công ty';
  if (/đang\s*thu\s*thập|collecting|đang\s*làm|in\s*progress|準備中|収集中/.test(s))
    return 'Đang thu thập';
  return 'Chưa làm';
}

export interface ResignationDeadlineAssessment {
  daysLeft: number | null;
  urgency: 'done' | 'overdue' | 'urgent3' | 'warn7' | 'normal';
  label: string;
}

export function assessResignationDeadline(
  deadlineYmd?: string,
  status?: string
): ResignationDeadlineAssessment {
  if (status === 'Hoàn tất') {
    return { daysLeft: null, urgency: 'done', label: 'Hoàn tất thủ tục' };
  }
  const ymd = normalizeFlexibleDateToYmd(deadlineYmd);
  if (!ymd) {
    return { daysLeft: null, urgency: 'normal', label: 'Chưa đặt hạn nộp' };
  }

  const todayYmd = getJstTodayYmd();
  const [ty, tm, td] = todayYmd.split('-').map(Number);
  const [dy, dm, dd] = ymd.split('-').map(Number);
  const todayUtc = Date.UTC(ty, tm - 1, td);
  const deadUtc = Date.UTC(dy, dm - 1, dd);
  const daysLeft = Math.round((deadUtc - todayUtc) / 86400000);

  if (daysLeft < 0) {
    return {
      daysLeft,
      urgency: 'overdue',
      label: `Quá hạn ${Math.abs(daysLeft)} ngày (${ymd})`,
    };
  }
  if (daysLeft === 0) {
    return { daysLeft: 0, urgency: 'urgent3', label: `Hạn hôm nay (${ymd})` };
  }
  if (daysLeft <= 3) {
    return {
      daysLeft,
      urgency: 'urgent3',
      label: `Còn ${daysLeft} ngày (${ymd})`,
    };
  }
  if (daysLeft <= 7) {
    return {
      daysLeft,
      urgency: 'warn7',
      label: `Còn ${daysLeft} ngày (${ymd})`,
    };
  }
  return { daysLeft, urgency: 'normal', label: `Còn ${daysLeft} ngày (${ymd})` };
}

export interface ResignationImportPreviewRow extends ResignationRecord {
  _rowIndex: number;
  _isValid: boolean;
  _warning?: string;
}

const COLUMN_ALIASES: Record<keyof Omit<ResignationRecord, 'id'>, RegExp[]> = {
  fullName: [/họ\s*(?:và\s*)?tên|tên\s*ứng\s*viên|氏名|名前|full\s*name|name/i],
  nationality: [/quốc\s*tịch|quoc\s*tich|国籍|nationality|country/i],
  previousCompany: [/công\s*ty\s*cũ|nơi\s*làm\s*cũ|前職|退職(?:会社|先|企業)|会社名|company/i],
  resignationDate: [/ngày\s*nghỉ\s*việc|退職日|resignation\s*date|nghỉ\s*việc/i],
  paperworkDeadline: [/hạn\s*(?:chót\s*)?(?:làm|nộp)?\s*giấy\s*tờ|hạn\s*nộp|書類提出期限|提出期限|deadline/i],
  paperworkStatus: [/trạng\s*thái|tình\s*trạng|進捗|ステータス|status/i],
  handoverNotes: [/ghi\s*chú\s*bàn\s*giao|bàn\s*giao|ghi\s*chú|メモ|引継|備考|notes?/i],
  newJobNeed: [/nhu\s*cầu\s*chuyển\s*việc|nhu\s*cầu|nguyện\s*vọng|転職希望|希望/i],
  phone: [/sđt|điện\s*thoại|phone|tel|連絡先/i],
  fbName: [/facebook|fb|tên\s*fb/i],
  fbProfileUrl: [/link\s*fb|fb\s*url|facebook\s*url/i],
  industry: [/ngành|職種|分野|industry/i],
  jlpt: [/jlpt|tiếng\s*nhật|日本語/i],
  convertedCandidateId: [],
  googleCalendarEventId: [],
  googleCalendarEventLink: [],
  manualProtectedFields: [],
  gasProtected: [],
  updatedAt: [],
};

export async function parseResignationSpreadsheetFile(
  file: File
): Promise<{ rows: ResignationImportPreviewRow[]; sheetName: string }> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array', cellDates: false });
  const sheetName = wb.SheetNames[0] || 'Sheet1';
  const ws = wb.Sheets[sheetName];
  if (!ws) return { rows: [], sheetName };

  const matrix = XLSX.utils.sheet_to_json<unknown[]>(ws, {
    header: 1,
    defval: '',
  });
  if (matrix.length === 0) return { rows: [], sheetName };

  // Locate header row within top 8 rows
  let headerRowIdx = 0;
  let bestScore = -1;
  for (let r = 0; r < Math.min(8, matrix.length); r++) {
    const rowCells = (matrix[r] || []).map((c) => String(c || '').trim());
    let score = 0;
    for (const cell of rowCells) {
      if (!cell) continue;
      if (Object.values(COLUMN_ALIASES).some((rxs) => rxs.some((rx) => rx.test(cell)))) {
        score += 1;
      }
    }
    if (score > bestScore) {
      bestScore = score;
      headerRowIdx = r;
    }
  }

  const headers = (matrix[headerRowIdx] || []).map((c) => String(c || '').trim());
  const colIndexMap: Partial<Record<keyof ResignationRecord, number>> = {};

  headers.forEach((h, idx) => {
    if (!h) return;
    (Object.keys(COLUMN_ALIASES) as Array<keyof Omit<ResignationRecord, 'id'>>).forEach(
      (field) => {
        if (colIndexMap[field] !== undefined) return;
        if (COLUMN_ALIASES[field].some((rx) => rx.test(h))) {
          colIndexMap[field] = idx;
        }
      }
    );
  });

  // Fallback positional mapping if headers weren't matched clearly
  if (colIndexMap.fullName === undefined && headers.length >= 2) {
    colIndexMap.fullName = 0;
    colIndexMap.nationality = colIndexMap.nationality ?? 1;
    colIndexMap.previousCompany = colIndexMap.previousCompany ?? 2;
    colIndexMap.resignationDate = colIndexMap.resignationDate ?? 3;
    colIndexMap.paperworkDeadline = colIndexMap.paperworkDeadline ?? 4;
    colIndexMap.paperworkStatus = colIndexMap.paperworkStatus ?? 5;
    colIndexMap.handoverNotes = colIndexMap.handoverNotes ?? 6;
    colIndexMap.newJobNeed = colIndexMap.newJobNeed ?? 7;
  }

  const getCell = (row: unknown[], field: keyof ResignationRecord): unknown => {
    const idx = colIndexMap[field];
    if (idx === undefined) return '';
    return row[idx] ?? '';
  };

  const parsedRows: ResignationImportPreviewRow[] = [];
  for (let r = headerRowIdx + 1; r < matrix.length; r++) {
    const row = matrix[r] || [];
    const nonEmptyCount = row.filter((c) => String(c || '').trim() !== '').length;
    if (nonEmptyCount === 0) continue;

    const fullName = String(getCell(row, 'fullName') || '').trim();
    const previousCompany = String(getCell(row, 'previousCompany') || '').trim();
    const rawNationality = getCell(row, 'nationality');
    const nationality = normalizeResignationNationality(rawNationality);
    const resignationDate = normalizeFlexibleDateToYmd(getCell(row, 'resignationDate'));
    const paperworkDeadline = normalizeFlexibleDateToYmd(
      getCell(row, 'paperworkDeadline')
    );
    const paperworkStatus = normalizePaperworkStatus(getCell(row, 'paperworkStatus'));
    const handoverNotes = String(getCell(row, 'handoverNotes') || '').trim();
    const newJobNeed = String(getCell(row, 'newJobNeed') || '').trim();
    const phone = String(getCell(row, 'phone') || '').trim();
    const fbName = String(getCell(row, 'fbName') || '').trim();
    const industry = String(getCell(row, 'industry') || '').trim();
    const jlpt = String(getCell(row, 'jlpt') || '').trim();

    const warnings: string[] = [];
    if (!fullName) warnings.push('Thiếu Họ tên');
    if (nationality === '国籍不明') warnings.push('Chưa rõ quốc tịch (cần chọn)');
    if (!previousCompany) warnings.push('Thiếu Công ty cũ');
    if (!paperworkDeadline) warnings.push('Chưa có hạn nộp giấy tờ');

    parsedRows.push({
      id: `resig_${Date.now()}_${r}`,
      fullName: fullName || `Chưa rõ tên (Dòng ${r + 1})`,
      nationality,
      previousCompany: previousCompany || 'Chưa rõ công ty cũ',
      resignationDate,
      paperworkDeadline,
      paperworkStatus,
      handoverNotes,
      newJobNeed,
      phone,
      fbName,
      industry,
      jlpt,
      _rowIndex: r + 1,
      _isValid: Boolean(fullName),
      _warning: warnings.length > 0 ? warnings.join(' / ') : undefined,
    });
  }

  return { rows: parsedRows, sheetName };
}

export function buildResignationReminderMessages(record: ResignationRecord): {
  vi: string;
  jp: string;
} {
  const deadlineText = record.paperworkDeadline || 'sớm nhất có thể';
  const companyText = record.previousCompany || 'công ty cũ';
  const name = record.fbName || record.fullName;

  return {
    vi: `Chào ${name}! Để hoàn tất thủ tục chuyển việc và hồ sơ xin visa mới sau khi nghỉ tại ${companyText} (ngày nghỉ: ${
      record.resignationDate || 'chưa rõ'
    }), bạn nhớ xin/nộp đầy đủ giấy tờ nghỉ việc (退職証明書, 源泉徴収票, 離職票 nếu có) trước hạn chót **${deadlineText}** nhé!\n${
      record.handoverNotes ? `Lưu ý hiện tại: ${record.handoverNotes}\n` : ''
    }Khi nào nhận được giấy tờ hoặc cần hỗ trợ nhắc công ty cũ, bạn chụp ảnh gửi lại cho mình ngay nha!`,
    jp: `${name}さん、お疲れ様です！${companyText}のご退職（退職日：${
      record.resignationDate || '未定'
    }）に伴う必要書類（退職証明書・源泉徴収票など）の提出期限は **${deadlineText}** となっております。\n${
      record.handoverNotes ? `メモ：${record.handoverNotes}\n` : ''
    }書類をお受け取りになりましたら、お写真でも構いませんので早めにお送りください！`,
  };
}

export function convertResignationToCandidate(record: ResignationRecord): Candidate {
  const todayYmd = getJstTodayYmd();
  const normalizedNationality = normalizeResignationNationality(record.nationality);
  const fullNotes = [
    `[Chuyển từ Quản lý nghỉ việc ${todayYmd}]`,
    `Công ty cũ: ${record.previousCompany}`,
    record.resignationDate ? `Ngày nghỉ việc (退職日): ${record.resignationDate}` : '',
    record.paperworkDeadline
      ? `Hạn nộp giấy tờ: ${record.paperworkDeadline} (${record.paperworkStatus})`
      : '',
    record.handoverNotes ? `Ghi chú bàn giao: ${record.handoverNotes}` : '',
    record.newJobNeed ? `Nhu cầu chuyển việc mới: ${record.newJobNeed}` : '',
  ]
    .filter(Boolean)
    .join('\n');

  return {
    id: `cand_resig_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    name: record.fullName,
    fbName: record.fbName || record.fullName,
    fbProfileUrl: record.fbProfileUrl,
    channel: 'personal',
    nationality: normalizedNationality,
    jlpt: record.jlpt || '要確認',
    desiredIndustry: record.industry || '',
    conversionGoal:
      record.newJobNeed ||
      `Chuyển đổi việc mới sau khi nghỉ tại ${record.previousCompany}`,
    visaStatus: '特定技能 / Đang chuyển việc (退職手続き中)',
    phone: record.phone || '',
    progressStatus: '新規相談',
    lastContactDate: todayYmd,
    negotiationLog: fullNotes,
    memo: fullNotes,
  };
}
