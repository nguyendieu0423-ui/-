import * as XLSX from 'xlsx';
import type {
  Candidate,
  ResignationPaperworkStatus,
  ResignationRecord,
} from '../types/crm';

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
  nationality: [/quốc\s*tịch|国籍|nationality|country/i],
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
    const nationality = String(getCell(row, 'nationality') || '').trim();
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
    if (!previousCompany) warnings.push('Thiếu Công ty cũ');
    if (!paperworkDeadline) warnings.push('Chưa có hạn nộp giấy tờ');

    parsedRows.push({
      id: `resig_${Date.now()}_${r}`,
      fullName: fullName || `Chưa rõ tên (Dòng ${r + 1})`,
      nationality: nationality || 'ベトナム',
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
    nationality: record.nationality || 'ベトナム',
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
