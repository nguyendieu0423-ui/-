export type KyujinSalaryType = 'hourly' | 'monthly' | 'unknown';

export interface ParsedKyujinSalary {
  salary: string;
  hourlyWage: string;
  salaryType: KyujinSalaryType;
  numericAmount: number | null;
}

const FULLWIDTH_DIGIT_OFFSET = 0xfee0;

export const normalizeJapaneseNumberText = (raw: string): string => {
  return String(raw || '')
    .replace(/[０-９]/g, (ch) =>
      String.fromCharCode(ch.charCodeAt(0) - FULLWIDTH_DIGIT_OFFSET)
    )
    .replace(/[，]/g, ',')
    .replace(/[．]/g, '.')
    .replace(/[：]/g, ':')
    .replace(/[〜～~]/g, '～')
    .replace(/[－−‐–—]/g, '-')
    .replace(/[￥¥]/g, '')
    .replace(/[Ｋｋ]/g, 'k');
};

const NON_SALARY_PREFIX_REGEX =
  /(?:夜勤手当|宿直手当|交通費|通勤手当|寮費|社宅費|家賃|光熱費|住宅手当|資格手当|家族手当|扶養手当|皆勤手当|処遇改善(?:加算)?|賞与|ボーナス|一時金|退職金|初期費用|控除(?!後)|社会保険|残業代|時間外手当|時給換算|上限|年収|想定年収|モデル年収|手取り|控除後)\s*[:：]?\s*[^\n,;、/／|]{0,36}$/i;

const SALARY_LABEL_SOURCE =
  '(?:想定月収|総支給額?|基本給|月給|時給|給与形態|給与|賃金|報酬|月収|Lương(?:\\s*(?:cơ bản|tháng|giờ|thực nhận))?|Salary)';

const SALARY_LINE_LABEL_REGEX = new RegExp(
  `^\\s*(?:[・●■▪▶★◆【\\[\\(]?\\s*)?(${SALARY_LABEL_SOURCE})\\s*[】\\]\\)]?\\s*(?:[:：|｜\\t]|\\s{2,}|(?=\\d|時給|月給))(.*)$`,
  'i'
);

const SALARY_INLINE_LABEL_REGEX = new RegExp(
  `(${SALARY_LABEL_SOURCE})\\s*(?:[:：|｜]|\\s+)\\s*([^\\n]+)`,
  'gi'
);

const NON_SALARY_UNIT_FOLLOWING_REGEX =
  /^\s*(?:年|ヶ月|カ月|か月|日|時間|時(?!給)|分|名|人|回|歳|才|%|％|駅|km|m\b|号|番|階)/i;

export const stripKyujinFormBoilerplate = (raw: string): string => {
  return String(raw || '')
    .replace(/社会保険加入事業者か[^\n]*/g, '')
    .replace(/[（(]例[)）][^\n]{0,45}/g, '')
    .replace(/※\s*[12１２][^\n]{0,60}/g, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
};

export const sanitizeExcelFloatText = (raw: unknown): string => {
  if (typeof raw === 'number' && Number.isFinite(raw)) {
    if (Math.abs(raw) >= 100 && Math.abs(raw - Math.round(raw)) < 0.05) {
      return String(Math.round(raw));
    }
    const rounded2 = Math.round(raw * 100) / 100;
    return String(rounded2);
  }
  const s = String(raw ?? '').trim();
  if (!s) return '';
  return s.replace(/(\d+)\.(?:9{4,}|0{4,}|59{4,}|[0-9]{7,})\d*/g, (m) => {
    const n = Number(m);
    if (!Number.isFinite(n)) return m;
    if (n >= 100 && Math.abs(n - Math.round(n)) < 0.05) {
      return String(Math.round(n));
    }
    return String(Math.round(n));
  });
};

export const extractEstimatedNetSalaryText = (rawText: string): string => {
  const cleaned = normalizeJapaneseNumberText(stripKyujinFormBoilerplate(rawText));
  const m = cleaned.match(
    /(?:手取り(?:目安)?|控除後(?:支給)?|実収入)\s*[：:・|｜\s]*(?:約)?\s*([0-9,]+(?:\.[0-9]+)?\s*(?:万円|万|円)?(?:\s*[～\-]\s*[0-9,]+(?:\.[0-9]+)?\s*(?:万円|万|円)?)?)/i
  );
  if (!m) return '';
  const amt = extractAmountSegment(m[1], { allowBareNumber: true });
  return amt ? amt.formatted : '';
};

export const extractAnnualSalaryText = (rawText: string): string => {
  const cleaned = normalizeJapaneseNumberText(stripKyujinFormBoilerplate(rawText));
  const m = cleaned.match(
    /(?:想定年収|モデル年収|年収)\s*[：:・|｜\s]*(?:約)?\s*([0-9,]+(?:\.[0-9]+)?\s*(?:万円|万|円)?(?:\s*[～\-]\s*[0-9,]+(?:\.[0-9]+)?\s*(?:万円|万|円)?)?)/i
  );
  if (!m) return '';
  const amt = extractAmountSegment(m[1], { allowBareNumber: false, maxYenLimit: 25000000 });
  return amt ? amt.formatted : '';
};

const parseAmountUnitToYen = (numStr: string, unitStr?: string): number | null => {
  const cleanNum = Number(String(numStr || '').replace(/,/g, ''));
  if (!Number.isFinite(cleanNum) || cleanNum <= 0) return null;
  const unit = String(unitStr || '').toLowerCase();
  if (unit.startsWith('万')) {
    return Math.round(cleanNum * 10000);
  }
  if (unit === 'k') {
    return Math.round(cleanNum * 1000);
  }
  return Math.round(cleanNum);
};

const formatYenNumber = (yen: number): string => `${Math.round(yen).toLocaleString('en-US')}円`;

interface RawAmountMatch {
  formatted: string;
  minYen: number;
  maxYen: number;
  hasHourlyHint: boolean;
  hasMonthlyHint: boolean;
  rawMatched: string;
}

const extractAmountSegment = (
  rawSegment: string,
  options: { allowBareNumber?: boolean; maxYenLimit?: number } = {}
): RawAmountMatch | null => {
  const normalized = normalizeJapaneseNumberText(rawSegment).trim();
  if (!normalized || /要確認|未定|応相談/.test(normalized)) {
    // Check if there is still an actual number after/before 要確認? Usually none.
    if (!/\d/.test(normalized)) return null;
  }

  const hasHourlyHint = /時給|Lương giờ|\/h\b|時間給/i.test(normalized);
  const hasMonthlyHint = /月給|基本給|想定月収|総支給|月収|Lương tháng|\/月/i.test(normalized);

  // Skip ranges that are clearly dates/times/phones/postal codes
  const sanitized = normalized
    .replace(/\b\d{4}[-/年]\d{1,2}[-/月]\d{1,2}日?\b/g, ' ')
    .replace(/\b\d{1,2}:\d{2}\s*[～\-]\s*\d{1,2}:\d{2}\b/g, ' ')
    .replace(/\b0\d{1,4}-\d{1,4}-\d{3,4}\b/g, ' ')
    .replace(/〒?\s*\d{3}-\d{4}/g, ' ');

  // Match pattern: <num><unit?> (optional range <num2><unit2?> or trailing ～) + optional (allowance note)
  const tokenPattern =
    '(?:(\\d{1,3}(?:,\\d{3})+|\\d+(?:\\.\\d+)?)\\s*(万円|万|k\\b|円)?)';
  const rangeRegex = new RegExp(
    `${tokenPattern}(?:\\s*([～\\-])\\s*${tokenPattern}|\\s*(～))?(\\s*[（(][^）)\\n]{1,42}[）)])?`,
    'gi'
  );

  let match: RegExpExecArray | null;
  while ((match = rangeRegex.exec(sanitized)) !== null) {
    const fullMatch = match[0];
    const num1 = match[1];
    const unit1 = match[2];
    const sep = match[3];
    const num2 = match[4];
    const unit2 = match[5];
    const trailingTilde = match[6];
    const parenNoteRaw = match[7];

    // Check what follows the numeric token (if no unit, avoid years/counts/dates)
    const afterMatchIdx = match.index + fullMatch.length;
    const afterText = sanitized.slice(afterMatchIdx);
    const beforeText = sanitized.slice(0, match.index);

    if (!unit1 && !unit2) {
      if (!options.allowBareNumber) continue;
      if (NON_SALARY_UNIT_FOLLOWING_REGEX.test(afterText)) continue;
    } else if (!unit1 && unit2 && sep === '-') {
      // ok: e.g. 1200-1500円
    }

    // Avoid matching non-salary allowance numbers appearing later in the text
    if (NON_SALARY_PREFIX_REGEX.test(beforeText)) {
      continue;
    }

    // If first part has no unit in a range like `20～25万円`, inherit `万` from unit2!
    const effectiveUnit1 =
      unit1 || (unit2 && unit2.startsWith('万') ? '万' : unit2 && unit2.toLowerCase() === 'k' ? 'k' : unit2 || '');
    const yen1 = parseAmountUnitToYen(num1, effectiveUnit1);
    const maxYenLimit = options.maxYenLimit ?? 3000000;
    if (!yen1 || yen1 < 500 || yen1 > maxYenLimit) continue;

    // Avoid matching standalone years like "2026" without unit
    if (!unit1 && !unit2 && yen1 >= 1990 && yen1 <= 2099) {
      continue;
    }

    let yen2: number | null = null;
    if (num2) {
      const effectiveUnit2 = unit2 || effectiveUnit1;
      yen2 = parseAmountUnitToYen(num2, effectiveUnit2);
      if (!yen2 || yen2 < 500 || yen2 > 3000000 || yen2 < yen1 * 0.5) {
        yen2 = null;
      }
    }

    const minYen = yen1;
    const maxYen = yen2 ?? yen1;
    let formatted = formatYenNumber(yen1);
    if (yen2) {
      formatted = `${formatYenNumber(yen1)}～${formatYenNumber(yen2)}`;
    } else if (trailingTilde || /～\s*$/.test(fullMatch)) {
      formatted = `${formatted}～`;
    }

    if (parenNoteRaw) {
      const cleanNote = parenNoteRaw
        .trim()
        .replace(/^\(/, '（')
        .replace(/\)$/, '）');
      if (cleanNote && !/^（\s*）$/.test(cleanNote)) {
        formatted += cleanNote;
      }
    }

    return {
      formatted,
      minYen,
      maxYen,
      hasHourlyHint,
      hasMonthlyHint,
      rawMatched: fullMatch.trim(),
    };
  }

  return null;
};

const classifyRawMatch = (
  match: RawAmountMatch,
  contextHintText = ''
): ParsedKyujinSalary => {
  const hintLower = normalizeJapaneseNumberText(contextHintText);
  const explicitHourly = match.hasHourlyHint || /時給|Lương giờ/i.test(hintLower);
  const explicitMonthly =
    match.hasMonthlyHint || /月給|基本給|想定月収|総支給|月収|Lương tháng/i.test(hintLower);

  // Rule: if contains 時給 OR numeric amount < 10,000円 -> hourly
  // If >= 10,000円 (and not explicitly labeled 時給 with small number) -> monthly
  const isHourly =
    match.minYen < 10000 || (explicitHourly && !explicitMonthly && match.minYen < 30000);

  if (isHourly) {
    return {
      salary: `時給 ${match.formatted}`,
      hourlyWage: match.formatted,
      salaryType: 'hourly',
      numericAmount: match.minYen,
    };
  }

  return {
    salary: match.formatted,
    hourlyWage: '',
    salaryType: 'monthly',
    numericAmount: match.minYen,
  };
};

/**
 * Extracts structured salary & hourlyWage from raw Japanese/Vietnamese job text,
 * supporting all common labels (給与, 月給, 時給, 基本給, 想定月収, 賃金, 報酬, 給与形態, Lương, Salary)
 * and formats (220,000円, 220000円, 1,200円, 1200円～1500円, 22万円, 22.5万円, 20万～25万円,
 * 月給23万, 220k, 時給1,150円〜（交通費支給）, 月給215,000円（夜勤手当含む）).
 */
export const extractKyujinSalaryInfo = (
  rawText: string,
  options: { assumeBareNumberIsSalary?: boolean } = {}
): ParsedKyujinSalary => {
  const text = stripKyujinFormBoilerplate(sanitizeExcelFloatText(String(rawText || '').trim()));
  if (!text) {
    return {
      salary: '給与要確認',
      hourlyWage: '',
      salaryType: 'unknown',
      numericAmount: null,
    };
  }

  const normalizedText = normalizeJapaneseNumberText(text);
  const lines = normalizedText
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);

  // Pass 1: Line-by-line explicit salary label check (with priority scoring)
  interface CandidateMatch {
    priority: number;
    parsed: ParsedKyujinSalary;
  }
  const candidates: CandidateMatch[] = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const lineMatch = line.match(SALARY_LINE_LABEL_REGEX);
    if (lineMatch) {
      const label = lineMatch[1];
      let valuePart = (lineMatch[2] || '').trim();
      // If value is on next line (common in PDF/Word/copy-paste)
      if ((!valuePart || !/\d/.test(valuePart)) && i + 1 < lines.length) {
        const nextLine = lines[i + 1];
        if (!SALARY_LINE_LABEL_REGEX.test(nextLine) && /\d/.test(nextLine)) {
          valuePart = `${valuePart} ${nextLine}`.trim();
        }
      }
      const extracted = extractAmountSegment(valuePart, { allowBareNumber: true });
      if (extracted) {
        const parsed = classifyRawMatch(extracted, `${label} ${valuePart}`);
        let priority = 50;
        const isTrialHourly =
          parsed.salaryType === 'hourly' && /試用期間|研修期間|見習い/.test(line);
        if (isTrialHourly) priority = 40;
        else if (/総支給|想定月収|月収/i.test(label)) priority = 100;
        else if (/月給|基本給/i.test(label)) priority = 90;
        else if (/時給/i.test(label)) priority = 85;
        else if (/給与|賃金|報酬|Salary|Lương/i.test(label)) priority = 80;
        else if (/給与形態/i.test(label)) priority = 70;
        candidates.push({ priority, parsed });
      }
    }

    // Inline labels inside multi-field single lines (e.g., "勤務地: 東京 / 月給: 23万")
    SALARY_INLINE_LABEL_REGEX.lastIndex = 0;
    let inlineM: RegExpExecArray | null;
    while ((inlineM = SALARY_INLINE_LABEL_REGEX.exec(line)) !== null) {
      const label = inlineM[1];
      const val = inlineM[2];
      const extracted = extractAmountSegment(val, { allowBareNumber: true });
      if (extracted) {
        const parsed = classifyRawMatch(extracted, `${label} ${val}`);
        let priority = 75;
        const isTrialHourly =
          parsed.salaryType === 'hourly' &&
          /試用期間|研修期間|見習い/.test(line.slice(0, (inlineM.index || 0) + label.length + 12));
        if (isTrialHourly) priority = 38;
        else if (/総支給|想定月収|月収/i.test(label)) priority = 98;
        else if (/月給|基本給/i.test(label)) priority = 88;
        else if (/時給/i.test(label)) priority = 84;
        candidates.push({ priority, parsed });
      }
    }
  }

  if (candidates.length > 0) {
    candidates.sort((a, b) => b.priority - a.priority);
    return candidates[0].parsed;
  }

  // Pass 2: Check for explicit 月給/時給/給与 words directly attached without colon (e.g. "月給23万")
  const attachedMatch = normalizedText.match(
    /(?:想定月収|総支給額?|基本給|月給|時給|月収)\s*[^\n,;]{0,40}/i
  );
  if (attachedMatch) {
    const extracted = extractAmountSegment(attachedMatch[0], { allowBareNumber: true });
    if (extracted) {
      return classifyRawMatch(extracted, attachedMatch[0]);
    }
  }

  // Pass 3: Unlabelled money figures in text (require currency unit unless assumeBareNumberIsSalary)
  const fallbackExtracted = extractAmountSegment(normalizedText, {
    allowBareNumber: Boolean(options.assumeBareNumberIsSalary),
  });
  if (fallbackExtracted) {
    return classifyRawMatch(fallbackExtracted, normalizedText);
  }

  return {
    salary: '給与要確認',
    hourlyWage: '',
    salaryType: 'unknown',
    numericAmount: null,
  };
};

/**
 * Extracts salary & hourlyWage from a 2D Excel/CSV sheet matrix, checking:
 * - same cell ("月給：23万円")
 * - right-hand cells (c+1..c+6)
 * - immediately below cell (r+1, c)
 * - fallback full block scan
 */
export const extractKyujinSalaryFromRows = (rows: unknown[][] = []): ParsedKyujinSalary => {
  if (!Array.isArray(rows) || rows.length === 0) {
    return {
      salary: '給与要確認',
      hourlyWage: '',
      salaryType: 'unknown',
      numericAmount: null,
    };
  }

  const labelRegex = new RegExp(SALARY_LABEL_SOURCE, 'i');
  const candidates: Array<{ priority: number; parsed: ParsedKyujinSalary }> = [];

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      const cell = normalizeJapaneseNumberText(String(row[c] || '')).trim();
      if (!cell || !labelRegex.test(cell)) continue;
      // Ignore allowance / non-salary headers that accidentally contain a word like 報酬/給与規定 etc if clearly unrelated
      if (/交通費|夜勤手当|賞与|寮費|住宅手当/.test(cell) && !/給与|月給|時給|基本給|想定月収|賃金/.test(cell)) {
        continue;
      }

      const priority = /総支給|想定月収|月収/i.test(cell)
        ? 100
        : /月給|基本給/i.test(cell)
        ? 90
        : /時給/i.test(cell)
        ? 85
        : /給与|賃金|報酬|Salary|Lương/i.test(cell)
        ? 80
        : 70;

      // 1) Try same cell first
      const sameCellExtracted = extractAmountSegment(cell, { allowBareNumber: true });
      if (sameCellExtracted) {
        candidates.push({ priority, parsed: classifyRawMatch(sameCellExtracted, cell) });
        continue;
      }

      // 2) Try adjacent right cells (c+1 .. c+6)
      const rightJoined = row
        .slice(c + 1, Math.min(row.length, c + 7))
        .map((v) => String(v || '').trim())
        .filter(Boolean)
        .join(' ');
      if (rightJoined) {
        const rightExtracted = extractAmountSegment(rightJoined, { allowBareNumber: true });
        if (rightExtracted) {
          candidates.push({
            priority,
            parsed: classifyRawMatch(rightExtracted, `${cell} ${rightJoined}`),
          });
          continue;
        }
      }

      // 3) Try cell directly below (r+1, c or c..c+1) — typical for table-header Excel/CSV layouts
      if (r + 1 < rows.length) {
        const belowRow = rows[r + 1] || [];
        const belowVal = String(belowRow[c] || '').trim();
        if (belowVal) {
          const belowExtracted = extractAmountSegment(belowVal, { allowBareNumber: true });
          if (belowExtracted) {
            candidates.push({
              priority: priority - 2,
              parsed: classifyRawMatch(belowExtracted, `${cell} ${belowVal}`),
            });
          }
        }
      }
    }
  }

  if (candidates.length > 0) {
    candidates.sort((a, b) => b.priority - a.priority);
    return candidates[0].parsed;
  }

  // Fallback: scan full text joined by newlines
  const fullText = rows
    .map((r) =>
      (r || [])
        .map((c) => String(c || '').trim())
        .filter(Boolean)
        .join(' ')
    )
    .filter(Boolean)
    .join('\n');

  return extractKyujinSalaryInfo(fullText);
};

/**
 * Normalizes a manual inline salary edit input from the UI card.
 * Preserves custom Japanese/Vietnamese text if no numeric yen value exists,
 * while automatically classifying <10,000円 or 時給 as hourlyWage and >=10,000円 as monthly salary.
 */
export const normalizeManualKyujinSalaryInput = (rawInput: string): ParsedKyujinSalary => {
  const trimmed = String(rawInput || '').trim();
  if (!trimmed || /要確認/.test(trimmed)) {
    return {
      salary: '給与要確認',
      hourlyWage: '',
      salaryType: 'unknown',
      numericAmount: null,
    };
  }

  const parsed = extractKyujinSalaryInfo(trimmed, { assumeBareNumberIsSalary: true });
  if (parsed.salaryType !== 'unknown') {
    return parsed;
  }

  // Custom textual salary without parseable yen amount (e.g. "応相談", "規定による")
  const isHourlyText = /時給|Lương giờ/i.test(trimmed);
  return {
    salary: trimmed,
    hourlyWage: isHourlyText ? trimmed.replace(/^時給[：:\s]*/i, '').trim() : '',
    salaryType: isHourlyText ? 'hourly' : 'monthly',
    numericAmount: null,
  };
};

export const isKyujinSalaryMissing = (
  slot: { salary?: string; hourlyWage?: string },
  isBaito = false
): boolean => {
  const primary = isBaito
    ? String(slot.hourlyWage || slot.salary || '').trim()
    : String(slot.salary || slot.hourlyWage || '').trim();
  if (!primary) return true;
  if (/要確認/.test(primary)) return true;
  return false;
};

export const getKyujinSlotSalaryDisplayMeta = (
  slot: { salary?: string; hourlyWage?: string; employmentType?: string },
  companyEmploymentType = ''
): {
  isHourly: boolean;
  isMissing: boolean;
  labelPrefix: '時給' | 'Lương';
  displayValue: string;
  editInitialValue: string;
} => {
  const empType = slot.employmentType || companyEmploymentType || '';
  const isBaito = /baito|アルバイト|パート/i.test(empType);
  const rawSalary = String(slot.salary || '').trim();
  const rawHourly = String(slot.hourlyWage || '').trim();
  const isMissing = isKyujinSalaryMissing(slot, isBaito);

  if (isMissing) {
    return {
      isHourly: isBaito || /時給/.test(rawSalary),
      isMissing: true,
      labelPrefix: isBaito || /時給/.test(rawSalary) ? '時給' : 'Lương',
      displayValue: '給与要確認',
      editInitialValue: '',
    };
  }

  const combined = rawHourly || rawSalary;
  const parsed = extractKyujinSalaryInfo(
    rawHourly ? `時給 ${rawHourly}` : rawSalary,
    { assumeBareNumberIsSalary: true }
  );

  const isHourly =
    Boolean(rawHourly) ||
    /時給/.test(rawSalary) ||
    parsed.salaryType === 'hourly' ||
    (isBaito && (parsed.numericAmount === null || parsed.numericAmount < 10000));

  if (isHourly) {
    const cleanHourly = (rawHourly || rawSalary).replace(/^時給[：:\s]*/i, '').trim();
    return {
      isHourly: true,
      isMissing: false,
      labelPrefix: '時給',
      displayValue: cleanHourly || '給与要確認',
      editInitialValue: cleanHourly,
    };
  }

  const cleanMonthly = rawSalary.replace(/^月給[：:\s]*/i, '').trim();
  return {
    isHourly: false,
    isMissing: false,
    labelPrefix: 'Lương',
    displayValue: cleanMonthly || rawSalary || '給与要確認',
    editInitialValue: rawSalary,
  };
};
