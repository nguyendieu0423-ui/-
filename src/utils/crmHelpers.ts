import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import mammoth from 'mammoth';
import type {
  Candidate,
  ExcelAnchor,
  KyujinCompany,
  KyujinMatchResult,
  KyujinSlot,
  MissingInfoItem,
  ReplyScenario,
  RirekishoDraft,
} from '../types/crm';
import { compressImageFileToBase64, fileToBase64Raw } from './imageCompression';

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
    nationality: cand.nationality || 'ベトナム',
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
  };
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

export const isActualProposedRecord = (person: Partial<Candidate> = {}): boolean => {
  const comp = getEffectiveProposedCompany(person);
  if (comp) return true;
  const st = String(person.progressStatus || '').trim();
  return [
    '企業提案中',
    '1次面接予定',
    '2次面接予定',
    '面接調整中',
    '内定',
    '入社手続き中',
  ].includes(st);
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

export const getUrgentFollowUpReason = (person: Partial<Candidate> = {}): string => {
  const now = new Date();
  const todayIso = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(
    now.getDate()
  ).padStart(2, '0')}`;
  if (String(person.lastContactDate || '').trim() === todayIso) {
    return '';
  }

  const st = String(person.progressStatus || '').trim();
  if (['保留・連絡待ち', '面談調整中', '面接調整中', 'ヒアリング中'].includes(st)) {
    return `ステータス：${st}`;
  }
  const mendan = extractNormalizedMendanInfo(person);
  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  if (mendan.timestamp > 0) {
    const diffDays = Math.round((mendan.timestamp - todayStart) / 86400000);
    if (diffDays >= 0 && diffDays <= 3) {
      return `面談/PV直前（${mendan.dateLabel}）`;
    }
    if (diffDays < 0 && !['内定', '入社手続き中', '辞退・終了'].includes(st)) {
      return `予定経過・要状況確認（${mendan.dateLabel}）`;
    }
  }
  if (!['内定', '入社手続き中', '辞退・終了'].includes(st)) {
    const lastTs = parseFlexibleDateToTimestamp(person.lastContactDate, 2026);
    if (!lastTs) return '最終連絡日未記録';
    const daysSince = Math.floor((todayStart - lastTs) / 86400000);
    if (daysSince >= 7) return `${daysSince}日間連絡なし`;
  }
  return '';
};

const jlptRank = (lvl = ''): number => {
  const s = String(lvl || '').toUpperCase();
  if (s.includes('N1')) return 5;
  if (s.includes('N2')) return 4;
  if (s.includes('N3')) return 3;
  if (s.includes('N4')) return 2;
  if (s.includes('N5') || s.includes('日常')) return 1;
  if (s.includes('不問')) return 0;
  return 2;
};

const PREFECTURE_PATTERNS: Array<{ rx: RegExp; label: string; keys: string[] }> = [
  { rx: /kyoto|京都府|京都/i, label: '京都府（Kyoto）', keys: ['京都', 'kyoto'] },
  { rx: /osaka|大阪府|大阪/i, label: '大阪府（Osaka）', keys: ['大阪', 'osaka'] },
  { rx: /hyogo|kobe|兵庫県|神戸/i, label: '兵庫県（Hyogo）', keys: ['兵庫', '神戸', 'hyogo', 'kobe'] },
  { rx: /nara|奈良県|奈良/i, label: '奈良県（Nara）', keys: ['奈良', 'nara'] },
  { rx: /shiga|滋賀県|滋賀/i, label: '滋賀県（Shiga）', keys: ['滋賀', 'shiga'] },
  { rx: /aichi|nagoya|愛知県|名古屋/i, label: '愛知県（Aichi）', keys: ['愛知', '名古屋', 'aichi', 'nagoya'] },
  { rx: /shizuoka|静岡県|静岡/i, label: '静岡県（Shizuoka）', keys: ['静岡', 'shizuoka'] },
  { rx: /fukuoka|福岡県|福岡/i, label: '福岡県（Fukuoka）', keys: ['福岡', 'fukuoka'] },
  { rx: /okinawa|沖縄県|沖縄/i, label: '沖縄県（Okinawa）', keys: ['沖縄', 'okinawa'] },
  { rx: /hokkaido|sapporo|北海道|札幌/i, label: '北海道（Hokkaido）', keys: ['北海道', '札幌', 'hokkaido', 'sapporo'] },
  { rx: /miyagi|sendai|宮城県|仙台/i, label: '宮城県（Miyagi）', keys: ['宮城', '仙台', 'miyagi', 'sendai'] },
  { rx: /hiroshima|広島県|広島/i, label: '広島県（Hiroshima）', keys: ['広島', 'hiroshima'] },
  { rx: /kanagawa|yokohama|kawasaki|神奈川県|神奈川|横浜|川崎/i, label: '神奈川県（Kanagawa）', keys: ['神奈川', '横浜', '川崎', 'kanagawa', 'yokohama', 'kawasaki'] },
  { rx: /saitama|埼玉県|埼玉/i, label: '埼玉県（Saitama）', keys: ['埼玉', 'saitama'] },
  { rx: /chiba|千葉県|千葉/i, label: '千葉県（Chiba）', keys: ['千葉', 'chiba'] },
  { rx: /ibaraki|茨城県|茨城/i, label: '茨城県（Ibaraki）', keys: ['茨城', 'ibaraki'] },
  { rx: /tochigi|栃木県|栃木/i, label: '栃木県（Tochigi）', keys: ['栃木', 'tochigi'] },
  { rx: /gunma|群馬県|群馬/i, label: '群馬県（Gunma）', keys: ['群馬', 'gunma'] },
  { rx: /tokyo|東京都|東京/i, label: '東京都（Tokyo）', keys: ['東京', 'tokyo'] },
  { rx: /nagano|長野県|長野/i, label: '長野県（Nagano）', keys: ['長野', 'nagano'] },
  { rx: /gifu|岐阜県|岐阜/i, label: '岐阜県（Gifu）', keys: ['岐阜', 'gifu'] },
  { rx: /mie|三重県|三重/i, label: '三重県（Mie）', keys: ['三重', 'mie'] },
  { rx: /niigata|新潟県|新潟/i, label: '新潟県（Niigata）', keys: ['新潟', 'niigata'] },
  { rx: /okayama|岡山県|岡山/i, label: '岡山県（Okayama）', keys: ['岡山', 'okayama'] },
  { rx: /kumamoto|熊本県|熊本/i, label: '熊本県（Kumamoto）', keys: ['熊本', 'kumamoto'] },
  { rx: /kagoshima|鹿児島県|鹿児島/i, label: '鹿児島県（Kagoshima）', keys: ['鹿児島', 'kagoshima'] },
];

export const extractAllAreaTokens = (rawArea = ''): Array<{ label: string; keys: string[] }> => {
  const s = String(rawArea || '').trim();
  if (!s) return [];
  const matched: Array<{ label: string; keys: string[] }> = [];
  for (const pref of PREFECTURE_PATTERNS) {
    if (pref.rx.test(s)) {
      matched.push({ label: pref.label, keys: pref.keys });
    }
  }
  return matched;
};

export const doesAreaMatchFilter = (itemArea = '', filterText = ''): boolean => {
  const f = String(filterText || '').trim().toLowerCase();
  if (!f) return true;
  const target = String(itemArea || '').toLowerCase();
  if (!target) return false;
  if (target.includes(f)) return true;

  const filterPrefs = extractAllAreaTokens(f);
  if (filterPrefs.length > 0) {
    return filterPrefs.some((fp) => fp.keys.some((k) => target.includes(k)));
  }
  return false;
};

export const findMatchingKyujinsForCandidate = (
  cand: Partial<Candidate> = {},
  kyujinsList: KyujinCompany[] = []
): KyujinMatchResult[] => {
  const candInd = String(cand.desiredIndustry || '').trim();
  const candArea = String(cand.desiredArea || '').trim();
  const candGoal = String(cand.conversionGoal || '').trim();
  const candExp = String(cand.japanExperience || '').trim();
  const candNat = String(cand.nationality || '').trim();
  const candSalaryExp = String(cand.salaryExpectation || '').trim();
  const candJlptScore = jlptRank(cand.jlpt);
  const candAreaPrefs = extractAllAreaTokens(candArea);

  const wantsBaito = /baito|アルバイト|パート|làm thêm/i.test(
    `${candGoal} ${cand.memo || ''} ${cand.visaStatus || ''}`
  );

  const results: KyujinMatchResult[] = [];
  (kyujinsList || []).forEach((comp) => {
    (comp.areaBreakdown || []).forEach((sl) => {
      if (sl.recruitmentStatus === '募集終了') return;

      if (
        candNat &&
        Array.isArray(comp.excludedNationalities) &&
        comp.excludedNationalities.includes(candNat)
      ) {
        return;
      }

      let rawScore = 0;
      const reasons: string[] = [];

      // 1. Sector / Industry (up to 35 pts)
      const slotInd = String(sl.industry || comp.industry || '').trim();
      if (candInd && slotInd && (slotInd.includes(candInd) || candInd.includes(slotInd))) {
        rawScore += 35;
        reasons.push(`✅ Đúng ngành ${slotInd}`);
      } else if (!candInd) {
        rawScore += 15;
      }

      // 2. Employment Type (Tokutei / Seishain / Baito) (up to 15 pts)
      const empType = String(sl.employmentType || comp.employmentType || 'Tokutei Ginou (特定技能)');
      const slotIsBaito = /baito|アルバイト|パート/i.test(empType);
      if (wantsBaito && slotIsBaito) {
        rawScore += 15;
        reasons.push('✅ Đúng hình thức Baito');
      } else if (!wantsBaito && !slotIsBaito) {
        rawScore += 10;
      } else if (wantsBaito && !slotIsBaito) {
        rawScore -= 10;
      }

      // 3. Multi-prefecture Area matching (up to 25 pts)
      const slotArea = String(sl.area || '').trim();
      if (candArea && slotArea) {
        const matchedPref = candAreaPrefs.find((p) =>
          p.keys.some((k) => slotArea.toLowerCase().includes(k))
        );
        if (matchedPref) {
          rawScore += 25;
          const shortPref = matchedPref.label.split('（')[0];
          reasons.push(`✅ Khớp ${shortPref}`);
        } else {
          const fallbackTokens = candArea
            .split(/[・、,\s\/]+/)
            .map((t) => t.replace(/[都道府県市区町村]/g, '').trim())
            .filter((t) => t.length >= 2);
          const hit = fallbackTokens.find((tok) => slotArea.includes(tok));
          if (hit) {
            rawScore += 25;
            reasons.push(`✅ Khớp ${hit}`);
          }
        }
      }

      // 4. JLPT requirement (up to 20 pts)
      const reqJlpt = String(sl.minJlpt || '要確認');
      if (reqJlpt === '要確認' || reqJlpt === '不問') {
        rawScore += 12;
      } else if (candJlptScore >= jlptRank(reqJlpt)) {
        rawScore += 20;
        reasons.push(`✅ Đạt ${cand.jlpt || reqJlpt}`);
      } else {
        rawScore -= 15;
      }

      // 5. Experience requirement (up to 10 pts)
      const expReq = String(sl.experienceRequirement || '');
      if (/未経験|không yêu cầu/i.test(expReq)) {
        rawScore += 10;
        reasons.push('✅ Không yêu cầu kinh nghiệm');
      } else if (/経験者|có kinh nghiệm/i.test(expReq)) {
        if (candExp) {
          rawScore += 10;
          reasons.push('✅ Đáp ứng kinh nghiệm');
        } else {
          rawScore -= 5;
        }
      } else if (candExp) {
        rawScore += 5;
      }

      // 6. Salary fit / bonus (up to 10 pts)
      const salaryText = `${sl.salary || ''} ${sl.estimatedNetSalary || ''} ${sl.hourlyWage || ''}`;
      const numSalaryM = (sl.salary || '').replace(/,/g, '').match(/(\d{5,6})/);
      const expSalaryM = candSalaryExp.match(/(\d{2})\s*man/i);
      if (numSalaryM && expSalaryM) {
        const offeredYen = Number(numSalaryM[1]);
        const desiredYen = Number(expSalaryM[1]) * 10000;
        if (offeredYen >= desiredYen) {
          rawScore += 10;
          reasons.push('✅ Đạt mức lương kỳ vọng');
        }
      } else if (sl.bonus === '賞与あり') {
        rawScore += 5;
        reasons.push('🎁 Có thưởng');
      } else if (salaryText.trim()) {
        rawScore += 5;
      }

      const clampedScore = Math.max(15, Math.min(100, rawScore));
      if (clampedScore >= 35 || reasons.length >= 2) {
        results.push({
          score: clampedScore,
          reasons: reasons.slice(0, 5),
          company: comp,
          slot: sl,
        });
      }
    });
  });

  results.sort((a, b) => b.score - a.score);
  return results.slice(0, 5);
};

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
  const expText =
    slot.experienceRequirement ||
    (slot.careQualification && slot.careQualification !== '不問'
      ? `Ưu tiên có chứng chỉ/kinh nghiệm (${slot.careQualification})`
      : 'Không yêu cầu kinh nghiệm (未経験OK)');

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
          slot.bonus === '賞与あり' ? `賞与あり（${slot.bonusDetail || '年2回'}）` : '',
        ]
          .filter(Boolean)
          .join(' ／ ');

    return [
      `【求人案内（匿名）】`,
      `・雇用形態：${empType}`,
      `・分野 / 職種：${sector} - ${industry}${slot.facilityType ? `（${slot.facilityType}）` : ''}`,
      `・勤務地：${publicArea}`,
      `・仕事内容：${slot.jobDescription || '面接時に詳しくご説明いたします。'}`,
      `・経験要件：${expText}${slot.experienceDetail ? `（${slot.experienceDetail}）` : ''}`,
      `・日本語レベル：${slot.minJlpt || '応相談'}`,
      `・給与・待遇：${compPay}`,
      `・住宅・寮：${slot.housingSupport || '要確認'}`,
      slot.shiftTime ? `・勤務時間：${slot.shiftTime}` : '',
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
        slot.bonus === '賞与あり' ? `Thưởng: Có (${slot.bonusDetail || 'theo quy định'})` : '',
      ]
        .filter(Boolean)
        .join(' | ');

  return [
    `📢 [ĐƠN TUYỂN DỤNG ẨN DANH]`,
    `• Hình thức: ${empType}`,
    `• Lĩnh vực - Ngành nghề: ${sector} - ${industry}${slot.facilityType ? ` (${slot.facilityType})` : ''}`,
    `• Khu vực làm việc: ${publicArea}`,
    `• Nội dung công việc: ${slot.jobDescription || 'Trao đổi chi tiết khi Mendan'}`,
    `• Yêu cầu kinh nghiệm: ${expText}${slot.experienceDetail ? ` — ${slot.experienceDetail}` : ''}`,
    `• Trình độ tiếng Nhật: ${slot.minJlpt || 'Trao đổi thêm'}`,
    `• Chế độ lương thưởng: ${paySummaryVi}`,
    `• Nhà ở / KTX: ${slot.housingSupport || 'Xác nhận khi phỏng vấn'}`,
    slot.shiftTime ? `• Thời gian làm việc: ${slot.shiftTime}` : '',
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
  if (b === '賞与あり' || /賞与あり|ボーナスあり|決算賞与/.test(`${b} ${d} ${fb}`)) {
    const detailMatch = `${d} ${fb}`.match(
      /年\s*\d+\s*回[^\s。、）)]*|\d+(?:,\d+)*\s*万?円[^\s。、）)]*/
    );
    return {
      bonus: '賞与あり',
      bonusDetail: d || (detailMatch ? detailMatch[0] : ''),
    };
  }
  if (b === '賞与なし' || /賞与なし|賞与無/.test(`${b} ${d}`)) {
    return { bonus: '賞与なし', bonusDetail: '' };
  }
  return { bonus: '要確認', bonusDetail: '' };
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
  return String(desc || '')
    .replace(
      /(?:受動喫煙|試用期間|選考方法|応募書類|担当者|連絡先|電話番号|FAX|メールアドレス|所在地|法人番号)[\s\S]*$/i,
      ''
    )
    .replace(/\s+/g, ' ')
    .trim();
};

export const normalizeKyujinItem = (
  comp: Partial<KyujinCompany> & Partial<KyujinSlot> = {}
): KyujinCompany => {
  const rawSlots: Partial<KyujinSlot>[] =
    Array.isArray(comp.areaBreakdown) && comp.areaBreakdown.length > 0
      ? comp.areaBreakdown
      : [
          {
            slotId: `${comp.id || 'kj'}_slot_1`,
            sector: comp.sector,
            industry: comp.industry,
            employmentType: comp.employmentType,
            jobCode: comp.jobCode,
            area: comp.area,
            shiftTime: comp.shiftTime,
            salary: comp.salary,
            nightShiftAllowance: comp.nightShiftAllowance,
            estimatedNetSalary: comp.estimatedNetSalary,
            hourlyWage: comp.hourlyWage,
            commuteAllowance: comp.commuteAllowance,
            shiftsPerWeek: comp.shiftsPerWeek,
            bonus: comp.bonus,
            bonusDetail: comp.bonusDetail,
            minJlpt: comp.minJlpt,
            housingSupport: comp.housingSupport,
            recruitmentStatus: comp.recruitmentStatus || '募集中',
            careQualification: comp.careQualification || '',
            experienceRequirement: comp.experienceRequirement,
            experienceDetail: comp.experienceDetail,
            jobDescription: comp.jobDescription || '',
          },
        ];

  const normalizedSlots: KyujinSlot[] = rawSlots.map((sl, idx) => {
    const cleanDesc = sanitizeJobDescriptionForCandidate(
      sl.jobDescription || comp.jobDescription || ''
    );
    const slIndustry = sl.industry || comp.industry || '特定技能';
    const slSector =
      sl.sector ||
      comp.sector ||
      detectSectorFromIndustryOrText(slIndustry, cleanDesc);
    const slEmploymentType =
      sl.employmentType ||
      comp.employmentType ||
      detectEmploymentTypeFromText(`${sl.salary || ''} ${sl.hourlyWage || ''} ${cleanDesc}`);
    const facilityType =
      sl.facilityType ||
      extractSingleKaigoFacilityTypeForSlot(
        { ...sl, jobDescription: cleanDesc },
        slIndustry
      );
    const bonusNorm = normalizeBonusInfo(
      sl.bonus || comp.bonus,
      sl.bonusDetail || comp.bonusDetail,
      `${sl.jobDescription || ''} ${comp.jobDescription || ''}`
    );
    const expReq =
      sl.experienceRequirement ||
      comp.experienceRequirement ||
      (/経験者必須|経験あり必須|要経験/.test(cleanDesc)
        ? 'Yêu cầu có kinh nghiệm (経験者必須)'
        : /未経験OK|未経験歓迎|不問/.test(`${cleanDesc} ${sl.careQualification || ''}`)
        ? 'Không yêu cầu kinh nghiệm (未経験OK)'
        : '');

    return {
      ...sl,
      slotId: sl.slotId || `${comp.id || 'kj'}_slot_${idx + 1}`,
      sector: slSector,
      industry: slIndustry,
      employmentType: slEmploymentType,
      jobCode: String(sl.jobCode || comp.jobCode || '').trim(),
      area: String(sl.area || comp.area || '').trim(),
      shiftTime: String(sl.shiftTime || comp.shiftTime || '').trim(),
      salary: String(sl.salary || comp.salary || (sl.hourlyWage ? `時給 ${sl.hourlyWage}` : '給与要確認')).trim(),
      nightShiftAllowance: String(sl.nightShiftAllowance || comp.nightShiftAllowance || '').trim(),
      estimatedNetSalary: String(sl.estimatedNetSalary || comp.estimatedNetSalary || '').trim(),
      hourlyWage: String(sl.hourlyWage || comp.hourlyWage || '').trim(),
      commuteAllowance: String(sl.commuteAllowance || comp.commuteAllowance || '').trim(),
      shiftsPerWeek: String(sl.shiftsPerWeek || comp.shiftsPerWeek || '').trim(),
      bonus: bonusNorm.bonus,
      bonusDetail: bonusNorm.bonusDetail,
      minJlpt: String(sl.minJlpt || comp.minJlpt || '要確認').trim(),
      housingSupport: normalizeHousingDisplay(sl.housingSupport || comp.housingSupport),
      recruitmentStatus:
        sl.recruitmentStatus === '募集終了' || comp.recruitmentStatus === '募集終了'
          ? '募集終了'
          : '募集中',
      careQualification: sl.careQualification || comp.careQualification || '',
      facilityType,
      experienceRequirement: expReq,
      experienceDetail: String(sl.experienceDetail || comp.experienceDetail || '').trim(),
      jobDescription: cleanDesc,
    };
  });

  const hasAnyOpen = normalizedSlots.some((s) => s.recruitmentStatus !== '募集終了');
  const primaryIndustry =
    normalizedSlots[0]?.industry || String(comp.industry || '特定技能').trim();
  const primarySector =
    normalizedSlots[0]?.sector ||
    comp.sector ||
    detectSectorFromIndustryOrText(primaryIndustry);
  const primaryEmpType =
    normalizedSlots[0]?.employmentType ||
    comp.employmentType ||
    'Tokutei Ginou (特定技能)';

  return {
    id: comp.id || `kj_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    companyName: String(comp.companyName || '名称未設定企業')
      .replace(/\s*代表者.*$/, '')
      .trim(),
    sector: primarySector,
    industry: primaryIndustry,
    employmentType: primaryEmpType,
    recruitmentStatus: hasAnyOpen ? '募集中' : '募集終了',
    excludedNationalities: comp.excludedNationalities,
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
      if (/[☑✓✔■●レ\uF0FE\uF052\u00FE\u00FC]\s*(N[1-5]|日常会話|不問)/i.test(cell)) {
        const m = cell.match(/(N[1-5]|日常会話|不問)/i);
        if (m) return m[1].toUpperCase();
      }
      if (
        /^(N[1-5]|日常会話|不問)$/i.test(cell) &&
        /^[☑✓✔■●レ\uF0FE\uF052\u00FE\u00FC]$/.test(prevCell)
      ) {
        return cell.toUpperCase();
      }
    }
  }

  const fullText = rows.map((r) => (r || []).join(' ')).join(' ');
  const condMatch = fullText.match(/日本語(?:能力|レベル)?\s*[：:]?\s*(N[1-5])\s*以上/i);
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

const extractSalaryFromRows = (rows: unknown[][] = []): string => {
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      const cell = String(row[c] || '').trim();
      if (/総支給/.test(cell)) {
        for (let k = c + 1; k < Math.min(row.length, c + 5); k++) {
          const val = String(row[k] || '').trim();
          const m = val.match(/(\d{2,3}(?:,\d{3})+|\d{5,6})\s*円?/);
          if (m) return `${Number(m[1].replace(/,/g, '')).toLocaleString()}円`;
        }
      }
    }
  }
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      const cell = String(row[c] || '').trim();
      if (/基本給|月給/.test(cell)) {
        for (let k = c + 1; k < Math.min(row.length, c + 5); k++) {
          const val = String(row[k] || '').trim();
          const m = val.match(/(\d{2,3}(?:,\d{3})+|\d{5,6})\s*円?/);
          if (m) return `${Number(m[1].replace(/,/g, '')).toLocaleString()}円`;
        }
      }
    }
  }
  return '給与要確認';
};

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
        .map((v) => String(v || '').trim())
        .filter(Boolean);

      if (!companyName && /^企業名$/.test(cell) && rightValues.length > 0) {
        companyName = rightValues[0].replace(/\s*代表者.*$/, '').trim();
      }
      if (!jobCode && /^案件番号$/.test(cell) && rightValues.length > 0) {
        jobCode = rightValues[0].trim();
      }
      if (!workArea && /^勤務先$/.test(cell) && rightValues.length > 0) {
        workArea = rightValues.join(' ').trim();
      }
      if (!fallbackLocation && /^所在地$/.test(cell) && rightValues.length > 0) {
        fallbackLocation = rightValues.join(' ').trim();
      }
      if (!jobDescription && /主な仕事内容|仕事内容/.test(cell) && rightValues.length > 0) {
        jobDescription = rightValues.join(' ').trim();
      }
      if (/賞与/.test(cell) && rightValues.length > 0) {
        bonusText += ' ' + rightValues.join(' ');
      }
    }
  }

  const fullText = blockRows.map((r) => (r || []).join(' ')).join('\n');
  if (!companyName && !jobCode && !workArea && !/求人票|特定技能/.test(fullText)) {
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
  const salary = extractSalaryFromRows(blockRows);
  const bonusNorm = normalizeBonusInfo('', '', `${bonusText} ${fullText}`);
  const hourlyM = fullText.match(/時給\s*([0-9,]+\s*円[^\n]*)/);
  const netM = fullText.match(/(?:手取り|実収入|ước tính thực nhận)[：:\s]*([^\n]{3,30})/i);
  const nightM = fullText.match(/(?:夜勤手当)[：:\s]*([^\n]{2,30})/);
  const expReq = /経験者必須|経験あり必須|要経験/.test(fullText)
    ? 'Yêu cầu có kinh nghiệm (経験者必須)'
    : /未経験OK|未経験歓迎|未経験可/.test(fullText)
    ? 'Không yêu cầu kinh nghiệm (未経験OK)'
    : '';

  return {
    companyName: finalCompany,
    sector,
    industry,
    employmentType,
    recruitmentStatus: '募集中',
    slot: {
      slotId: `slot_xl_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
      sector,
      industry,
      employmentType,
      jobCode: jobCode || sheetName || cleanArea || '新規枠',
      area: cleanArea || '勤務地要確認',
      salary,
      nightShiftAllowance: nightM ? nightM[1].trim() : '',
      estimatedNetSalary: netM ? netM[1].trim() : '',
      hourlyWage: hourlyM ? hourlyM[1].trim() : '',
      bonus: bonusNorm.bonus,
      bonusDetail: bonusNorm.bonusDetail,
      minJlpt,
      housingSupport,
      recruitmentStatus: '募集中',
      careQualification: '',
      experienceRequirement: expReq,
      jobDescription: sanitizeJobDescriptionForCandidate(jobDescription),
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
  const blocks = rawText
    .split(/(?:\n\s*[-=]{4,}\s*\n|\n{3,}(?=(?:企業名|会社名|【|■|Đơn|Job)))/i)
    .map((b) => b.trim())
    .filter((b) => b.length >= 20);
  const targetBlocks = blocks.length > 0 ? blocks : [rawText.trim()];

  return targetBlocks.map((block, idx) => {
    const compM = block.match(/(?:企業名|会社名|Công ty)[：:\s]*([^\n]+)/i);
    const areaM = block.match(/(?:勤務地|エリア|Khu vực|Địa điểm)[：:\s]*([^\n]+)/i);
    const salM = block.match(/(?:総支給|基本給|月給|Lương|時給)[：:\s]*([^\n]+)/i);
    const jlptM = block.match(/\b(N[1-5])\b/i);
    const hourlyM = block.match(/時給\s*([0-9,]+\s*円[^\n]*)/);
    const industry = detectIndustryFromText(block, sourceTitle);
    const sector = detectSectorFromIndustryOrText(industry, block);
    const employmentType = detectEmploymentTypeFromText(block, sourceTitle);
    const inferredArea = areaM ? areaM[1].trim() : extractPreferredAreaFromChat(block) || 'Nhật Bản';

    return {
      companyName: compM ? compM[1].trim() : `${sourceTitle} #${idx + 1}`,
      sector,
      industry,
      employmentType,
      jobCode: `JOB-${idx + 1}`,
      area: inferredArea,
      salary: salM ? salM[1].trim() : hourlyM ? `時給 ${hourlyM[1].trim()}` : '給与要確認',
      hourlyWage: hourlyM ? hourlyM[1].trim() : '',
      minJlpt: jlptM ? jlptM[1].toUpperCase() : '要確認',
      housingSupport: /自社寮/.test(block)
        ? '自社寮あり'
        : /借上げ|借り上げ/.test(block)
        ? '借上げ宿舎あり'
        : /寮あり|社宅/.test(block)
        ? '寮・社宅あり'
        : '要確認',
      bonus: /賞与あり|ボーナスあり/.test(block) ? '賞与あり' : '要確認',
      experienceRequirement: /未経験OK|không yêu cầu kinh nghiệm/i.test(block)
        ? 'Không yêu cầu kinh nghiệm (未経験OK)'
        : /経験者必須|cần kinh nghiệm/i.test(block)
        ? 'Yêu cầu có kinh nghiệm (経験者必須)'
        : '',
      jobDescription: block.slice(0, 600),
    };
  });
};

export const extractKyujinViaServer = async (
  file: File,
  idToken?: string | null
) => {
  const lower = file.name.toLowerCase();
  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
  };
  if (idToken) {
    headers.Authorization = `Bearer ${idToken}`;
  }

  let payload: Record<string, unknown> = { fileName: file.name };

  if (lower.endsWith('.docx')) {
    const rawText = await extractWordTextViaMammoth(file);
    payload = { rawText, fileName: file.name };
  } else if (lower.endsWith('.txt') || lower.endsWith('.csv') || file.type.startsWith('text/')) {
    const rawText = await file.text();
    payload = { rawText, fileName: file.name };
  } else if (file.type.startsWith('image/') || /\.(png|jpe?g|webp|gif)$/i.test(lower)) {
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

  const response = await fetch('/api/gemini/extract-kyujin', {
    method: 'POST',
    headers,
    body: JSON.stringify(payload),
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data?.error || `HTTP ${response.status}`);
  }
  return Array.isArray(data.items) ? data.items : [];
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
  const data = await response.json();
  if (!response.ok) {
    const err = new Error(data?.error || `HTTP ${response.status}`) as Error & { code?: string };
    err.code = data?.code;
    throw err;
  }
  return (data.candidate || {}) as Partial<Candidate>;
};

// ============================================================================
// CHAT PARSING & SMART REPLY BUILDER
// ============================================================================
export const extractPreferredAreaFromChat = (chatText = ''): string => {
  const text = String(chatText || '');
  if (!text.trim()) return '';

  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  const wishLineRx =
    /(?:ở|tại|về|sang|khu vực|tìm.*ở|muốn.*ở|希望|エリア|勤務地|方面|あたり|近く|ないですか|ありますか|tỉnh)/i;

  const wishMatches = new Set<string>();
  for (const line of lines) {
    if (wishLineRx.test(line)) {
      for (const item of PREFECTURE_PATTERNS) {
        if (item.rx.test(line)) {
          wishMatches.add(item.label);
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

  if (/kansai|関西/i.test(text)) return '関西エリア';
  if (/kanto|関東/i.test(text)) return '関東エリア';
  return '';
};

export const parseChatIntoUpdates = (chatText = '') => {
  const text = String(chatText || '').trim();
  const todayIso = new Date(Date.now() + 9 * 60 * 60 * 1000).toISOString().slice(0, 10);
  const updates: Partial<Candidate> & { memoLine: string } = {
    lastContactDate: todayIso,
    memoLine: '',
  };

  const lines = text
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);

  // Separate explicit Legal Name & Facebook Name if provided, otherwise first short line
  const legalNameM = text.match(/(?:họ\s*(?:và\s*)?tên|tên\s*khai\s*sinh|氏名|名前)[：:\s]*([^\n,;()]+)/i);
  const fbNameM = text.match(/(?:fb|facebook|tên\s*fb)[：:\s]*([^\n,;()]+)/i);

  if (legalNameM) {
    updates.name = legalNameM[1].trim();
  }
  if (fbNameM) {
    updates.fbName = fbNameM[1].trim();
  }
  if (!updates.name && !updates.fbName && lines.length > 0 && lines[0].length <= 45 && !/[:：。]/.test(lines[0])) {
    updates.fbName = lines[0];
    updates.name = lines[0].toUpperCase();
  }

  const jlptM = text.match(/\b(N[1-5]|JFT(?:\s*Basic)?)\b/i);
  if (jlptM) updates.jlpt = jlptM[1].toUpperCase();

  const phoneM = text.match(/(0[789]0[-\s]?\d{4}[-\s]?\d{4})/);
  if (phoneM) updates.phone = phoneM[1];

  // Parse YYYY-MM-DD, YYYY/MM/DD, or DD/MM/YYYY for Visa Expiry
  const expYmdM = text.match(
    /(?:在留期限|ビザ期限|hạn\s*(?:thẻ|visa|ngoại\s*kiều)?|thẻ\s*(?:hết\s*hạn|đến))[：:\s]*(\d{4}[年\/\-]\d{1,2}[月\/\-]\d{1,2})/i
  );
  const expDmyM = text.match(
    /(?:在留期限|ビザ期限|hạn\s*(?:thẻ|visa|ngoại\s*kiều)?|thẻ\s*(?:hết\s*hạn|đến))[：:\s]*(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})/i
  );
  if (expYmdM) {
    const clean = expYmdM[1].replace(/[年月]/g, '-').replace(/日/g, '');
    const parts = clean.split(/[\/\-]/);
    if (parts.length === 3) {
      updates.visaExpiry = `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
    }
  } else if (expDmyM) {
    updates.visaExpiry = `${expDmyM[3]}-${expDmyM[2].padStart(2, '0')}-${expDmyM[1].padStart(2, '0')}`;
  }

  // Current visa status
  if (/thực\s*tập\s*sinh|tts|技能実習/i.test(text)) {
    updates.visaStatus = '技能実習（Thực tập sinh）';
  } else if (/du\s*học|留学/i.test(text)) {
    updates.visaStatus = '留学（Du học sinh）';
  } else {
    const remPeriodM = text.match(
      /(?:còn|tokutei\s*(?:còn)?|thời hạn\s*(?:còn)?|残り)?\s*(\d+(?:[.,]\d+)?)\s*(?:n\b|năm|年)/i
    );
    if (remPeriodM && /(?:tokutei|特定|còn|残り|visa|ビザ|\d+[.,]\d+\s*n\b)/i.test(text)) {
      const yrs = remPeriodM[1].replace(',', '.');
      updates.visaStatus = `特定技能1号（残り約${yrs}年）`;
    } else if (/tokutei|特定技能/i.test(text)) {
      updates.visaStatus = '特定技能1号';
    }
  }

  if (/介護|特養|デイサービス|グループホーム|有料老人|shoninsha|初任者|viện de|viện có phí|kaigo/i.test(text)) {
    updates.desiredIndustry = '介護';
  } else if (/外食|レストラン|居酒屋|キッチン|ngoại thực|nhà hàng/i.test(text)) {
    updates.desiredIndustry = '外食業';
  } else if (/ドライバー|トラック|運送|運転|lái xe|vận tải/i.test(text)) {
    updates.desiredIndustry = '自動車運送業';
  } else if (/ホテル|宿泊|フロント|khách sạn/i.test(text)) {
    updates.desiredIndustry = '宿泊';
  } else if (/食品|弁当|工場|thực phẩm/i.test(text)) {
    updates.desiredIndustry = '飲食料品製造業';
  }

  const extractedArea = extractPreferredAreaFromChat(text);
  if (extractedArea) {
    updates.desiredArea = extractedArea;
  }

  const examList: string[] = [];
  if (/shoninsha|初任者/i.test(text)) examList.push('介護職員初任者研修修了');
  if (/jitsumusha|実務者/i.test(text)) examList.push('介護福祉士実務者研修修了');
  if (/kaigo fukushishi|介護福祉士/i.test(text)) examList.push('介護福祉士');
  if (/gino|技能評価試験|特定技能試験|chứng chỉ tokutei/i.test(text)) examList.push('特定技能評価試験合格');
  if (examList.length > 0) {
    updates.passedExam = examList.join('・');
  }

  const expParts: string[] = [];
  const knMatch =
    text.match(/(?:kn|kinh nghiệm|経験)\s*(\d+(?:[.,]\d+)?)\s*(?:n\b|năm|年)/i) ||
    text.match(/(\d+(?:[.,]\d+)?)\s*(?:năm|年)\s*(?:kn|kinh nghiệm|経験)/i);
  if (knMatch) {
    expParts.push(`経験${knMatch[1].replace(',', '.')}年`);
  }
  if (/kaigo năm cuối|năm cuối/i.test(text)) {
    expParts.push('Đang là Kaigo năm cuối');
  }
  const facilities: string[] = [];
  if (/viện\s*de\b|デイサービス|デイ\b/i.test(text)) facilities.push('デイサービス');
  if (/viện\s*có\s*phí|有料老人ホーム|有料/i.test(text)) facilities.push('有料老人ホーム');
  if (/tokuyou|特別養護老人ホーム|特養/i.test(text)) facilities.push('特別養護老人ホーム');
  if (/group\s*home|グループホーム|GH\b/i.test(text)) facilities.push('グループホーム');
  if (/rouken|老健/i.test(text)) facilities.push('老健');
  if (facilities.length > 0) {
    expParts.push(`施設経験：${facilities.join('・')}`);
  }
  if (expParts.length > 0) {
    updates.japanExperience = expParts.join('／');
  }

  const salaryLineM = text.match(/(?:lương|thực nhận|man|手取り|基本給)[^\n]*/i);
  if (salaryLineM) {
    updates.salaryExpectation = salaryLineM[0].trim();
  }

  const goalLineM = text.match(/(?:mục tiêu|nhu cầu|tìm công ty|chuyển đổi|nguyện vọng|希望)[^\n]*/i);
  if (goalLineM) {
    updates.conversionGoal = goalLineM[0].trim();
  }

  const mendanDateM = text.match(
    /(\d{4}[年\/\-]\d{1,2}[月\/\-]\d{1,2}|\d{1,2}[月\/]\d{1,2}日?)\s*(?:に|の)?\s*(?:面談|面接)/
  );
  if (mendanDateM) {
    const ts = parseFlexibleDateToTimestamp(mendanDateM[1], 2026);
    if (ts > 0) updates.mendanDate = formatTimestampToIso(ts);
  }

  if (/内定/.test(text)) updates.progressStatus = '内定';
  else if (/2次面接/.test(text)) updates.progressStatus = '2次面接予定';
  else if (/1次面接|面接予定|面接決ま/.test(text)) updates.progressStatus = '1次面接予定';
  else if (/面談した|面談済|面談終了/.test(text)) updates.progressStatus = '面談済み';
  else if (/面談調整|日程調整/.test(text)) updates.progressStatus = '面談調整中';

  const compM = text.match(/([^\s、。,\n]+(?:株式会社|有限会社|医療法人|社会福祉法人))/);
  if (compM) updates.proposedCompany = compM[1];

  // Zero truncation: preserve 100% of original text verbatim!
  updates.negotiationLog = `[${todayIso} Ghi chú nguyên văn]\n${text}`;
  updates.memoLine = `[${todayIso} Nhật ký trao đổi]\n${text}`;
  return updates;
};

export const buildSmartReplyMessages = (
  person: Partial<Candidate> = {},
  scenario: ReplyScenario = 'first_mendan',
  kyujins: KyujinCompany[] = []
) => {
  const name = person.fbName || person.name || 'Bạn';
  const missing = getMissingInfoForCandidate(person);
  const matches = findMatchingKyujinsForCandidate(person, kyujins);
  const topMatch = matches[0];
  const proposedComp =
    getEffectiveProposedCompany(person) || topMatch?.company?.companyName || '企業様';
  const mendanInfo = extractNormalizedMendanInfo(person);
  const dateText = mendanInfo.dateLabel || person.mendanDate || '近日中';

  const missingVi =
    missing.length > 0
      ? `\nĐể chuẩn bị hồ sơ 履歴書 và chọn đơn tốt nhất cho bạn, bạn cho mình xin thêm các thông tin sau nhé:\n${missing
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
      vi: `Chào ${name}! Để mình lọc chính xác 3 đơn hàng lương tốt và hỗ trợ làm Rirekisho gửi công ty cho bạn nhanh nhất, bạn bổ sung giúp mình mấy thông tin còn thiếu sau nhé:\n${listVi}\n\nBạn nhắn ngắn gọn hoặc chụp ảnh mặt trước thẻ ngoại kiều (che số thẻ nếu muốn) gửi mình cũng được nha!`,
      jp: `${name}さん、こんにちは！ご希望に合う求人のご提案と履歴書作成を進めるため、不足している以下の情報をお送りいただけますでしょうか。\n${listJp}\n\nメッセージまたは在留カードのお写真でお知らせいただければすぐに対応いたします！`,
      en: `Hello ${name}! To match the best job openings and prepare your resume quickly, could you please share the missing details below:\n${listEn}`,
    };
  }

  if (scenario === 'post_interview_followup') {
    return {
      vi: `Chào ${name}! Buổi phỏng vấn với ${proposedComp} (${dateText}) vừa rồi thế nào bạn ơi?\nBạn thấy không khí công ty, nội dung công việc và mức lương phía công ty chia sẻ có phù hợp với kỳ vọng của bạn không? Có điều gì băn khoăn cứ nhắn thật lòng cho mình để mình trao đổi thêm với công ty hoặc tính phương án tốt nhất cho bạn nhé!`,
      jp: `${name}さん、お疲れ様でした！${proposedComp}との面接（${dateText}）はいかがでしたでしょうか？\n業務内容や給与条件・職場の雰囲気など、ご感想や気になる点がございましたら遠慮なくお聞かせください。結果連絡や次のステップに向けてしっかりサポートいたします！`,
      en: `Hello ${name}! How did your interview with ${proposedComp} (${dateText}) go? Please let me know your thoughts on the role, salary, and workplace so I can support your next steps!`,
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
      vi: `Chào ${name}! Dựa theo hồ sơ của bạn, mình gửi 3 đơn phù hợp nhất hiện tại (ẩn danh tên công ty trước khi phỏng vấn):\n\n${jobBlockVi}\n${missingVi}\nNếu bạn ưng đơn số mấy, nhắn lại cho mình để mình gửi lịch phỏng vấn sớm nhất nhé!`,
      jp: `${name}さん、こんにちは！ご希望条件に合う求人3件をご案内いたします（選考前のため企業名は非公開です）。\n\n${jobBlockJp}\n${missingJp}\n気になる求人がございましたら番号をお知らせください！`,
      en: `Hello ${name}! Here are the top matching openings for your profile:\n\n${jobBlockEn}\n${missingEn}\nPlease let me know which one interests you!`,
    };
  }

  if (scenario === 'remind_interview') {
    return {
      vi: `Chào ${name}! Mình nhắn nhắc bạn lịch hẹn Mendan / Phỏng vấn sắp tới:\n・Thời gian: ${dateText}\n・Nội dung / Công ty: ${proposedComp}\n\nBạn nhớ kiểm tra đường truyền mạng, trang phục gọn gàng và vào trước giờ hẹn 5–10 phút nhé. Nếu cần hỗ trợ luyện tập câu hỏi phỏng vấn thì nhắn mình ngay nha!`,
      jp: `${name}さん、こんにちは！面談・面接のリマインド連絡です。\n・日時：${dateText}\n・対象企業／内容：${proposedComp}\n\n当日は開始5〜10分前にご準備をお願いいたします。事前練習やご不明点があればいつでもご連絡ください！`,
      en: `Hello ${name}! This is a friendly reminder for your upcoming Mendan / Interview:\n・Date/Time: ${dateText}\n・Company / Details: ${proposedComp}\n\nPlease be ready 5–10 minutes early in neat attire. Let me know if you need any interview prep!`,
    };
  }

  if (scenario === 'followup_silent') {
    return {
      vi: `Chào ${name}! Lâu rồi chưa thấy bạn phản hồi, không biết tình hình công việc và tìm việc Tokutei của bạn dạo này thế nào rồi?\nHiện bên mình mới về thêm nhiều đơn lương tốt và có thưởng (賞与あり). Nếu bạn vẫn đang tìm việc hoặc muốn tham khảo thêm đơn mới thì nhắn lại cho mình nhé!`,
      jp: `${name}さん、こんにちは！その後お仕事探しの状況はいかがでしょうか？\n現在、賞与あり・好条件の特定技能求人も新しく入っております。まだ転職・就職をご検討中でしたら、ぜひお気軽にご返信ください！`,
      en: `Hello ${name}! How has your job search been going lately?\nWe have received several new Tokutei Ginou openings with good salary and bonuses. If you are still looking, feel free to reply anytime!`,
    };
  }

  return {
    vi: `Chào ${name}! Cảm ơn bạn đã liên hệ tư vấn việc làm Tokutei Ginou.\nBên mình hỗ trợ miễn phí 100% từ khâu làm 履歴書 đến luyện phỏng vấn và xin visa.\n${missingVi}\n\nHôm nay hoặc ngày mai khung giờ nào bạn tiện nghe điện thoại/gọi video Mendan khoảng 15 phút để mình tư vấn đơn phù hợp nhất cho bạn?`,
    jp: `${name}さん、こんにちは！特定技能のお仕事相談にお問い合わせいただきありがとうございます。\n履歴書作成から面接対策・ビザ申請まで完全無料でサポートしております。\n${missingJp}\n\n本日または明日で、15分ほどお電話・オンライン面談が可能なお時間帯はございますでしょうか？`,
    en: `Hello ${name}! Thank you for reaching out regarding Tokutei Ginou job support.\nWe provide 100% free support from resume preparation to interview coaching and visa procedures.\n${missingEn}\n\nWhat time today or tomorrow would be convenient for a quick 15-minute Mendan call?`,
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
      <table>
        <tr>
          <th style="width:22%">フリガナ</th>
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
        <tr>
          <th>現住所</th>
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
