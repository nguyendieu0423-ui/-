import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import mammoth from 'mammoth';
import type {
  Candidate,
  ExcelAnchor,
  KyujinCompany,
  KyujinMatchEvaluation,
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
    rx: /\btokyo\b|東京都|(?:^|[^京都])東京|23区|新宿区|渋谷区|杉並区|立川市|足立区|八王子市|世田谷区|品川区|大田区|中野区|練馬区|板橋区|豊島区|港区|中央区|千代田区/i,
    label: '東京都（Tokyo）',
    shortLabel: '東京都',
    keys: ['東京都', '東京', 'tokyo'],
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

export const evaluateCandidateKyujinMatches = (
  cand: Partial<Candidate> = {},
  kyujinsList: KyujinCompany[] = []
): KyujinMatchEvaluation => {
  const candInd = String(cand.desiredIndustry || '').trim();
  const candArea = String(cand.desiredArea || '').trim();
  const candGoal = String(cand.conversionGoal || '').trim();
  const candExp = String(cand.japanExperience || '').trim();
  const candPassedExam = String(cand.passedExam || '').trim();
  const candNat = String(cand.nationality || '').trim();
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
    (comp.areaBreakdown || []).forEach((sl) => {
      if (sl.recruitmentStatus === '募集終了') return;

      if (
        candNat &&
        Array.isArray(comp.excludedNationalities) &&
        comp.excludedNationalities.includes(candNat)
      ) {
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
        reasons.push(`✅ Đúng ngành ${slotInd || candInd}`);
      } else {
        rawScore += 15;
      }

      // Employment Type bonus
      if (mentionsBaito && slotIsBaito) {
        rawScore += 5;
        reasons.push('✅ Đúng hình thức Baito');
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
          reasons.push(`✅ Khớp ${matchedPref.shortLabel}`);
        } else if (candRegions.length > 0 && slotAreaPrefs.length > 0) {
          const matchedViaRegion = slotAreaPrefs.find((sp) =>
            candRegions.some((reg) => reg.prefectures.includes(sp.shortLabel))
          );
          if (matchedViaRegion) {
            areaMatched = true;
            rawScore += 22;
            reasons.push(`✅ Khớp ${matchedViaRegion.shortLabel}`);
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
            reasons.push(`✅ Khớp ${hit}`);
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
        reasons.push('✅ Không yêu cầu JLPT');
      } else if (reqJlpt === '要確認' || reqJlptScore < 0) {
        rawScore += 5;
      } else if (candJlptScore > 0 && candJlptScore >= reqJlptScore) {
        rawScore += 15;
        reasons.push(`✅ Đạt ${cand.jlpt}`);
      } else if (candJlptScore > 0 && candJlptScore < reqJlptScore) {
        rawScore -= 20;
      }

      // 5. Experience Requirement
      const expReq = String(sl.experienceRequirement || '');
      const hasCandExp = Boolean(candExp || candPassedExam);
      if (/未経験|không yêu cầu/i.test(expReq)) {
        rawScore += 10;
        reasons.push('✅ Không yêu cầu kinh nghiệm');
      } else if (/経験者|có kinh nghiệm/i.test(expReq)) {
        if (hasCandExp) {
          rawScore += 10;
          reasons.push('✅ Đáp ứng kinh nghiệm');
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
            reasons.push('✅ Đạt mức lương kỳ vọng');
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
            reasons.push('✅ Đạt mức lương kỳ vọng');
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
          reasons.push('✅ Có hỗ trợ nhà ở / KTX');
        } else if (sl.bonus === '賞与あり') {
          rawScore += 5;
          reasons.push('🎁 Có thưởng');
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

  const hasExactIndustryAndAreaMatch =
    hasCandidateArea && results.some((r) => Boolean(r.areaMatched));

  let noticeMessage = '';
  if (hasCandidateIndustry) {
    if (results.length === 0) {
      noticeMessage = `Chưa có đơn đúng ngành ${candInd}${
        hasCandidateArea ? ` tại ${candArea}` : ''
      } trong kho Tin tuyển dụng.`;
    } else if (hasCandidateArea && !hasExactIndustryAndAreaMatch) {
      noticeMessage = `Chưa có đơn đúng ngành ${candInd} tại ${candArea} — hiển thị ${results.length} đơn cùng ngành ở khu vực khác để tham khảo.`;
    }
  }

  return {
    matches: results.slice(0, 5),
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
  const rawCand = (data.candidate || {}) as Partial<Candidate> & { fbUrl?: string };
  return {
    ...rawCand,
    fbProfileUrl: rawCand.fbProfileUrl || rawCand.fbUrl || '',
    visaExpiry: rawCand.visaExpiry ? normalizeExtractedDateYmd(rawCand.visaExpiry) : '',
    birthDate: rawCand.birthDate ? normalizeExtractedDateYmd(rawCand.birthDate) : '',
  };
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

    const turnM = line.match(/^([^\n:：]{2,40})[:：]\s*(.+)$/);
    if (turnM) {
      const label = turnM[1].trim();
      const body = turnM[2].trim();
      const isFieldLabel =
        /^(?:họ\s*(?:và\s*)?tên|tên\s*(?:khai\s*sinh|fb|facebook)|fb|facebook|sđt|sdt|điện\s*thoại|tel|phone|email|mail|visa|hạn\s*(?:thẻ|visa)|trình\s*độ|jlpt|tiếng\s*nhật|khu\s*vực|tỉnh|ngành|kinh\s*nghiệm|chứng\s*chỉ|lương|mục\s*tiêu|nhu\s*cầu|ghi\s*chú|ngày\s*sinh|tuổi|quốc\s*tịch|địa\s*chỉ|link|氏名|名前|在留資格|在留期限|希望|エリア|給料|経験|資格)$/i.test(
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

  // 2. Legal Name vs Facebook Display Name (never invent or copy staff name)
  const legalNameM = text.match(
    /(?:họ\s*(?:và\s*)?tên(?:\s*khai\s*sinh)?|tên\s*khai\s*sinh|full\s*name|氏名|本名|名前)[：:\s]+([^\n,;()]+)/i
  );
  const fbNameM = text.match(
    /(?:tên\s*(?:nick\s*)?(?:fb|facebook|messenger)|nick\s*(?:fb|facebook)|fb|facebook)[：:\s]+([^\n,;()]+)/i
  );
  const inlineParenFbM = text.match(
    /^([A-ZÀ-Ỹ][^\n()]{3,45})\s*\((?:fb|facebook)[：:\s]*([^)\n]+)\)/im
  );

  if (inlineParenFbM) {
    updates.name = inlineParenFbM[1].trim();
    updates.fbName = inlineParenFbM[2].trim();
  }
  if (legalNameM && !updates.name) {
    const candidateName = legalNameM[1].trim();
    if (candidateName && !CONSULTANT_CONTENT_RX.test(candidateName)) {
      updates.name = candidateName;
    }
  }
  if (fbNameM && !updates.fbName) {
    const candidateFb = fbNameM[1].trim();
    if (
      candidateFb &&
      !/^https?:\/\//i.test(candidateFb) &&
      !CONSULTANT_CONTENT_RX.test(candidateFb)
    ) {
      updates.fbName = candidateFb;
    }
  }

  // Check self-introduction patterns in candidate turns only: "em tên là ..." / "mình tên ..."
  if (!updates.name) {
    const selfIntroM = candText.match(
      /(?:em|mình|tôi)\s*(?:tên\s*(?:là|thật\s*là)?|là)\s+([A-ZÀ-Ỹ][a-zà-ỹ]+(?:\s+[A-ZÀ-Ỹ][a-zà-ỹ]+){1,4}|[A-Z]{2,}(?:\s+[A-Z]{2,}){1,4})/
    );
    if (selfIntroM) {
      const introName = selfIntroM[1].trim();
      if (!/^(?:Em|Mình|Tôi|Bạn|Anh|Chị)\b/i.test(introName)) {
        updates.name = introName;
      }
    }
  }

  // Fallback for FB Nickname ONLY (do NOT promote to legal name)
  if (!updates.fbName && detectedSpeakerFbName) {
    updates.fbName = detectedSpeakerFbName;
  } else if (
    !updates.fbName &&
    !updates.name &&
    cleanedLines.length > 0 &&
    cleanedLines[0].length <= 40 &&
    !/[:：。?!？！]/.test(cleanedLines[0]) &&
    !CONSULTANT_CONTENT_RX.test(cleanedLines[0]) &&
    !CANDIDATE_CONTENT_RX.test(cleanedLines[0]) &&
    !/^(?:em|anh|chị|bạn|chào|alo|hello|hi|xin chào)\b/i.test(cleanedLines[0])
  ) {
    updates.fbName = cleanedLines[0];
  }

  // 2b. Default Nationality & Channel & Initial Progress Status
  const hasVietnamese =
    /[àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ]|\b(?:dhs|em|mình|anh|chị|tìm|chuyển|visa|tốt nghiệp|năm sau|kinh nghiệm)\b/i.test(
      text
    );
  if (hasVietnamese) {
    updates.nationality = 'ベトナム';
  }
  if (/fanpage|page\s*công\s*ty|fb\s*công\s*ty|公式/i.test(text)) {
    updates.channel = 'company';
  } else {
    updates.channel = 'personal';
  }
  updates.progressStatus = '新規相談';

  // 3. Language (JLPT / JFT / 初級) — NEVER promote upcoming target exam (e.g. "sắp tới thi N3") as current level!
  const upcomingExamM = candText.match(
    /(?:sắp\s*(?:tới)?|chuẩn\s*bị|đang\s*ôn|dự\s*định|tháng\s*\d{1,2}(?:\/\d{4})?)\s*(?:sẽ\s*)?thi\s*(?:jlpt\s*)?(N[1-5])/i
  );
  const upcomingExamNote = upcomingExamM
    ? `Sắp tới thi JLPT ${upcomingExamM[1].toUpperCase()}`
    : '';
  const candTextCurrentJlpt = candText.replace(
    /(?:sắp\s*(?:tới)?|chuẩn\s*bị|đang\s*ôn|dự\s*định|tháng\s*\d{1,2}(?:\/\d{4})?)\s*(?:sẽ\s*)?thi\s*(?:jlpt\s*)?N[1-5]/gi,
    ''
  );
  const jlptM =
    candTextCurrentJlpt.match(/\b(N[1-5]|JFT(?:\s*Basic)?)\b/i) ||
    text
      .replace(
        /(?:sắp\s*(?:tới)?|chuẩn\s*bị|đang\s*ôn|dự\s*định|tháng\s*\d{1,2}(?:\/\d{4})?)\s*(?:sẽ\s*)?thi\s*(?:jlpt\s*)?N[1-5]/gi,
        ''
      )
      .match(/\b(N[1-5]|JFT(?:\s*Basic)?)\b/i);
  if (jlptM) {
    updates.jlpt = jlptM[1].toUpperCase().replace(/\s+/g, ' ');
  } else if (/初級|sơ\s*cấp/i.test(candText)) {
    updates.jlpt = '初級';
  } else if (/giao tiếp cơ bản|日常会話/i.test(candText)) {
    updates.jlpt = '日常会話';
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

  // 6. Visa Expiry (careful DD/MM/YYYY vs YYYY/MM/DD)
  const expPrefix =
    '(?:在留期限|ビザ期限|hạn\\s*(?:thẻ\\s*ngoại\\s*kiều|thẻ|visa|ngoại\\s*kiều)?|thẻ\\s*(?:ngoại\\s*kiều\\s*)?(?:hết\\s*hạn|đến)|hết\\s*hạn)(?:\\s*ngày)?';
  const expYmdM = candText.match(
    new RegExp(`${expPrefix}[：:\\s]*(\\d{4}[年\\/\\-]\\d{1,2}[月\\/\\-]\\d{1,2}日?)`, 'i')
  );
  const expDmyM = candText.match(
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
    if (remPeriodM && /(?:tokutei|特定|còn|残り|visa|ビザ|\d+[.,]\d+\s*n\b)/i.test(candText)) {
      const yrs = remPeriodM[1].replace(',', '.');
      updates.visaStatus = `特定技能1号（残り約${yrs}年）`;
    } else if (/tokutei|特定技能/i.test(candText)) {
      updates.visaStatus = '特定技能1号';
    }
  }

  // 8. Desired Industry (evaluate candidate text first)
  if (/介護|特養|デイサービス|グループホーム|有料老人|shoninsha|初任者|viện de|viện có phí|kaigo|điều\s*dưỡng|hộ\s*lý/i.test(candText)) {
    updates.desiredIndustry = '介護';
  } else if (/外食|レストラン|居酒屋|キッチン|ngoại\s*thực|nhà\s*hàng|quán\s*ăn|bếp/i.test(candText)) {
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
  if (/gino|技能評価試験|特定技能試験|chứng\s*chỉ\s*tokutei|đỗ\s*tay\s*nghề|senmon/i.test(candText)) {
    examList.push('特定技能評価試験合格');
  }
  if (/bằng\s*lái|gplx|免許/i.test(candText)) examList.push('運転免許');
  if (updates.jlpt && /^N[1-5]$/.test(updates.jlpt)) {
    examList.push(`JLPT ${updates.jlpt}`);
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

  // 11. Salary expectation & housing / night-shift lines (preserve all matching lines)
  const salaryLines = candidateLines.filter((l) =>
    /(?:lương|thực\s*nhận|cơ\s*bản|trừ\s*nhà|tiền\s*nhà|ký\s*túc|ca\s*đêm|yakin|jikyu|時給|手取り|基本給|総支給|寮|夜勤|\d+\s*(?:man|vạn|yên|円|万円))/i.test(
      l
    )
  );
  if (salaryLines.length > 0) {
    updates.salaryExpectation = Array.from(new Set(salaryLines)).join(' | ');
  }

  // 12. Conversion Goal / Need
  const goalLines = candidateLines.filter((l) =>
    /(?:mục\s*tiêu|nhu\s*cầu|tìm\s*(?:công\s*ty|đơn|việc|baito)|chuyển\s*(?:đổi|visa|việc|công\s*ty|sang\s*tokutei)|nguyện\s*vọng|希望|転職)/i.test(
      l
    )
  );
  if (goalLines.length > 0) {
    updates.conversionGoal = Array.from(new Set(goalLines)).join(' | ');
  } else if (isDhs) {
    updates.conversionGoal = 'Du học sinh tìm đơn chuyển đổi sang visa Tokutei Ginou';
  }

  // 13. Mendan / Interview (YYYY-MM-DD HH:mm) / Status / Proposed Company
  const mendanDateM = text.match(
    /(?:mendan|面談)[^\n\d]{0,15}(\d{4}[年\/\-]\d{1,2}[月\/\-]\d{1,2}|\d{1,2}[月\/]\d{1,2}日?)(?:[^\n\d]{0,12}(\d{1,2}[:h]\d{2}))?/i
  ) || text.match(
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

  const interviewDateM = text.match(
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

  if (/内定/.test(text)) updates.progressStatus = '内定';
  else if (/2次面接/.test(text)) updates.progressStatus = '2次面接予定';
  else if (/1次面接|面接予定|面接決ま/.test(text)) updates.progressStatus = '1次面接予定';
  else if (/面談した|面談済|面談終了/.test(text)) updates.progressStatus = '面談済み';
  else if (/面談調整|日程調整/.test(text)) updates.progressStatus = '面談調整中';

  const compM = text.match(/([^\s、。,\n]+(?:株式会社|有限会社|医療法人|社会福祉法人))/);
  if (compM) updates.proposedCompany = compM[1];

  // Zero truncation: preserve 100% of original text verbatim!
  const missingLegalNote =
    !updates.name && updates.fbName
      ? `[Chưa có họ tên khai sinh — mới ghi nhận tên FB: ${updates.fbName}]\n`
      : '';
  updates.negotiationLog = `[${todayIso} Ghi chú nguyên văn]\n${missingLegalNote}${text}`;
  updates.memoLine = `[${todayIso} Nguyên văn đoạn chat]\n${missingLegalNote}${text}`;
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
