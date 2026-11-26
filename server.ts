import crypto from 'crypto';
import express from 'express';
import path from 'path';
import { fileURLToPath } from 'url';
import dotenv from 'dotenv';
import { GoogleGenAI, Type } from '@google/genai';
import {
  optionalAuth,
  requireAuth,
  type AuthRequest,
} from './src/middleware/auth.ts';
import { getPrimaryCrmUser } from './src/db/users.ts';
import {
  appendUserChatExchange,
  clearUserChatHistory,
  getUserChatHistory,
  loadUserCrmData,
  saveUserCrmSnapshot,
  updateCandidateCalendarEvent,
  updateResignationCalendarEvent,
  upsertMessengerCandidate,
  type MessengerExtractedProfile,
} from './src/db/crmRepository.ts';
import type { Candidate, ResignationRecord } from './src/types/crm.ts';
import {
  extractAnnualSalaryText,
  extractEstimatedNetSalaryText,
  extractKyujinSalaryInfo,
  sanitizeExcelFloatText,
  stripKyujinFormBoilerplate,
} from './src/utils/salaryParser.ts';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = 3000;
const MODEL_NAME = 'gemini-4-hydrogen';

function getGeminiClient(): GoogleGenAI {
  const apiKey = (process.env.GEMINI_API_KEY || '').trim();
  if (!apiKey) {
    const err = new Error('GEMINI_API_KEY is not configured in environment secrets.') as Error & {
      code?: string;
    };
    err.code = 'MISSING_GEMINI_KEY';
    throw err;
  }
  return new GoogleGenAI({
    apiKey,
    httpOptions: {
      headers: {
        'User-Agent': 'aistudio-build',
      },
    },
  });
}

function safeParseGeminiJson<T>(rawText: string | undefined, fallback: T): T {
  const cleaned = String(rawText || '')
    .trim()
    .replace(/^```(?:json)?\s*/i, '')
    .replace(/\s*```$/i, '')
    .trim();
  if (!cleaned) return fallback;
  try {
    return JSON.parse(cleaned) as T;
  } catch {
    const firstObj = cleaned.indexOf('{');
    const firstArr = cleaned.indexOf('[');
    const startIdx =
      firstObj === -1
        ? firstArr
        : firstArr === -1
        ? firstObj
        : Math.min(firstObj, firstArr);
    if (startIdx >= 0) {
      const endChar = cleaned[startIdx] === '[' ? ']' : '}';
      const endIdx = cleaned.lastIndexOf(endChar);
      if (endIdx > startIdx) {
        try {
          return JSON.parse(cleaned.slice(startIdx, endIdx + 1)) as T;
        } catch {
          return fallback;
        }
      }
    }
    return fallback;
  }
}

const KYUJIN_EXTRACTION_PROMPT = `あなたは日本の人材紹介会社（特定技能・正社員・アルバイト/パート）向けの求人票精密データ抽出AIです。
添付されたPDF・画像（複数ページ含む）またはテキストデータを確認し、求人票が複数含まれる場合や、1つの求人票に複数の施設・勤務先・店舗一覧表がある場合は、必ず施設・勤務地・枠ごとに1件ずつ分離してJSON配列で返してください。

【重要ルール：分野・雇用形態（Baito含む）・待遇の精密分類】
1. companyName: 「企業名・法人名・店舗運営会社名」（後ろの「代表者〜」は削除）。
2. jobCode: 「案件番号」または「施設名・店舗名」。
3. sector (Lĩnh vực lớn): 以下の8分類から最も適切なものを1つ選択:
   - "Kaigo / Điều dưỡng"
   - "Thực phẩm & Dịch vụ ăn uống"
   - "Sản xuất & Nhà máy"
   - "Xây dựng & Cơ khí"
   - "Khách sạn & Du lịch"
   - "Vệ sinh tòa nhà & Nông nghiệp"
   - "Vận tải & Logistics"
   - "Bán lẻ / Dịch vụ Baito"
4. industry (Ngành nghề cụ thể): 介護, 外食業, 飲食料品製造業, 宿泊, 自動車運送業, ビルクリーニング, 建設, 自動車整備, 農業, 素形材・産業機械・電気電子情報関連製造業, コンビニ・小売（アルバイト） 等。
5. employmentType (Hình thức tuyển dụng): 以下の3分類から必ず判定:
   - "Tokutei Ginou (特定技能)"
   - "Nhân viên chính thức / Chuyển đổi visa (正社員・登用)"
   - "Baito / Làm thêm (アルバイト・パート)"
6. facilityType: 介護の場合の施設形態（デイサービス、グループホーム、有料老人ホーム、特別養護老人ホーム、老健、訪問介護）。
7. area: 勤務地を「都道府県＋区・市」レベルで正確に記載。
8. shiftTime: 勤務時間・シフト枠。就業時間欄に複数のシフト枠（例: ①08:00〜17:00 / ②09:00〜18:00 / ③11:00〜20:00 / ④14:00〜22:00）が並んでいる場合は1つにまとめず、番号・時刻をすべて原文どおり列挙すること。
9. shiftType: 勤務形態を「日勤のみ」「シフト制」「夜勤あり」から正確に判定（22:00以降の深夜・夜勤シフトがなければ「日勤のみ」または「シフト制（夜勤なし）」として扱う）。
10. holidays: 年間休日日数・休日形態（例: 110日/年・月7〜10日、完全週休2日制 等）。
11. overtimeHours: 残業時間・みなし残業・月平均残業時間（例: 月平均10時間、なし）。
12. salary: 給与・月給・時給・基本給・想定月収・賃金・報酬・給与形態・総支給から必ず抽出（例: 220,000円, 22万円, 22.5万円, 20万～25万円, 220k, 月給215,000円（夜勤手当含む））。時給の場合は「時給 1,150円〜（交通費支給）」のように記載。金額の記載がある場合は絶対に「給与要確認」にしないこと。浮動小数点誤差（22.999999等）は整数に丸める。
13. nightShiftAllowance: 夜勤手当（例: 6,000円/回、なし）。
14. estimatedNetSalary: 手取り目安・控除後の記載があればそのまま記載（例: 約185,000円）。
15. annualSalary: 想定年収・モデル年収の記載があれば記載（例: 320万円〜380万円）。
16. hourlyWage: 時給（アルバイト・パート、または10,000円未満の時給表記がある場合: 1,250円 / 1,150円〜（交通費支給） / 1,200円～1,500円）。
17. commuteAllowance: 交通費・通勤手当（例: 全額支給 / 月上限20,000円）。
18. shiftsPerWeek: 週の勤務日数・時間（例: 週3〜5日 / 週28時間以内）。
19. recruitCount: 募集人数（例: 2名、若干名）。
20. interviewProcess: 選考フロー・面接回数（例: 書類選考 → 面接1〜2回）。
21. specialBenefits: 福利厚生・待遇・各種手当・退職金・社会保険等（「社会保険加入事業者か〜」「(例)〜」「※1〜」等の求人票フォーム注記定型文は絶対に除去すること）。
22. minJlpt: N1, N2, N3, N4, N5, 日常会話, 不問, 要確認（「N4以上」→「N4」のように基準級に正規化）。
23. housingSupport: 「自社寮あり」「借上げ宿舎あり」「寮・社宅あり」「提供無」「要確認」。
24. bonus / bonusDetail: 「賞与あり」「賞与なし」「要確認」＋詳細（例: 年2回 計2.5ヶ月分）。重要：チェックボックス☑・✔・●や「無」「賞与なし」「退職金・賞与無」が指定されている場合は、選択肢に「賞与あり」の文字が並んでいても必ず「賞与なし」を優先すること。
25. experienceRequirement: 必ず「Không yêu cầu kinh nghiệm (未経験OK)」「Yêu cầu có kinh nghiệm (経験者必須)」「Ưu tiên có kinh nghiệm (経験者歓迎)」の3分類から選択。
26. experienceDetail: 必須資格や歓迎経験の詳細。
27. jobDescription: 業務内容（業務の特徴・具体的な作業内容を省略せず詳しく記載。フォーム定型文や社内メモは除去）。
28. excludedNationalities: 採用不可・NGと明記された国籍の配列（例: ["ネパール"], ["ベトナム"]）。記載がなければ []。
29. excludedGenders: 採用不可・NGと明記された性別の配列（例: "男性NG" や "女性限定" なら ["男性"]、"女性NG" や "男性限定" なら ["女性"]）。記載がなければ []。
30. requiresInJapan: "国内在住限定"、"日本在住のみ"、"在日限定"、"海外NG"、"海外招へい 不可" などの記載があれば true、それ以外は false。
31. requiresTokutei: 特定技能（または特定技能試験合格）が必須であれば true、それ以外は false。
32. requiredConditions: 必須応募条件の配列（例: ["特定技能", "日本国内在住", "JLPT N4", "介護初任者研修"]）。
33. ngConditions: 除外・NG条件の配列（例: ["ネパール人NG", "男性NG", "海外NG", "タトゥーNG"]）。`;

const RIREKISHO_EXTRACTION_PROMPT = `あなたは日本の特定技能人材紹介会社の履歴書データ抽出AIです。
提供された履歴書ファイル（PDF/画像/Excelテキスト/Wordテキスト）または応募者とのチャット履歴から、2カラム標準履歴書の全項目を正確に抽出してJSONで返してください。
不明な項目は空文字 "" または空配列 [] にしてください（推測で捏造しないこと）。

抽出対象フィールド:
- furigana: フリガナ（カタカナ）
- fullName: 氏名
- birthDate: 生年月日（YYYY-MM-DD または YYYY年MM月DD日）
- gender: 性別（男性 / 女性）
- postalCode: 郵便番号（000-0000）
- address: 現住所
- phone: 電話番号
- email: メールアドレス
- nationality: 国籍（ベトナム / ミャンマー / インドネシア / フィリピン / ネパール 等）
- visaExpiry: 在留期限（YYYY-MM-DD）
- visaType: 在留資格（特定技能1号 / 技能実習2号 / 留学 等）
- educationList: 学歴の配列 [{ "year": "2018", "month": "09", "content": "〇〇高等学校 卒業" }]
- workHistoryList: 職歴の配列 [{ "year": "2022", "month": "04", "content": "株式会社〇〇 入社（介護業務）" }]
- qualificationsList: 免許・資格の配列 [{ "year": "2024", "month": "07", "content": "日本語能力試験 N3 合格" }]
- nativeLanguage: 母国語
- spokenLanguages: 使用可能言語
- speakingLevel: 日本語能力CA評価（話す）
- listeningLevel: 日本語能力CA評価（聞く）
- motivation: 志望動機
- selfPr: 自己PR
- allergies: アレルギー（例: 特になし）
- dependentsCount: 扶養家族数(配偶者を除く)（例: 0人）
- hasSpouse: 配偶者（有 / 無）
- spouseSupportObligation: 配偶者の扶養義務（有 / 無）
- wishes: 本人希望記入欄（希望職種・勤務地・入社可能時期など）`;

interface ClientPart {
  text?: string;
  inlineData?: {
    mimeType: string;
    data: string;
  };
}

const ALLOWED_MIME_TYPES = new Set([
  'application/pdf',
  'image/png',
  'image/jpeg',
  'image/jpg',
  'image/webp',
  'image/gif',
]);

const MESSENGER_EXTRACTION_PROMPT = `Bạn là AI trích xuất thông tin ứng viên Tokutei Ginou (特定技能) từ tin nhắn Facebook Messenger.
Hãy đọc tin nhắn của ứng viên (tiếng Việt, tiếng Nhật hoặc tiếng Anh) và trích xuất chính xác các trường thông tin sau dưới dạng JSON.
QUY TẮC:
- Chỉ trích xuất thông tin thực sự có trong tin nhắn; không suy đoán hay tự bịa. Trường nào không đề cập hãy trả về chuỗi rỗng "".
- name: Họ tên ứng viên (chỉ khi ứng viên tự xưng tên hoặc ghi rõ họ tên).
- jlpt: Trình độ tiếng Nhật (N1, N2, N3, N4, N5, 日常会話 hoặc "").
- phone: Số điện thoại liên hệ (giữ nguyên số hoặc định dạng sạch).
- desiredIndustry: Ngành nghề mong muốn (ví dụ: 介護, 外食, 宿泊, 飲食料品製造, 自動車運送業, 農業, 建築, ビルクリーニング, hoặc tên ngành họ nhắc tới).
- desiredArea: Khu vực/tỉnh thành mong muốn tại Nhật (ví dụ: 東京都（Tokyo）, 大阪府（Osaka）, 愛知県（Aichi）, 神奈川県（Kanagawa）, v.v.).
- visaExpiry: Hạn thẻ ngoại kiều / hạn visa nếu có đề cập (ưu tiên YYYY-MM-DD hoặc nguyên văn ngắn gọn).`;

interface RawBodyRequest extends express.Request {
  rawBody?: Buffer;
}

function timingSafeStringEqual(a: string, b: string): boolean {
  const hashA = crypto.createHash('sha256').update(a, 'utf8').digest();
  const hashB = crypto.createHash('sha256').update(b, 'utf8').digest();
  return crypto.timingSafeEqual(hashA, hashB);
}

function verifyMetaSignature(
  rawBody: Buffer | undefined,
  signatureHeader: string | string[] | undefined,
  appSecret: string
): boolean {
  if (!rawBody || !appSecret || typeof signatureHeader !== 'string') {
    return false;
  }
  const trimmedSig = signatureHeader.trim();
  if (!trimmedSig.startsWith('sha256=')) {
    return false;
  }
  const expected = `sha256=${crypto
    .createHmac('sha256', appSecret)
    .update(rawBody)
    .digest('hex')}`;
  const sigBuf = Buffer.from(trimmedSig, 'utf8');
  const expBuf = Buffer.from(expected, 'utf8');
  if (sigBuf.length !== expBuf.length) {
    return false;
  }
  return crypto.timingSafeEqual(sigBuf, expBuf);
}

async function extractMessengerCandidateWithGemini(
  messageText: string
): Promise<MessengerExtractedProfile> {
  const ai = getGeminiClient();
  const response = await ai.models.generateContent({
    model: MODEL_NAME,
    contents: `Tin nhắn từ ứng viên Messenger:\n"""\n${messageText.slice(0, 4000)}\n"""`,
    config: {
      systemInstruction: MESSENGER_EXTRACTION_PROMPT,
      responseMimeType: 'application/json',
      responseSchema: {
        type: Type.OBJECT,
        properties: {
          name: { type: Type.STRING },
          jlpt: { type: Type.STRING },
          phone: { type: Type.STRING },
          desiredIndustry: { type: Type.STRING },
          desiredArea: { type: Type.STRING },
          visaExpiry: { type: Type.STRING },
        },
      },
    },
  });

  const rawJson = (response.text ?? '{}').trim();
  const parsed = JSON.parse(rawJson) as Record<string, unknown>;
  return {
    name: typeof parsed.name === 'string' ? parsed.name : '',
    jlpt: typeof parsed.jlpt === 'string' ? parsed.jlpt : '',
    phone: typeof parsed.phone === 'string' ? parsed.phone : '',
    desiredIndustry:
      typeof parsed.desiredIndustry === 'string' ? parsed.desiredIndustry : '',
    desiredArea: typeof parsed.desiredArea === 'string' ? parsed.desiredArea : '',
    visaExpiry: typeof parsed.visaExpiry === 'string' ? parsed.visaExpiry : '',
  };
}

const CANDIDATE_EXTRACTION_PROMPT = `Bạn là AI chuyên gia bóc tách dữ liệu ứng viên cho hệ thống Tokutei CRM (giới thiệu việc làm tại Nhật Bản).
Thời điểm hiện tại: năm 2026 (tháng 11/2026).

NHIỆM VỤ CỐT LÕI:
Đọc kỹ đoạn hội thoại tin nhắn Messenger/Facebook HOẶC ảnh chụp màn hình Messenger / Thẻ ngoại kiều (Zairyu Card) / Ảnh bài đăng tuyển dụng ứng viên gửi kèm và trả về JSON chuẩn xác để tự động điền vào form "+ 新規相談者登録".

QUY TẮC BÓC TÁCH BẮT BUỘC (TUÂN THỦ TUYỆT ĐỐI):
1. Chỉ trích xuất thông tin thực sự có trong ảnh / hội thoại hoặc suy luận chắc chắn từ ngữ cảnh. KHÔNG tự bịa thông tin.
2. PHÂN BIỆT ỨNG VIÊN VS NGƯỜI TƯ VẤN:
   - **ỨNG VIÊN** là người hỏi tìm việc, hỏi lương/ca làm/điều kiện đơn tuyển dụng, gửi ảnh bài đăng, chia sẻ visa, hạn thẻ, tiếng Nhật, kinh nghiệm.
   - **NGƯỜI TƯ VẤN / NHÂN VIÊN** là người chào hỏi tư vấn, giới thiệu công ty, hẹn lịch Mendan. Tuyệt đối không lấy tên/SĐT của nhân viên tư vấn gán cho ứng viên.
3. QUY TẮC CHI TIẾT TỪNG TRƯỜNG:
   - **fullName** (氏名 / Họ tên khai sinh): CHỈ điền nếu trong ảnh có giấy tờ tùy thân (Thẻ ngoại kiều / Zairyu Card, Sơ yếu lý lịch / CV) ghi rõ họ tên khai sinh hoặc ứng viên tự khai rõ họ tên thật. **Nếu chỉ là ảnh chụp đoạn chat Messenger, BẮT BUỘC để trống \`""\`**.
   - **fbName** (FBアカウント名 / Tên nick FB hoặc LINE): Lấy chính xác tên hiển thị ở thanh tiêu đề trên cùng của khung chat Messenger/LINE (ví dụ: \`"Dương Như"\`, \`"Nguyễn Lyly"\`). **TUYỆT ĐỐI KHÔNG** copy đường dẫn URL (\`https://...\`, \`facebook.com/...\`) vào trường \`fbName\` (URL phải đưa vào \`fbProfileUrl\`).
   - **nationality** (国籍): Tự nhận diện quốc tịch theo tên người nhắn, ngôn ngữ chat hoặc thông tin trên thẻ ngoại kiều và chuẩn hóa về tiếng Nhật SẠCH, KHÔNG kèm chú thích ngoặc đơn: \`ベトナム\`, \`インドネシア\`, \`ミャンマー\`, \`ネパール\`, \`フィリピン\`, \`カンボジア\`, \`中国\`...
   - **japaneseLevel** (日本語レベル): Chỉ điền trình độ thực tế của ứng viên (\`N1\`, \`N2\`, \`N3\`, \`N4\`, \`N5\`, \`JFT\`, \`初級\`, \`日常会話\`) nếu ứng viên tự khai (vd: "em có N4"). **QUAN TRỌNG**: Nếu \`N3\` / \`N4\` chỉ xuất hiện trong ảnh bài đăng tuyển dụng / tin tuyển dụng mà ứng viên đang hỏi thăm (vd: yêu cầu đơn là N3/N4) chứ chưa rõ trình độ của ứng viên, BẮT BUỘC giữ nguyên là \`"要確認"\`. Nếu nói "có N4 và sắp thi N3" -> chỉ điền \`"N4"\`, ghi chú "Sắp tới thi N3" vào \`japanExperience\`.
   - **desiredField** (希望分野): Tự nhận diện ngành nghề chuẩn tiếng Nhật từ nội dung chat hoặc ảnh bài đăng tuyển dụng ứng viên gửi kèm (ví dụ: thấy "nhà hàng / quán ăn / bếp / phục vụ bàn" -> \`外食業\`; các ngành chuẩn khác: \`介護\`, \`飲食料品製造業\`, \`ビルクリーニング\`, \`宿泊\`, \`自動車運送業\`, \`建設\`, \`自動車整備\`, \`農業\`, \`素形材・産業機械・電気電子情報関連製造業\`, \`コンビニ・小売（アルバイト）\`).
   - **desiredArea** (希望勤務エリア): Tự nhận diện khu vực làm việc chuẩn tiếng Nhật từ lời nhắn hoặc bài đăng ứng viên hỏi (ví dụ: thấy \`都内\` -> điền \`東京都\`; thấy \`関東\` -> điền \`関東エリア\`; thấy \`大阪\` -> điền \`大阪府\`; nhiều tỉnh nối bằng \`・\`).
   - **visaType** & **visaExpiry** (在留資格 / 在留期限): Trích xuất loại visa (\`特定技能1号\`, \`留学\`, \`技能実習\`, \`家族滞在\`, \`技術・人文知識・国際業務\`...) và hạn thẻ \`YYYY-MM-DD\` nếu có (chú ý giải mã "dhs" -> \`留学（Du học sinh）\`).
   - **certifications** (合格済み試験・資格): Trích xuất bằng tay nghề Tokutei, JFT, bằng lái xe, sơ cấp cứu Kaigo... nếu có; không có thì \`""\`.
   - **goal** (Nhu cầu & Mục tiêu chuyển đổi): Tóm tắt ngắn gọn bằng tiếng Việt mục đích của ứng viên (bao gồm cả hình thức làm việc như Baito / Làm thêm hay Tokutei, mức lương/điều kiện của đơn mà ứng viên đang hỏi).
   - **rawLog** (Nhật ký & Tình trạng trao đổi nguyên văn): **BẮT BUỘC** trích xuất TOÀN BỘ chữ có trong ảnh (bao gồm CẢ nội dung văn bản trong ảnh bài đăng tuyển dụng mà ứng viên gửi kèm lẫn toàn bộ tin nhắn hội thoại kèm mốc giờ), giữ nguyên từng dòng không cắt bớt.
4. CÁC TRƯỜNG BỔ SUNG (điền nếu có thực tế trong ảnh/chat; nếu không có thì trả về \`""\`, TUYỆT ĐỐI KHÔNG điền chuỗi giữ chỗ kiểu \`"(Chưa có...)"\`, \`"Chưa rõ"\`, \`"未定"\`, và KHÔNG tự suy đoán \`mendanDate\` / \`interviewDate\` từ câu hỏi giả định):
   - \`fbProfileUrl\`, \`channel\` (\`"personal"\` FB cá nhân mặc định, \`"company"\` Fanpage công ty, \`"school"\` Trường Nhật ngữ / DHS, \`"line"\` nếu nguồn là LINE, \`"other"\` nếu nguồn khác / giới thiệu trực tiếp), \`age\`, \`birthDate\` (\`YYYY-MM-DD\`), \`phone\`, \`email\`, \`address\`, \`startDate\` (năm hiện tại 2026 -> "tháng X năm sau" là 2027), \`graduationDate\`, \`japanExperience\`, \`salaryExpectation\`, \`progressStatus\` (mặc định \`"新規相談"\`), \`proposedCompany\`, \`mendanDate\`, \`interviewDate\`.`;

async function startServer() {
  const app = express();
  app.use(
    express.json({
      limit: '35mb',
      verify: (req, _res, buf) => {
        (req as RawBodyRequest).rawBody = buf;
      },
    })
  );

  const pendingWebhookCandidates = new Map<
    string,
    {
      senderId: string;
      rawTexts: string[];
      extracted: MessengerExtractedProfile;
      candidate: Candidate;
    }
  >();

  function getJstDateYmd(): string {
    return new Date(Date.now() + 9 * 60 * 60 * 1000)
      .toISOString()
      .slice(0, 10);
  }

  function upsertPendingWebhookCandidate(
    senderId: string,
    rawText: string,
    extracted: MessengerExtractedProfile
  ): Candidate {
    const todayYmd = getJstDateYmd();
    const existing = pendingWebhookCandidates.get(senderId);
    const memoLine = `[Messenger ${todayYmd}] ${rawText}`;

    if (existing) {
      const nextCand: Candidate = {
        ...existing.candidate,
        name: existing.candidate.name.startsWith('FB相談者') && extracted.name
          ? extracted.name
          : existing.candidate.name,
        fbName:
          extracted.name && existing.candidate.fbName?.startsWith('Messenger ')
            ? extracted.name
            : existing.candidate.fbName,
        jlpt: existing.candidate.jlpt || extracted.jlpt || '',
        phone: existing.candidate.phone || extracted.phone || '',
        desiredIndustry:
          existing.candidate.desiredIndustry || extracted.desiredIndustry || '',
        desiredArea:
          existing.candidate.desiredArea || extracted.desiredArea || '',
        visaExpiry: existing.candidate.visaExpiry || extracted.visaExpiry || '',
        lastContactDate: todayYmd,
        memo: existing.candidate.memo
          ? `${existing.candidate.memo}\n${memoLine}`
          : memoLine,
      };
      existing.rawTexts.push(rawText);
      existing.extracted = {
        name: existing.extracted.name || extracted.name,
        jlpt: existing.extracted.jlpt || extracted.jlpt,
        phone: existing.extracted.phone || extracted.phone,
        desiredIndustry:
          existing.extracted.desiredIndustry || extracted.desiredIndustry,
        desiredArea: existing.extracted.desiredArea || extracted.desiredArea,
        visaExpiry: existing.extracted.visaExpiry || extracted.visaExpiry,
      };
      existing.candidate = nextCand;
      return nextCand;
    }

    const created: Candidate = {
      id: `fb_${senderId}`,
      fbSenderId: senderId,
      name: extracted.name || `FB相談者 (${senderId.slice(-6)})`,
      fbName: extracted.name || `Messenger ${senderId.slice(-6)}`,
      channel: 'company',
      progressStatus: '新規相談',
      jlpt: extracted.jlpt || '',
      phone: extracted.phone || '',
      desiredIndustry: extracted.desiredIndustry || '',
      desiredArea: extracted.desiredArea || '',
      visaExpiry: extracted.visaExpiry || '',
      lastContactDate: todayYmd,
      memo: memoLine,
    };
    pendingWebhookCandidates.set(senderId, {
      senderId,
      rawTexts: [rawText],
      extracted: { ...extracted },
      candidate: created,
    });
    return created;
  }

  // =========================================================================
  // AUTH SESSION & FACEBOOK MESSENGER WEBHOOK
  // =========================================================================
  app.get('/api/auth/session', requireAuth, (req: AuthRequest, res) => {
    res.json({
      authorized: true,
      email: req.dbUser!.email,
      displayName: req.dbUser!.displayName,
      storageMode: req.dbAvailable === false ? 'local_fallback' : 'cloudsql',
    });
  });

  app.get('/api/webhooks/messenger/status', optionalAuth, (req: AuthRequest, res) => {
    const hasVerifyToken = Boolean((process.env.FB_VERIFY_TOKEN || '').trim());
    const hasAppSecret = Boolean((process.env.FB_APP_SECRET || '').trim());
    const rawAppUrl = (process.env.APP_URL || '').trim().replace(/\/+$/, '');
    const webhookUrl =
      rawAppUrl && /^https?:\/\//i.test(rawAppUrl) && !rawAppUrl.includes('MY_APP_URL')
        ? `${rawAppUrl}/api/webhooks/messenger`
        : 'https://fb-crm.ai.studio/api/webhooks/messenger';
    const sharedWebhookUrl = webhookUrl.includes('://ais-dev-')
      ? webhookUrl.replace('://ais-dev-', '://ais-pre-')
      : '';
    res.json({
      configured: true,
      hasVerifyToken,
      hasAppSecret,
      verifyTokenEnforced: hasVerifyToken,
      signatureEnforced: hasAppSecret,
      storageMode: req.dbAvailable === false ? 'local_fallback' : 'cloudsql',
      webhookUrl,
      sharedWebhookUrl,
    });
  });

  app.get('/api/webhooks/messenger', (req, res) => {
    const mode = req.query['hub.mode'];
    const token = req.query['hub.verify_token'];
    const challenge = req.query['hub.challenge'];
    const verifyToken = (process.env.FB_VERIFY_TOKEN || '').trim();

    if (mode !== 'subscribe') {
      res.sendStatus(403);
      return;
    }

    if (verifyToken.length > 0) {
      if (
        typeof token === 'string' &&
        timingSafeStringEqual(token, verifyToken)
      ) {
        res.status(200).send(typeof challenge === 'string' ? challenge : '');
        return;
      }
      res.sendStatus(403);
      return;
    }

    // Allow easy webhook verification when FB_VERIFY_TOKEN is not yet configured
    res.status(200).send(typeof challenge === 'string' ? challenge : '');
  });

  app.post('/api/webhooks/messenger', async (req: RawBodyRequest, res) => {
    const appSecret = (process.env.FB_APP_SECRET || '').trim();
    const signatureHeader = req.headers['x-hub-signature-256'];

    if (
      appSecret.length > 0 &&
      !verifyMetaSignature(req.rawBody, signatureHeader, appSecret)
    ) {
      res.status(403).json({ error: 'Forbidden: Invalid X-Hub-Signature-256' });
      return;
    }

    const body = req.body as {
      object?: string;
      entry?: Array<{
        id?: string;
        time?: number;
        messaging?: Array<{
          sender?: { id?: string };
          recipient?: { id?: string };
          timestamp?: number;
          message?: { mid?: string; text?: string; is_echo?: boolean };
        }>;
      }>;
    };

    if (body?.object !== 'page') {
      res.sendStatus(404);
      return;
    }

    try {
      const primaryUser = await getPrimaryCrmUser();
      const entries = Array.isArray(body.entry) ? body.entry : [];
      let processedCount = 0;

      for (const entry of entries) {
        const messagingEvents = Array.isArray(entry?.messaging)
          ? entry.messaging
          : [];
        for (const event of messagingEvents) {
          const senderId = event?.sender?.id;
          const isEcho = Boolean(event?.message?.is_echo);
          const text = event?.message?.text;

          if (
            !senderId ||
            isEcho ||
            typeof text !== 'string' ||
            !text.trim()
          ) {
            continue;
          }

          let extracted: MessengerExtractedProfile = {};
          try {
            extracted = await extractMessengerCandidateWithGemini(text.trim());
          } catch (aiErr) {
            const errMsg =
              aiErr instanceof Error ? aiErr.message : 'Gemini parse failed';
            console.warn(
              '[/api/webhooks/messenger] Gemini extraction skipped:',
              errMsg
            );
          }

          let savedToDb = false;
          if (primaryUser && primaryUser.id > 0) {
            try {
              await upsertMessengerCandidate(primaryUser.id, {
                senderId: String(senderId),
                rawText: text.trim(),
                extracted,
              });
              savedToDb = true;
            } catch (dbErr) {
              console.warn(
                '[/api/webhooks/messenger] DB store fallback to memory:',
                dbErr instanceof Error ? dbErr.message : dbErr
              );
            }
          }

          if (!savedToDb) {
            upsertPendingWebhookCandidate(
              String(senderId),
              text.trim(),
              extracted
            );
          }
          processedCount += 1;
        }
      }

      res.status(200).json({ status: 'EVENT_RECEIVED', processedCount });
    } catch (err) {
      const msg =
        err instanceof Error ? err.message : 'Webhook processing failed';
      console.error('[/api/webhooks/messenger] Error:', msg);
      res.status(500).json({ error: msg });
    }
  });

  app.post('/api/webhooks/messenger/test', optionalAuth, async (req: AuthRequest, res) => {
    const rawText =
      typeof req.body?.text === 'string' && req.body.text.trim()
        ? req.body.text.trim().slice(0, 4000)
        : 'Em tên Phạm Minh Tuấn, có bằng N3, SĐT 080-1234-5678, muốn tìm việc 介護 tại 東京都, hạn visa 2027-08-15 ạ.';
    const senderId = `test_${Date.now().toString(36)}`;

    let extracted: MessengerExtractedProfile = {};
    try {
      extracted = await extractMessengerCandidateWithGemini(rawText);
    } catch (aiErr) {
      console.warn(
        '[/api/webhooks/messenger/test] Gemini extraction fallback:',
        aiErr instanceof Error ? aiErr.message : aiErr
      );
    }

    if (req.dbAvailable !== false && req.dbUser && req.dbUser.id > 0) {
      try {
        const result = await upsertMessengerCandidate(req.dbUser.id, {
          senderId,
          rawText,
          extracted,
        });
        res.json({
          ok: true,
          persisted: true,
          candidate: result.candidate,
        });
        return;
      } catch (dbErr) {
        console.warn(
          '[/api/webhooks/messenger/test] DB write skipped:',
          dbErr instanceof Error ? dbErr.message : dbErr
        );
      }
    }

    const todayYmd = getJstDateYmd();
    res.json({
      ok: true,
      persisted: false,
      candidate: {
        id: `fb_${senderId}`,
        fbSenderId: senderId,
        name: extracted.name || 'Phạm Minh Tuấn (Messenger Test)',
        fbName: extracted.name || 'Messenger Test',
        channel: 'company',
        progressStatus: '新規相談',
        jlpt: extracted.jlpt || 'N3',
        phone: extracted.phone || '',
        desiredIndustry: extracted.desiredIndustry || '介護',
        desiredArea: extracted.desiredArea || '東京都',
        visaExpiry: extracted.visaExpiry || '',
        lastContactDate: todayYmd,
        memo: `[Messenger Test] ${rawText}`,
      },
    });
  });

  app.post('/api/gemini/extract-candidate', optionalAuth, async (req, res) => {
    try {
      const { text, parts } = req.body ?? {};
      const contentParts: Array<
        { text: string } | { inlineData: { mimeType: string; data: string } }
      > = [{ text: CANDIDATE_EXTRACTION_PROMPT }];

      if (typeof text === 'string' && text.trim()) {
        contentParts.push({ text: `Dữ liệu gốc cần bóc tách (giữ nguyên 100% chi tiết):\n\n${text}` });
      }
      if (Array.isArray(parts)) {
        for (const p of parts as ClientPart[]) {
          if (p && typeof p.text === 'string' && p.text.trim()) {
            contentParts.push({ text: p.text });
          } else if (
            p?.inlineData &&
            typeof p.inlineData.mimeType === 'string' &&
            ALLOWED_MIME_TYPES.has(p.inlineData.mimeType) &&
            typeof p.inlineData.data === 'string' &&
            p.inlineData.data.length > 0
          ) {
            contentParts.push({
              inlineData: {
                mimeType: p.inlineData.mimeType,
                data: p.inlineData.data,
              },
            });
          }
        }
      }

      if (contentParts.length <= 1) {
        res.status(400).json({ error: 'Vui lòng dán văn bản hoặc ảnh chụp màn hình để trích xuất.' });
        return;
      }

      const ai = getGeminiClient();
      const response = await ai.models.generateContent({
        model: MODEL_NAME,
        contents: { parts: contentParts },
        config: {
          responseMimeType: 'application/json',
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              fullName: { type: Type.STRING },
              name: { type: Type.STRING },
              fbName: { type: Type.STRING },
              fbProfileUrl: { type: Type.STRING },
              channel: { type: Type.STRING },
              nationality: { type: Type.STRING },
              gender: { type: Type.STRING },
              birthDate: { type: Type.STRING },
              age: { type: Type.STRING },
              japaneseLevel: { type: Type.STRING },
              jlpt: { type: Type.STRING },
              visaType: { type: Type.STRING },
              visaStatus: { type: Type.STRING },
              visaExpiry: { type: Type.STRING },
              desiredField: { type: Type.STRING },
              desiredIndustry: { type: Type.STRING },
              desiredArea: { type: Type.STRING },
              goal: { type: Type.STRING },
              conversionGoal: { type: Type.STRING },
              certifications: { type: Type.STRING },
              passedExam: { type: Type.STRING },
              japanExperience: { type: Type.STRING },
              salaryExpectation: { type: Type.STRING },
              rawLog: { type: Type.STRING },
              negotiationLog: { type: Type.STRING },
              startDate: { type: Type.STRING },
              graduationDate: { type: Type.STRING },
              phone: { type: Type.STRING },
              email: { type: Type.STRING },
              address: { type: Type.STRING },
              progressStatus: { type: Type.STRING },
              proposedCompany: { type: Type.STRING },
              mendanDate: { type: Type.STRING },
              interviewDate: { type: Type.STRING },
            },
          },
        },
      });

      const raw = safeParseGeminiJson<Record<string, unknown>>(response.text, {});
      const pickStr = (...vals: unknown[]) => {
        for (const v of vals) {
          if (typeof v === 'string' && v.trim()) return v.trim();
        }
        return '';
      };
      const cleanOptionalField = (val: unknown): string => {
        const s = typeof val === 'string' ? val.trim() : '';
        if (
          !s ||
          /^(?:\(?chưa\s*(?:có|rõ|biết|xác\s*định)[^)]*\)?|\(?không\s*(?:có|rõ)[^)]*\)?|未定|なし|不明|該当なし|n\/a|-+)$/i.test(
            s
          )
        ) {
          return '';
        }
        return s;
      };
      const cleanFbName = (val: unknown): string => {
        const s = cleanOptionalField(val);
        if (!s || /^https?:\/\//i.test(s) || /(?:facebook\.com|fb\.com|m\.me)\//i.test(s)) {
          return '';
        }
        return s;
      };
      const cleanNat = (val: unknown): string => {
        const s = cleanOptionalField(val).replace(/\s*[（(][^）)]*[）)]\s*/g, '').trim();
        if (!s) return '';
        if (/viet|việt|ベトナム/i.test(s)) return 'ベトナム';
        if (/indo|インドネシア/i.test(s)) return 'インドネシア';
        if (/myanmar|miến|ミャンマー/i.test(s)) return 'ミャンマー';
        if (/nepal|ネパール/i.test(s)) return 'ネパール';
        if (/philip|フィリピン/i.test(s)) return 'フィリピン';
        if (/cambo|campu|カンボジア/i.test(s)) return 'カンボジア';
        if (/china|trung\s*quốc|中国/i.test(s)) return '中国';
        return s;
      };
      const normalizeIndustryStr = (val: unknown): string => {
        const s = cleanOptionalField(val);
        if (!s) return '';
        if (/nhà\s*hàng|quán\s*ăn|bếp|phục\s*vụ\s*bàn|ẩm\s*thực/i.test(s) && !/外食/.test(s)) {
          return '外食業';
        }
        return s;
      };
      const rawChan = pickStr(raw.channel).toLowerCase();
      const normalizedChannel =
        rawChan === 'company' ||
        rawChan === 'school' ||
        rawChan === 'line' ||
        rawChan === 'other' ||
        rawChan === 'personal'
          ? rawChan
          : /line/i.test(pickStr(raw.rawLog, raw.negotiationLog))
          ? 'line'
          : 'personal';

      const resolvedArea = pickStr(raw.desiredArea)
        .replace(/都内/g, '東京都')
        .replace(/^関東$/g, '関東エリア');

      const candidate: Record<string, unknown> = {
        ...raw,
        name: cleanOptionalField(pickStr(raw.fullName, raw.name)),
        fullName: cleanOptionalField(pickStr(raw.fullName, raw.name)),
        fbName: cleanFbName(raw.fbName),
        fbProfileUrl: cleanOptionalField(pickStr(raw.fbProfileUrl, raw.fbUrl)),
        fbUrl: cleanOptionalField(pickStr(raw.fbProfileUrl, raw.fbUrl)),
        channel: normalizedChannel,
        nationality: cleanNat(raw.nationality),
        jlpt: pickStr(raw.japaneseLevel, raw.jlpt) || '要確認',
        japaneseLevel: pickStr(raw.japaneseLevel, raw.jlpt) || '要確認',
        desiredIndustry: normalizeIndustryStr(pickStr(raw.desiredField, raw.desiredIndustry)),
        desiredField: normalizeIndustryStr(pickStr(raw.desiredField, raw.desiredIndustry)),
        desiredArea: resolvedArea,
        visaStatus: cleanOptionalField(pickStr(raw.visaType, raw.visaStatus)),
        visaType: cleanOptionalField(pickStr(raw.visaType, raw.visaStatus)),
        visaExpiry: cleanOptionalField(raw.visaExpiry),
        passedExam: cleanOptionalField(pickStr(raw.certifications, raw.passedExam)),
        certifications: cleanOptionalField(pickStr(raw.certifications, raw.passedExam)),
        conversionGoal: cleanOptionalField(pickStr(raw.goal, raw.conversionGoal)),
        goal: cleanOptionalField(pickStr(raw.goal, raw.conversionGoal)),
        phone: cleanOptionalField(raw.phone),
        email: cleanOptionalField(raw.email),
        address: cleanOptionalField(raw.address),
        mendanDate: cleanOptionalField(raw.mendanDate),
        interviewDate: cleanOptionalField(raw.interviewDate),
        proposedCompany: cleanOptionalField(raw.proposedCompany),
        negotiationLog: pickStr(raw.rawLog, raw.negotiationLog),
        rawLog: pickStr(raw.rawLog, raw.negotiationLog),
      };
      res.json({ candidate });
    } catch (err) {
      const code = (err as { code?: string })?.code;
      const message = err instanceof Error ? err.message : 'Candidate extraction failed';
      console.error('[/api/gemini/extract-candidate] Error:', message);
      res.status(code === 'MISSING_GEMINI_KEY' ? 503 : 500).json({ error: message, code });
    }
  });

  app.post('/api/gemini/extract-kyujin', optionalAuth, async (req, res) => {
    try {
      const { mimeType, base64Data, rawText, fileName, images } = req.body ?? {};
      const parts: Array<
        { text: string } | { inlineData: { mimeType: string; data: string } }
      > = [{ text: KYUJIN_EXTRACTION_PROMPT }];

      if (typeof rawText === 'string' && rawText.trim()) {
        parts.push({
          text: `Tên nguồn/file: ${typeof fileName === 'string' ? fileName : 'Văn bản thô'}\n\nNội dung tuyển dụng:\n${rawText}`,
        });
      } else if (Array.isArray(images) && images.length > 0) {
        const validImages = images.filter(
          (img): img is { mimeType: string; base64Data: string } =>
            Boolean(
              img &&
                typeof img === 'object' &&
                typeof (img as { mimeType?: unknown }).mimeType === 'string' &&
                ALLOWED_MIME_TYPES.has((img as { mimeType: string }).mimeType) &&
                typeof (img as { base64Data?: unknown }).base64Data === 'string' &&
                (img as { base64Data: string }).base64Data.length > 0
            )
        );
        if (validImages.length === 0) {
          res.status(400).json({ error: 'Không có trang ảnh hợp lệ để nhận diện OCR.' });
          return;
        }
        parts.push({
          text: `Tên nguồn/file: ${typeof fileName === 'string' ? fileName : 'PDF_Scan'} (${validImages.length} trang ảnh scan). Hãy đọc toàn bộ chữ tiếng Nhật, bảng lương (給与/月給/時給/万), địa điểm và điều kiện tuyển dụng trên tất cả các trang ảnh này.`,
        });
        validImages.slice(0, 12).forEach((img) => {
          parts.push({
            inlineData: { mimeType: img.mimeType, data: img.base64Data },
          });
        });
      } else {
        if (typeof mimeType !== 'string' || !ALLOWED_MIME_TYPES.has(mimeType)) {
          res.status(400).json({ error: 'Định dạng file không được hỗ trợ cho AI OCR.' });
          return;
        }
        if (typeof base64Data !== 'string' || base64Data.length === 0) {
          res.status(400).json({ error: 'Nội dung file bị rỗng.' });
          return;
        }
        parts.push({ inlineData: { mimeType, data: base64Data } });
      }

      const ai = getGeminiClient();
      const response = await ai.models.generateContent({
        model: MODEL_NAME,
        contents: { parts },
        config: {
          responseMimeType: 'application/json',
          responseSchema: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                companyName: { type: Type.STRING },
                jobCode: { type: Type.STRING },
                sector: { type: Type.STRING },
                industry: { type: Type.STRING },
                employmentType: { type: Type.STRING },
                facilityType: { type: Type.STRING },
                area: { type: Type.STRING },
                shiftTime: { type: Type.STRING },
                shiftType: { type: Type.STRING },
                holidays: { type: Type.STRING },
                overtimeHours: { type: Type.STRING },
                salary: { type: Type.STRING },
                nightShiftAllowance: { type: Type.STRING },
                estimatedNetSalary: { type: Type.STRING },
                annualSalary: { type: Type.STRING },
                hourlyWage: { type: Type.STRING },
                commuteAllowance: { type: Type.STRING },
                shiftsPerWeek: { type: Type.STRING },
                recruitCount: { type: Type.STRING },
                interviewProcess: { type: Type.STRING },
                specialBenefits: { type: Type.STRING },
                minJlpt: { type: Type.STRING },
                housingSupport: { type: Type.STRING },
                bonus: { type: Type.STRING },
                bonusDetail: { type: Type.STRING },
                experienceRequirement: { type: Type.STRING },
                experienceDetail: { type: Type.STRING },
                jobDescription: { type: Type.STRING },
                excludedNationalities: {
                  type: Type.ARRAY,
                  items: { type: Type.STRING },
                },
                excludedGenders: {
                  type: Type.ARRAY,
                  items: { type: Type.STRING },
                },
                requiresInJapan: { type: Type.BOOLEAN },
                requiresTokutei: { type: Type.BOOLEAN },
                requiredConditions: {
                  type: Type.ARRAY,
                  items: { type: Type.STRING },
                },
                ngConditions: {
                  type: Type.ARRAY,
                  items: { type: Type.STRING },
                },
              },
              required: [
                'companyName',
                'jobCode',
                'sector',
                'industry',
                'employmentType',
                'area',
                'salary',
                'minJlpt',
                'housingSupport',
              ],
            },
          },
        },
      });

      const parsed = safeParseGeminiJson<unknown>(response.text, []);
      const rawItems = (Array.isArray(parsed) ? parsed : [parsed]).filter(
        (x): x is Record<string, unknown> => Boolean(x && typeof x === 'object')
      );
      const rawTextFallbackSalary =
        typeof rawText === 'string' && rawText.trim() && rawItems.length <= 1
          ? extractKyujinSalaryInfo(rawText)
          : null;
      const cleanKyujinField = (v: unknown): string =>
        sanitizeExcelFloatText(stripKyujinFormBoilerplate(String(v ?? '').trim()));

      const normalizedItems = rawItems.map((item) => {
        const aiSal = cleanKyujinField(item.salary);
        const aiHourly = cleanKyujinField(item.hourlyWage);
        const combinedHint = [
          aiSal && !/要確認/.test(aiSal) ? aiSal : '',
          aiHourly && !/要確認/.test(aiHourly) ? `時給 ${aiHourly}` : '',
          cleanKyujinField(item.jobDescription),
        ]
          .filter(Boolean)
          .join('\n');

        let salaryParsed = extractKyujinSalaryInfo(combinedHint, {
          assumeBareNumberIsSalary: Boolean(aiSal && !/要確認/.test(aiSal)),
        });

        if (salaryParsed.salaryType === 'unknown' && rawTextFallbackSalary?.salaryType !== 'unknown' && rawTextFallbackSalary) {
          salaryParsed = rawTextFallbackSalary;
        }

        const finalSalary =
          salaryParsed.salaryType !== 'unknown'
            ? salaryParsed.salary
            : aiSal || (aiHourly ? `時給 ${aiHourly}` : '給与要確認');
        const finalHourly =
          salaryParsed.salaryType === 'hourly'
            ? salaryParsed.hourlyWage
            : aiHourly;

        const cleanNet =
          cleanKyujinField(item.estimatedNetSalary) ||
          (typeof rawText === 'string' && rawItems.length <= 1
            ? extractEstimatedNetSalaryText(rawText)
            : '');
        const cleanAnnual =
          cleanKyujinField(item.annualSalary) ||
          (typeof rawText === 'string' && rawItems.length <= 1
            ? extractAnnualSalaryText(rawText)
            : '');
        const rawJlpt = cleanKyujinField(item.minJlpt);
        const normJlptMatch = rawJlpt.match(/(N[1-5])\s*(?:以上|程度|相当)?/i);
        const finalJlpt = normJlptMatch ? normJlptMatch[1].toUpperCase() : rawJlpt;

        const bonusRaw = cleanKyujinField(item.bonus);
        const bonusDetailRaw = cleanKyujinField(item.bonusDetail);
        const bonusCombined = `${bonusRaw} ${bonusDetailRaw}`;
        const finalBonus = /(?:賞与なし|賞与\s*[:：]?\s*無|退職金・賞与無|ボーナス無|年0回)/.test(
          bonusCombined
        )
          ? '賞与なし'
          : bonusRaw;

        return {
          ...item,
          companyName: cleanKyujinField(item.companyName),
          jobCode: cleanKyujinField(item.jobCode),
          area: cleanKyujinField(item.area),
          shiftTime: cleanKyujinField(item.shiftTime),
          shiftType: cleanKyujinField(item.shiftType),
          holidays: cleanKyujinField(item.holidays),
          overtimeHours: cleanKyujinField(item.overtimeHours),
          salary: sanitizeExcelFloatText(finalSalary),
          hourlyWage: sanitizeExcelFloatText(finalHourly),
          estimatedNetSalary: cleanNet,
          annualSalary: cleanAnnual,
          nightShiftAllowance: cleanKyujinField(item.nightShiftAllowance),
          commuteAllowance: cleanKyujinField(item.commuteAllowance),
          shiftsPerWeek: cleanKyujinField(item.shiftsPerWeek),
          recruitCount: cleanKyujinField(item.recruitCount),
          interviewProcess: cleanKyujinField(item.interviewProcess),
          specialBenefits: cleanKyujinField(item.specialBenefits),
          minJlpt: finalJlpt || rawJlpt,
          bonus: finalBonus || item.bonus,
          bonusDetail: bonusDetailRaw,
          experienceDetail: cleanKyujinField(item.experienceDetail),
          jobDescription: cleanKyujinField(item.jobDescription),
        };
      });

      res.json({ items: normalizedItems });
    } catch (err) {
      const code = (err as { code?: string })?.code;
      const message = err instanceof Error ? err.message : 'Kyujin extraction failed';
      console.error('[/api/gemini/extract-kyujin] Error:', message);
      res.status(code === 'MISSING_GEMINI_KEY' ? 503 : 500).json({ error: message, code });
    }
  });

  app.post('/api/gemini/extract-rirekisho', optionalAuth, async (req, res) => {
    try {
      const rawParts = req.body?.parts;
      if (!Array.isArray(rawParts) || rawParts.length === 0) {
        res.status(400).json({ error: 'Thiếu dữ liệu đầu vào để tạo Sơ yếu lý lịch.' });
        return;
      }

      const sanitizedParts: Array<
        { text: string } | { inlineData: { mimeType: string; data: string } }
      > = [{ text: RIREKISHO_EXTRACTION_PROMPT }];

      for (const p of rawParts as ClientPart[]) {
        if (p && typeof p.text === 'string' && p.text.trim()) {
          sanitizedParts.push({ text: p.text });
        } else if (
          p?.inlineData &&
          typeof p.inlineData.mimeType === 'string' &&
          ALLOWED_MIME_TYPES.has(p.inlineData.mimeType) &&
          typeof p.inlineData.data === 'string' &&
          p.inlineData.data.length > 0
        ) {
          sanitizedParts.push({
            inlineData: {
              mimeType: p.inlineData.mimeType,
              data: p.inlineData.data,
            },
          });
        }
      }

      if (sanitizedParts.length <= 1) {
        res.status(400).json({ error: 'Không có văn bản hoặc ảnh/PDF hợp lệ để đọc.' });
        return;
      }

      const ai = getGeminiClient();
      const response = await ai.models.generateContent({
        model: MODEL_NAME,
        contents: { parts: sanitizedParts },
        config: {
          responseMimeType: 'application/json',
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              furigana: { type: Type.STRING },
              fullName: { type: Type.STRING },
              birthDate: { type: Type.STRING },
              gender: { type: Type.STRING },
              postalCode: { type: Type.STRING },
              address: { type: Type.STRING },
              phone: { type: Type.STRING },
              email: { type: Type.STRING },
              nationality: { type: Type.STRING },
              visaExpiry: { type: Type.STRING },
              visaType: { type: Type.STRING },
              educationList: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    year: { type: Type.STRING },
                    month: { type: Type.STRING },
                    content: { type: Type.STRING },
                  },
                },
              },
              workHistoryList: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    year: { type: Type.STRING },
                    month: { type: Type.STRING },
                    content: { type: Type.STRING },
                  },
                },
              },
              qualificationsList: {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    year: { type: Type.STRING },
                    month: { type: Type.STRING },
                    content: { type: Type.STRING },
                  },
                },
              },
              nativeLanguage: { type: Type.STRING },
              spokenLanguages: { type: Type.STRING },
              speakingLevel: { type: Type.STRING },
              listeningLevel: { type: Type.STRING },
              motivation: { type: Type.STRING },
              selfPr: { type: Type.STRING },
              allergies: { type: Type.STRING },
              dependentsCount: { type: Type.STRING },
              hasSpouse: { type: Type.STRING },
              spouseSupportObligation: { type: Type.STRING },
              wishes: { type: Type.STRING },
            },
          },
        },
      });

      const extracted = safeParseGeminiJson<Record<string, unknown>>(response.text, {});
      res.json({ extracted });
    } catch (err) {
      const code = (err as { code?: string })?.code;
      const message = err instanceof Error ? err.message : 'Rirekisho extraction failed';
      console.error('[/api/gemini/extract-rirekisho] Error:', message);
      res.status(code === 'MISSING_GEMINI_KEY' ? 503 : 500).json({ error: message, code });
    }
  });

  // =========================================================================
  // MULTI-TURN GEMINI CHATBOT ENDPOINT
  // =========================================================================
  const CHATBOT_ROLE_CONFIG: Record<
    string,
    { model: 'gemini-4-hydrogen' | 'gemini-3.1-flash-lite'; instruction: string }
  > = {
    crm_strategist: {
      model: 'gemini-4-hydrogen',
      instruction:
        'Bạn là Chuyên viên Tư vấn Tuyển dụng & Chiến lược Tokutei Ginou (特定技能) trong hệ thống CRM. Hãy hỗ trợ điều phối viên phân tích tiến độ ứng viên, ưu tiên follow-up, ghép đơn hàng (Kaigo, Ngoại thực, Khách sạn, Chế biến thực phẩm, Vận tải, Baito) và giải đáp quy định visa/JLPT/thủ tục nghỉ việc rõ ràng, súc tích, thực tế. Trả lời bằng ngôn ngữ mà người dùng đang sử dụng (tiếng Việt hoặc tiếng Nhật).',
    },
    reply_coach: {
      model: 'gemini-3.1-flash-lite',
      instruction:
        'Bạn là Trợ lý Soạn tin nhắn nhanh đa ngôn ngữ (Việt / Nhật / Anh) chuyên cho tư vấn viên Tokutei Ginou nhắn tin Messenger/Zalo với ứng viên hoặc gửi email/chat cho doanh nghiệp tiếp nhận tại Nhật. Phản hồi cực kỳ ngắn gọn, lịch sự, tự nhiên, sẵn sàng copy-paste ngay.',
    },
    screening_advisor: {
      model: 'gemini-4-hydrogen',
      instruction:
        'Bạn là Cố vấn Đánh giá Hồ sơ & Lịch phỏng vấn (Mendan/面接) cho lao động kỹ năng đặc định và du học sinh tại Nhật. Hãy giúp kiểm tra các thông tin còn thiếu (JLPT, hạn thẻ ngoại kiều, chứng chỉ ngành, khu vực ưu tiên), gợi ý câu hỏi phỏng vấn sơ vấn và lưu ý khi gửi Rirekisho 2 cột cho xí nghiệp.',
    },
  };

  app.post('/api/gemini/chat', optionalAuth, async (req: AuthRequest, res) => {
    try {
      const {
        message,
        history,
        roleId = 'crm_strategist',
        crmSummary,
        lang: rawLang,
      } = req.body ?? {};
      const lang: 'vi' | 'jp' = rawLang === 'jp' ? 'jp' : 'vi';

      if (typeof message !== 'string' || !message.trim()) {
        res.status(400).json({ error: 'Tin nhắn không được để trống.' });
        return;
      }
      if (message.length > 8000) {
        res.status(400).json({ error: 'Nội dung tin nhắn quá dài.' });
        return;
      }

      const verifiedDbUserId = req.dbUser?.id || 0;

      const safeRoleKey =
        typeof roleId === 'string' && roleId in CHATBOT_ROLE_CONFIG
          ? roleId
          : 'crm_strategist';
      const roleCfg = CHATBOT_ROLE_CONFIG[safeRoleKey];

      const formattedHistory = Array.isArray(history)
        ? history
            .filter(
              (item): item is { role: 'user' | 'model'; text: string } =>
                Boolean(item) &&
                (item.role === 'user' || item.role === 'model') &&
                typeof item.text === 'string' &&
                item.text.trim().length > 0
            )
            .slice(-24)
            .map((item) => ({
              role: item.role,
              parts: [{ text: item.text.slice(0, 8000) }],
            }))
        : [];

      const strictPolicy =
        lang === 'jp'
          ? '【厳守ルール】1. 必ず自然で丁寧な日本語のみで回答してください（求人票・送信文サンプルとしてベトナム語を明示的に依頼された箇所のみ例外）。ベトナム語の解説や見出しを混在させないでください。2. 件数・人数・企業数・退職者数などの数値は、必ず下記の「CRM集計データ」に記載された実数のみを正確に引用し、記載されていない数値を推測・創作しないでください。'
          : '【NGUYÊN TẮC BẮT BUỘC】1. Trả lời bằng tiếng Việt tự nhiên, rõ ràng (chỉ chèn tiếng Nhật khi là thuật ngữ/tên trường hoặc đoạn tin nhắn mẫu tiếng Nhật được yêu cầu). 2. Mọi số liệu thống kê (số ứng viên, kênh, trạng thái, công ty, người nghỉ việc...) PHẢI lấy chính xác tuyệt đối từ phần [Ngữ cảnh dữ liệu CRM hiện tại] bên dưới; tuyệt đối KHÔNG tự bịa hay ước lượng số liệu không có trong ngữ cảnh.';

      const systemInstruction =
        typeof crmSummary === 'string' && crmSummary.trim()
          ? `${roleCfg.instruction}\n\n${strictPolicy}\n\n[Ngữ cảnh dữ liệu CRM hiện tại]\n${crmSummary.slice(0, 4500)}`
          : `${roleCfg.instruction}\n\n${strictPolicy}`;

      const ai = getGeminiClient();
      const chat = ai.chats.create({
        model: roleCfg.model,
        config: {
          systemInstruction,
        },
        history: formattedHistory,
      });

      const response = await chat.sendMessage({
        message: message.trim(),
      });

      const replyText = (response.text ?? '').trim();
      if (!replyText) {
        res.status(502).json({ error: 'AI không trả về nội dung phản hồi.' });
        return;
      }

      let persisted = false;
      if (req.dbAvailable !== false && verifiedDbUserId > 0) {
        try {
          await appendUserChatExchange(
            verifiedDbUserId,
            'default',
            safeRoleKey,
            roleCfg.model,
            message.trim(),
            replyText
          );
          persisted = true;
        } catch (persistErr) {
          console.warn(
            '[/api/gemini/chat] DB persistence skipped:',
            persistErr instanceof Error ? persistErr.message : persistErr
          );
        }
      }

      res.json({
        reply: replyText,
        modelUsed: roleCfg.model,
        roleId: safeRoleKey,
        persisted,
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Gemini chat failed';
      console.error('[/api/gemini/chat] Error:', msg);
      res.status(500).json({ error: msg });
    }
  });

  // =========================================================================
  // PROTECTED CLOUD SQL CRM & CHAT PERSISTENCE ROUTES
  // =========================================================================
  app.get('/api/crm/state', requireAuth, async (req: AuthRequest, res) => {
    const bufferedCandidates = Array.from(
      pendingWebhookCandidates.values()
    ).map((item) => item.candidate);

    if (req.dbAvailable === false || !req.dbUser || req.dbUser.id <= 0) {
      res.json({
        candidates: [],
        schoolStudents: [],
        kyujins: [],
        ptPartners: [],
        resignations: [],
        pendingWebhookCandidates: bufferedCandidates,
        hasExistingData: false,
        storageMode: 'local_fallback',
      });
      return;
    }

    try {
      if (pendingWebhookCandidates.size > 0) {
        for (const [senderId, pending] of pendingWebhookCandidates.entries()) {
          await upsertMessengerCandidate(req.dbUser.id, {
            senderId,
            rawText: pending.rawTexts.join('\n'),
            extracted: pending.extracted,
          });
        }
        pendingWebhookCandidates.clear();
      }
      const data = await loadUserCrmData(req.dbUser.id);
      res.json({ ...data, pendingWebhookCandidates: [], storageMode: 'cloudsql' });
    } catch (err) {
      console.warn(
        '[/api/crm/state] Falling back to local mode due to DB error:',
        err instanceof Error ? err.message : err
      );
      res.json({
        candidates: [],
        schoolStudents: [],
        kyujins: [],
        ptPartners: [],
        resignations: [],
        pendingWebhookCandidates: bufferedCandidates,
        hasExistingData: false,
        storageMode: 'local_fallback',
      });
    }
  });

  app.put('/api/crm/state', requireAuth, async (req: AuthRequest, res) => {
    try {
      const { candidates, schoolStudents, kyujins, ptPartners, resignations } =
        req.body ?? {};
      if (
        !Array.isArray(candidates) ||
        !Array.isArray(schoolStudents) ||
        !Array.isArray(kyujins) ||
        !Array.isArray(ptPartners)
      ) {
        res.status(400).json({ error: 'Invalid CRM snapshot payload.' });
        return;
      }

      if (req.dbAvailable === false || !req.dbUser || req.dbUser.id <= 0) {
        res.json({ ok: true, storageMode: 'local_fallback' });
        return;
      }

      await saveUserCrmSnapshot(req.dbUser.id, {
        candidates,
        schoolStudents,
        kyujins,
        ptPartners,
        resignations: Array.isArray(resignations)
          ? (resignations as ResignationRecord[])
          : undefined,
      });

      res.json({ ok: true, storageMode: 'cloudsql' });
    } catch (err) {
      console.warn(
        '[/api/crm/state PUT] DB write skipped; using local fallback:',
        err instanceof Error ? err.message : err
      );
      res.json({ ok: true, storageMode: 'local_fallback' });
    }
  });

  app.get('/api/chat/history', requireAuth, async (req: AuthRequest, res) => {
    if (req.dbAvailable === false || !req.dbUser || req.dbUser.id <= 0) {
      res.json({ messages: [], storageMode: 'local_fallback' });
      return;
    }

    try {
      const messages = await getUserChatHistory(req.dbUser.id, 'default');
      res.json({ messages, storageMode: 'cloudsql' });
    } catch (err) {
      console.warn(
        '[/api/chat/history] DB read skipped:',
        err instanceof Error ? err.message : err
      );
      res.json({ messages: [], storageMode: 'local_fallback' });
    }
  });

  app.delete('/api/chat/history', requireAuth, async (req: AuthRequest, res) => {
    if (req.dbAvailable === false || !req.dbUser || req.dbUser.id <= 0) {
      res.json({ ok: true, storageMode: 'local_fallback' });
      return;
    }

    try {
      await clearUserChatHistory(req.dbUser.id, 'default');
      res.json({ ok: true, storageMode: 'cloudsql' });
    } catch (err) {
      console.warn(
        '[/api/chat/history DELETE] DB delete skipped:',
        err instanceof Error ? err.message : err
      );
      res.json({ ok: true, storageMode: 'local_fallback' });
    }
  });

  // =========================================================================
  // GOOGLE CALENDAR SYNC ROUTES
  // =========================================================================
  const SAFE_CALENDAR_EVENT_ID_RX = /^[A-Za-z0-9_@.-]{4,256}$/;

  interface CalendarSyncInputItem {
    candidateId: string;
    entityType?: 'candidate' | 'resignation';
    eventTitle: string;
    description: string;
    dateIso?: string;
    startIso: string;
    endIso: string;
    existingEventId?: string;
    reminderMinutes?: number[];
  }

  function mapGoogleCalendarHttpError(
    statusCode: number | undefined,
    rawMessage: string,
    reason = ''
  ): { httpStatus: number; code: string; message: string } {
    const lower = `${rawMessage} ${reason}`.toLowerCase();
    if (statusCode === 401) {
      return {
        httpStatus: 401,
        code: 'CALENDAR_TOKEN_EXPIRED',
        message: 'Phiên xác thực Google Calendar đã hết hạn (401 Unauthorized). Vui lòng đăng nhập lại Google.',
      };
    }
    if (statusCode === 403) {
      if (lower.includes('accessnotconfigured') || lower.includes('has not been used')) {
        return {
          httpStatus: 403,
          code: 'CALENDAR_API_DISABLED',
          message:
            'Lỗi 403: Google Calendar API chưa được kích hoạt trên Google Cloud Console. Vui lòng bật Google Calendar API rồi thử lại.',
        };
      }
      return {
        httpStatus: 403,
        code: 'CALENDAR_FORBIDDEN',
        message:
          'Lỗi 403 (access_denied / insufficientPermissions): Tài khoản Google chưa được thêm vào danh sách Test Users trên Google Cloud Console hoặc chưa cấp quyền https://www.googleapis.com/auth/calendar.events.',
      };
    }
    if (statusCode === 400) {
      return {
        httpStatus: 400,
        code: 'CALENDAR_BAD_REQUEST',
        message: `Dữ liệu ngày giờ gửi lên Google Calendar chưa hợp lệ (400 Invalid Request): ${rawMessage}`,
      };
    }
    return {
      httpStatus: 502,
      code: 'CALENDAR_API_ERROR',
      message: rawMessage || 'Không thể kết nối Google Calendar API.',
    };
  }

  async function upsertSingleGoogleCalendarEvent(
    googleAccessToken: string,
    userId: number,
    item: CalendarSyncInputItem
  ): Promise<{
    candidateId: string;
    eventId: string;
    htmlLink: string;
    action: 'created' | 'updated';
  }> {
    const minutesList =
      Array.isArray(item.reminderMinutes) && item.reminderMinutes.length > 0
        ? item.reminderMinutes
            .map((n) => Math.max(0, Math.min(40320, Math.round(Number(n)))))
            .filter((n) => !Number.isNaN(n))
        : [30, 10];

    let startIso = String(item.startIso || '').trim();
    let endIso = String(item.endIso || '').trim();
    if (!startIso && item.dateIso && /^\d{4}-\d{2}-\d{2}$/.test(item.dateIso)) {
      startIso = `${item.dateIso}T09:00:00+09:00`;
      endIso = `${item.dateIso}T10:00:00+09:00`;
    }
    if (
      Number.isNaN(Date.parse(startIso)) ||
      Number.isNaN(Date.parse(endIso))
    ) {
      const err = new Error(
        `Ngày giờ không đúng chuẩn ISO Asia/Tokyo (${startIso || 'trống'})`
      ) as Error & { statusCode?: number };
      err.statusCode = 400;
      throw err;
    }

    const bodyPayload = {
      summary: String(item.eventTitle).slice(0, 300),
      description: String(item.description || '').slice(0, 8000),
      start: {
        dateTime: startIso,
        timeZone: 'Asia/Tokyo',
      },
      end: {
        dateTime: endIso,
        timeZone: 'Asia/Tokyo',
      },
      reminders: {
        useDefault: false,
        overrides: minutesList.map((minutes) => ({
          method: 'popup',
          minutes,
        })),
      },
    };

    const existingId =
      typeof item.existingEventId === 'string' &&
      SAFE_CALENDAR_EVENT_ID_RX.test(item.existingEventId.trim())
        ? item.existingEventId.trim()
        : '';

    let targetMethod: 'POST' | 'PATCH' = existingId ? 'PATCH' : 'POST';
    let targetUrl = existingId
      ? `https://www.googleapis.com/calendar/v3/calendars/primary/events/${encodeURIComponent(existingId)}`
      : 'https://www.googleapis.com/calendar/v3/calendars/primary/events';

    let calResp = await fetch(targetUrl, {
      method: targetMethod,
      headers: {
        Authorization: `Bearer ${googleAccessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(bodyPayload),
    });

    // If previously linked event was manually removed in Google Calendar, create a new one
    if (existingId && (calResp.status === 404 || calResp.status === 410)) {
      targetMethod = 'POST';
      targetUrl = 'https://www.googleapis.com/calendar/v3/calendars/primary/events';
      calResp = await fetch(targetUrl, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${googleAccessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(bodyPayload),
      });
    }

    if (!calResp.ok) {
      const errJson = await calResp.json().catch(() => ({}));
      const errMsg =
        typeof errJson?.error?.message === 'string'
          ? errJson.error.message
          : `Google Calendar HTTP ${calResp.status}`;
      const errReason =
        typeof errJson?.error?.errors?.[0]?.reason === 'string'
          ? errJson.error.errors[0].reason
          : typeof errJson?.error?.status === 'string'
          ? errJson.error.status
          : '';
      const error = new Error(errMsg) as Error & { statusCode?: number; reason?: string };
      error.statusCode = calResp.status;
      error.reason = errReason;
      throw error;
    }

    const createdOrUpdated = (await calResp.json()) as {
      id?: string;
      htmlLink?: string;
    };
    const eventId = String(createdOrUpdated.id || '');
    const htmlLink = String(createdOrUpdated.htmlLink || '');

    if (eventId && userId > 0) {
      try {
        if (item.entityType === 'resignation') {
          await updateResignationCalendarEvent(
            userId,
            item.candidateId,
            eventId,
            htmlLink || null
          );
        } else {
          await updateCandidateCalendarEvent(
            userId,
            item.candidateId,
            eventId,
            htmlLink || null
          );
        }
      } catch (dbErr) {
        console.warn(
          '[Google Calendar] DB link update skipped:',
          dbErr instanceof Error ? dbErr.message : dbErr
        );
      }
    }

    return {
      candidateId: item.candidateId,
      eventId,
      htmlLink,
      action: targetMethod === 'PATCH' ? 'updated' : 'created',
    };
  }

  app.post('/api/calendar/sync-candidate', requireAuth, async (req: AuthRequest, res) => {
    const googleAccessToken = String(
      req.headers['x-google-access-token'] || ''
    ).trim();
    if (!googleAccessToken) {
      res.status(401).json({
        error: 'Thiếu quyền truy cập Google Calendar (chưa có Access Token). Vui lòng kết nối lại Google.',
        code: 'MISSING_CALENDAR_TOKEN',
      });
      return;
    }

    const {
      candidateId,
      entityType,
      eventTitle,
      description,
      dateIso,
      startIso,
      endIso,
      existingEventId,
      reminderMinutes,
    } = (req.body ?? {}) as Partial<CalendarSyncInputItem>;

    if (
      typeof candidateId !== 'string' ||
      !candidateId.trim() ||
      typeof eventTitle !== 'string' ||
      !eventTitle.trim() ||
      (typeof startIso !== 'string' && typeof dateIso !== 'string')
    ) {
      res.status(400).json({ error: 'Thiếu thông tin lịch hẹn để đồng bộ Google Calendar.' });
      return;
    }

    try {
      const result = await upsertSingleGoogleCalendarEvent(
        googleAccessToken,
        req.dbUser!.id,
        {
          candidateId: candidateId.trim(),
          entityType: entityType === 'resignation' ? 'resignation' : 'candidate',
          eventTitle: eventTitle.trim(),
          description: typeof description === 'string' ? description : '',
          dateIso: typeof dateIso === 'string' ? dateIso.trim() : undefined,
          startIso: typeof startIso === 'string' ? startIso.trim() : '',
          endIso: typeof endIso === 'string' ? endIso.trim() : '',
          existingEventId:
            typeof existingEventId === 'string' ? existingEventId : undefined,
          reminderMinutes: Array.isArray(reminderMinutes)
            ? reminderMinutes
            : undefined,
        }
      );
      res.json(result);
    } catch (err) {
      const status = (err as { statusCode?: number }).statusCode;
      const reason = (err as { reason?: string }).reason || '';
      const msg = err instanceof Error ? err.message : 'Calendar sync failed';
      const mapped = mapGoogleCalendarHttpError(status, msg, reason);
      res.status(mapped.httpStatus).json({ error: mapped.message, code: mapped.code });
    }
  });

  app.post('/api/calendar/sync-bulk', requireAuth, async (req: AuthRequest, res) => {
    const googleAccessToken = String(
      req.headers['x-google-access-token'] || ''
    ).trim();
    if (!googleAccessToken) {
      res.status(401).json({
        error: 'Thiếu quyền truy cập Google Calendar (chưa có Access Token). Vui lòng kết nối lại Google.',
        code: 'MISSING_CALENDAR_TOKEN',
      });
      return;
    }

    const rawItems = req.body?.items;
    if (!Array.isArray(rawItems) || rawItems.length === 0) {
      res.status(400).json({ error: 'Không có lịch hẹn hợp lệ để đồng bộ.' });
      return;
    }

    const synced: Array<{
      candidateId: string;
      eventId: string;
      htmlLink: string;
      action: 'created' | 'updated';
    }> = [];
    const failed: Array<{ candidateId: string; error: string }> = [];

    for (const rawItem of rawItems.slice(0, 50) as Partial<CalendarSyncInputItem>[]) {
      if (
        !rawItem ||
        typeof rawItem.candidateId !== 'string' ||
        typeof rawItem.eventTitle !== 'string' ||
        (typeof rawItem.startIso !== 'string' && typeof rawItem.dateIso !== 'string')
      ) {
        continue;
      }
      try {
        const result = await upsertSingleGoogleCalendarEvent(
          googleAccessToken,
          req.dbUser!.id,
          {
            candidateId: rawItem.candidateId.trim(),
            entityType:
              rawItem.entityType === 'resignation' ? 'resignation' : 'candidate',
            eventTitle: rawItem.eventTitle.trim(),
            description:
              typeof rawItem.description === 'string' ? rawItem.description : '',
            dateIso: typeof rawItem.dateIso === 'string' ? rawItem.dateIso.trim() : undefined,
            startIso: typeof rawItem.startIso === 'string' ? rawItem.startIso.trim() : '',
            endIso: typeof rawItem.endIso === 'string' ? rawItem.endIso.trim() : '',
            existingEventId:
              typeof rawItem.existingEventId === 'string'
                ? rawItem.existingEventId
                : undefined,
            reminderMinutes: Array.isArray(rawItem.reminderMinutes)
              ? rawItem.reminderMinutes
              : undefined,
          }
        );
        synced.push(result);
      } catch (err) {
        const status = (err as { statusCode?: number }).statusCode;
        const reason = (err as { reason?: string }).reason || '';
        const msg = err instanceof Error ? err.message : 'Sync failed';
        const mapped = mapGoogleCalendarHttpError(status, msg, reason);
        if (status === 401 || status === 403) {
          res.status(mapped.httpStatus).json({
            error: mapped.message,
            code: mapped.code,
            partialSynced: synced,
          });
          return;
        }
        failed.push({
          candidateId: rawItem.candidateId,
          error: mapped.message,
        });
      }
    }

    res.json({ synced, failed });
  });

  app.delete(
    '/api/calendar/events/:eventId',
    requireAuth,
    async (req: AuthRequest, res) => {
      const eventId = String(req.params.eventId || '').trim();
      if (!SAFE_CALENDAR_EVENT_ID_RX.test(eventId)) {
        res.status(400).json({ error: 'Invalid calendar event ID.' });
        return;
      }

      const candidateId =
        typeof req.query.candidateId === 'string' ? req.query.candidateId.trim() : '';
      const entityType =
        req.query.entityType === 'resignation' ? 'resignation' : 'candidate';

      const googleAccessToken = String(
        req.headers['x-google-access-token'] || ''
      ).trim();
      if (!googleAccessToken) {
        res.status(401).json({
          error: 'Thiếu quyền truy cập Google Calendar. Vui lòng kết nối lại Google.',
          code: 'MISSING_CALENDAR_TOKEN',
        });
        return;
      }

      try {
        const delResp = await fetch(
          `https://www.googleapis.com/calendar/v3/calendars/primary/events/${encodeURIComponent(
            eventId
          )}`,
          {
            method: 'DELETE',
            headers: {
              Authorization: `Bearer ${googleAccessToken}`,
            },
          }
        );

        if (delResp.status === 401 || delResp.status === 403) {
          const errJson = await delResp.json().catch(() => ({}));
          const rawMsg =
            typeof errJson?.error?.message === 'string'
              ? errJson.error.message
              : `Google Calendar DELETE HTTP ${delResp.status}`;
          const mapped = mapGoogleCalendarHttpError(delResp.status, rawMsg);
          res.status(mapped.httpStatus).json({ error: mapped.message, code: mapped.code });
          return;
        }

        if (!delResp.ok && delResp.status !== 404 && delResp.status !== 410) {
          const errJson = await delResp.json().catch(() => ({}));
          res.status(502).json({
            error:
              typeof errJson?.error?.message === 'string'
                ? errJson.error.message
                : `Google Calendar DELETE HTTP ${delResp.status}`,
          });
          return;
        }

        if (candidateId && req.dbAvailable !== false && req.dbUser && req.dbUser.id > 0) {
          try {
            if (entityType === 'resignation') {
              await updateResignationCalendarEvent(
                req.dbUser.id,
                candidateId,
                null,
                null
              );
            } else {
              await updateCandidateCalendarEvent(
                req.dbUser.id,
                candidateId,
                null,
                null
              );
            }
          } catch (dbErr) {
            console.warn(
              '[Google Calendar DELETE] DB link cleanup skipped:',
              dbErr instanceof Error ? dbErr.message : dbErr
            );
          }
        }

        res.json({ deleted: true, eventId });
      } catch (err) {
        const msg = err instanceof Error ? err.message : 'Calendar delete failed';
        res.status(500).json({ error: msg });
      }
    }
  );

  app.use(
    '/pdfjs-assets',
    express.static(path.resolve(__dirname, 'node_modules/pdfjs-dist'))
  );

  if (process.env.NODE_ENV !== 'production') {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distDir = path.resolve(__dirname, 'dist');
    app.use(express.static(distDir));
    app.get('*', (_req, res) => {
      res.sendFile(path.join(distDir, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server listening on http://0.0.0.0:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error('Failed to start server:', err);
  process.exit(1);
});
