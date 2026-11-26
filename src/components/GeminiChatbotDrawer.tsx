import React, { useEffect, useRef, useState } from 'react';
import { Copy, MessageSquareText, RotateCcw, Send, X } from 'lucide-react';
import type { UiLang } from '../i18n/crmTranslations';

export type ChatbotRoleId = 'crm_strategist' | 'reply_coach' | 'screening_advisor';

export interface ChatTurnMessage {
  id: string;
  role: 'user' | 'model';
  assistantRole: ChatbotRoleId;
  modelUsed: string;
  content: string;
  createdAt?: string;
}

interface GeminiChatbotDrawerProps {
  lang?: UiLang;
  isOpen: boolean;
  onClose: () => void;
  getIdToken: () => Promise<string | null>;
  userEmail: string | null;
  isSignedIn: boolean;
  crmSummary: string;
  onCopySuccess?: (msg: string) => void;
}

const ROLE_OPTIONS_BY_LANG: Record<
  UiLang,
  Array<{
    id: ChatbotRoleId;
    label: string;
    subtitle: string;
    modelHint: string;
    prompts: string[];
  }>
> = {
  vi: [
    {
      id: 'crm_strategist',
      label: 'Tư vấn CRM',
      subtitle: 'Phân tích ứng viên, ưu tiên follow-up & chiến lược ghép đơn',
      modelHint: 'gemini-4-hydrogen',
      prompts: [
        'Đề xuất thứ tự ưu tiên chăm sóc nhóm ứng viên đang có trong CRM.',
        'Những tiêu chí quan trọng nhất khi ghép ứng viên Kaigo N4 vào cơ sở phù hợp?',
        'Gợi ý quy trình giảm tỷ lệ từ chối sau khi gửi hồ sơ cho xí nghiệp.',
      ],
    },
    {
      id: 'reply_coach',
      label: 'Soạn tin nhanh',
      subtitle: 'Soạn tin nhắn Messenger/Zalo & email Nhật-Việt sẵn sàng gửi ngay',
      modelHint: 'gemini-3.1-flash-lite',
      prompts: [
        'Soạn tin nhắn tiếng Việt + tiếng Nhật nhắc ứng viên lịch Mendan ngày mai lúc 14:00.',
        'Viết mẫu tin nhắn hỏi thăm nhẹ nhàng bằng tiếng Việt cho ứng viên im lặng 5 ngày.',
        'Viết email ngắn bằng tiếng Nhật đề xuất ứng viên JLPT N3 kinh nghiệm Kaigo cho xí nghiệp.',
      ],
    },
    {
      id: 'screening_advisor',
      label: 'Đánh giá hồ sơ',
      subtitle: 'Kiểm tra thiếu sót visa/JLPT & gợi ý câu hỏi phỏng vấn sơ vấn',
      modelHint: 'gemini-4-hydrogen',
      prompts: [
        'Liệt kê 5 câu hỏi kiểm tra động lực và khả năng làm việc lâu dài khi Mendan thực tập sinh chuyển Tokutei.',
        'Du học sinh sắp tốt nghiệp trường tiếng cần chuẩn bị giấy tờ gì để chuyển sang Tokutei Ginou?',
        'Cần kiểm tra những điểm nào trên thẻ ngoại kiều trước khi tiến cử ứng viên?',
      ],
    },
  ],
  jp: [
    {
      id: 'crm_strategist',
      label: 'CRM戦略相談',
      subtitle: '求職者分析・フォロー優先順位・求人マッチング戦略を提案',
      modelHint: 'gemini-4-hydrogen',
      prompts: [
        '現在のCRMデータをもとに、優先してフォローすべき求職者の順序を提案してください。',
        '介護N4の求職者を最適な施設へマッチングする際の重要チェックポイントは？',
        '書類推薦後の辞退率を下げるための事前すり合わせ手順を提案してください。',
      ],
    },
    {
      id: 'reply_coach',
      label: '返信文作成',
      subtitle: 'Messenger/Zalo・企業向け日越メールの文面を即座に作成',
      modelHint: 'gemini-3.1-flash-lite',
      prompts: [
        '明日14:00の面談リマインド文をベトナム語＋日本語で作成してください。',
        '5日間返信がない求職者向けの丁寧なフォロー連絡文（ベトナム語）を作成してください。',
        'JLPT N3・介護経験ありの求職者を受入企業へ推薦する短い日本語メールを作成してください。',
      ],
    },
    {
      id: 'screening_advisor',
      label: '書類・面談助言',
      subtitle: 'ビザ・JLPT不足チェックと一次面談のスクリーニング質問を提案',
      modelHint: 'gemini-4-hydrogen',
      prompts: [
        '技能実習生から特定技能へ移行する方の定着意欲を確認する面談質問を5つ挙げてください。',
        '日本語学校卒業見込みの留学生が特定技能へ変更する際の必要書類は？',
        '企業推薦前に在留カードで必ず確認すべきポイントを整理してください。',
      ],
    },
  ],
};

const DRAWER_COPY = {
  vi: {
    ariaLabel: 'Trợ lý AI Gemini',
    headerTitle: 'Trợ lý AI Gemini · Tokutei CRM',
    newChatTitle: 'Bắt đầu hội thoại mới',
    closeAria: 'Đóng Trợ lý AI',
    attachSummary: 'Đính kèm số liệu CRM hiện tại',
    needLogin: 'Vui lòng đăng nhập Google để sử dụng Trợ lý AI Gemini.',
    needLoginDesc:
      'Sau khi đăng nhập, bạn có thể hỏi đáp với AI theo ngữ cảnh hồ sơ ứng viên trong CRM và lưu lịch sử trò chuyện.',
    loadingHistory: 'Đang tải lịch sử hội thoại từ Cloud SQL...',
    selectedRolePrefix: 'Vai trò đang chọn:',
    autoSavedSuffix: ' và tự động lưu vào Cloud SQL',
    roleSessionDesc: (sub: string, hasEmail: boolean) =>
      `${sub}. Lịch sử trò chuyện nhiều lượt được duy trì liên tục trong phiên làm việc${
        hasEmail ? ' và tự động lưu vào Cloud SQL' : ''
      }.`,
    quickStartLabel: 'Gợi ý bắt đầu nhanh:',
    copyBtn: 'Sao chép',
    copiedToast: 'Đã sao chép phản hồi AI',
    copyFailToast: 'Không thể truy cập clipboard trên trình duyệt này',
    typingMsg: 'Đang soạn câu trả lời...',
    errorTitle: 'Không thể kết nối Trợ lý AI',
    fallbackErr:
      'Không thể nhận phản hồi từ Gemini. Vui lòng kiểm tra GEMINI_API_KEY trong Settings > Secrets.',
    httpErr: (s: number) => `Lỗi kết nối AI (HTTP ${s})`,
    placeholderSignedIn: 'Nhập câu hỏi về ứng viên, ghép đơn hoặc nhờ soạn tin nhắn...',
    placeholderGuest: 'Đăng nhập Google để gửi câu hỏi...',
    sendAria: 'Gửi tin nhắn',
    sendBtn: 'Gửi',
    shortcutHint: 'Enter để gửi · Shift + Enter xuống dòng',
  },
  jp: {
    ariaLabel: 'Gemini AIアシスタント',
    headerTitle: 'Gemini AIアシスタント · Tokutei CRM',
    newChatTitle: '新しい会話を始める',
    closeAria: 'AIアシスタントを閉じる',
    attachSummary: '現在のCRM集計データを添付する',
    needLogin: 'Gemini AIアシスタントを利用するにはGoogleでログインしてください。',
    needLoginDesc:
      'ログイン後、CRM内の求職者・求人データを踏まえたAI相談や会話履歴の保存が可能になります。',
    loadingHistory: 'Cloud SQLから会話履歴を読み込み中...',
    selectedRolePrefix: '選択中の役割:',
    autoSavedSuffix: '（Cloud SQLに自動保存）',
    roleSessionDesc: (sub: string, hasEmail: boolean) =>
      `${sub}。セッション中の会話文脈は維持されます${
        hasEmail ? '（Cloud SQLに自動保存）' : ''
      }。`,
    quickStartLabel: 'クイック質問テンプレート:',
    copyBtn: 'コピー',
    copiedToast: 'AIの回答をコピーしました',
    copyFailToast: 'クリップボードにアクセスできませんでした',
    typingMsg: '回答を作成中...',
    errorTitle: 'AIアシスタントに接続できません',
    fallbackErr:
      'Geminiから応答を取得できませんでした。Settings > Secrets の GEMINI_API_KEY を確認してください。',
    httpErr: (s: number) => `AI接続エラー (HTTP ${s})`,
    placeholderSignedIn: '求職者フォロー、マッチング、返信文の作成などを入力...',
    placeholderGuest: '質問を送信するにはGoogleでログインしてください...',
    sendAria: 'メッセージを送信',
    sendBtn: '送信',
    shortcutHint: 'Enterで送信 · Shift + Enterで改行',
  },
} as const;

export function GeminiChatbotDrawer({
  lang = 'vi',
  isOpen,
  onClose,
  getIdToken,
  userEmail,
  isSignedIn,
  crmSummary,
  onCopySuccess,
}: GeminiChatbotDrawerProps) {
  const [selectedRole, setSelectedRole] = useState<ChatbotRoleId>('crm_strategist');
  const [includeCrmContext, setIncludeCrmContext] = useState(true);
  const [messages, setMessages] = useState<ChatTurnMessage[]>([]);
  const [inputText, setInputText] = useState('');
  const [isSending, setIsSending] = useState(false);
  const [isLoadingHistory, setIsLoadingHistory] = useState(false);
  const [errorMsg, setErrorMsg] = useState('');
  const threadEndRef = useRef<HTMLDivElement | null>(null);
  const loadedForUserRef = useRef<string | null>(null);

  const roleOptions = ROLE_OPTIONS_BY_LANG[lang];
  const copy = DRAWER_COPY[lang];
  const activeRoleMeta =
    roleOptions.find((r) => r.id === selectedRole) ?? roleOptions[0];

  useEffect(() => {
    if (!isOpen) return;
    threadEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, isOpen]);

  useEffect(() => {
    let cancelled = false;
    if (!userEmail || !isSignedIn) {
      loadedForUserRef.current = null;
      if (!isSignedIn) {
        setMessages([]);
      }
      return;
    }
    if (loadedForUserRef.current === userEmail) return;

    const loadHistory = async () => {
      setIsLoadingHistory(true);
      try {
        const token = await getIdToken();
        if (!token || cancelled) return;
        const resp = await fetch('/api/chat/history', {
          headers: { Authorization: `Bearer ${token}` },
        });
        if (!resp.ok) return;
        const data = (await resp.json()) as {
          messages?: Array<{
            id: number;
            role: 'user' | 'model';
            assistantRole: string;
            modelUsed: string;
            content: string;
            createdAt?: string;
          }>;
        };
        if (!cancelled && Array.isArray(data.messages)) {
          setMessages(
            data.messages.map((m) => ({
              id: `db_${m.id}`,
              role: m.role,
              assistantRole:
                m.assistantRole === 'reply_coach' ||
                m.assistantRole === 'screening_advisor'
                  ? m.assistantRole
                  : 'crm_strategist',
              modelUsed: m.modelUsed,
              content: m.content,
              createdAt: m.createdAt,
            }))
          );
          loadedForUserRef.current = userEmail;
        }
      } catch {
        // Ignore silent history fetch failures until user interacts
      } finally {
        if (!cancelled) setIsLoadingHistory(false);
      }
    };

    loadHistory();
    return () => {
      cancelled = true;
    };
  }, [userEmail, isSignedIn, getIdToken]);

  const handleSendMessage = async (overrideText?: string) => {
    const trimmed = (overrideText ?? inputText).trim();
    if (!trimmed || isSending) return;
    if (!isSignedIn) {
      setErrorMsg(copy.needLogin);
      return;
    }

    setErrorMsg('');
    if (!overrideText) setInputText('');

    const optimisticUserMsg: ChatTurnMessage = {
      id: `local_u_${Date.now()}`,
      role: 'user',
      assistantRole: selectedRole,
      modelUsed: activeRoleMeta.modelHint,
      content: trimmed,
    };

    const historyPayload = messages.slice(-20).map((m) => ({
      role: m.role,
      text: m.content,
    }));

    setMessages((prev) => [...prev, optimisticUserMsg]);
    setIsSending(true);

    try {
      const token = userEmail ? await getIdToken() : null;
      const headers: Record<string, string> = {
        'Content-Type': 'application/json',
      };
      if (token) {
        headers.Authorization = `Bearer ${token}`;
      }

      const resp = await fetch('/api/gemini/chat', {
        method: 'POST',
        headers,
        body: JSON.stringify({
          message: trimmed,
          history: historyPayload,
          roleId: selectedRole,
          crmSummary: includeCrmContext ? crmSummary : undefined,
        }),
      });

      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) {
        throw new Error(
          typeof data?.error === 'string' && data.error
            ? data.error
            : copy.httpErr(resp.status)
        );
      }

      const modelReply: ChatTurnMessage = {
        id: `local_m_${Date.now()}`,
        role: 'model',
        assistantRole: selectedRole,
        modelUsed: data.modelUsed || activeRoleMeta.modelHint,
        content: data.reply,
      };
      setMessages((prev) => [...prev, modelReply]);
    } catch (err) {
      const msg = err instanceof Error ? err.message : copy.fallbackErr;
      setErrorMsg(msg);
      setMessages((prev) => prev.filter((m) => m.id !== optimisticUserMsg.id));
      if (!overrideText) setInputText(trimmed);
    } finally {
      setIsSending(false);
    }
  };

  const handleClearConversation = async () => {
    setErrorMsg('');
    setMessages([]);
    if (userEmail) {
      try {
        const token = await getIdToken();
        if (token) {
          await fetch('/api/chat/history', {
            method: 'DELETE',
            headers: { Authorization: `Bearer ${token}` },
          });
        }
      } catch {
        // Ignore cleanup failures cleanly
      }
    }
  };

  const handleCopyReply = async (text: string) => {
    try {
      await navigator.clipboard.writeText(text);
      onCopySuccess?.(copy.copiedToast);
    } catch {
      onCopySuccess?.(copy.copyFailToast);
    }
  };

  if (!isOpen) return null;

  return (
    <div
      className="fixed inset-0 z-50 flex justify-end bg-slate-900/40 backdrop-blur-[2px]"
      role="dialog"
      aria-modal="true"
      aria-label={copy.ariaLabel}
    >
      <div className="flex h-full w-full max-w-lg flex-col border-l border-[#E2E8F0] bg-[#F8FAFC] text-[#0F172A] shadow-2xl">
        {/* Drawer Header */}
        <div className="flex items-center justify-between border-b border-[#E2E8F0] bg-white px-5 py-3.5">
          <div className="flex items-center gap-2.5 min-w-0">
            <MessageSquareText className="h-5 w-5 shrink-0 text-[#16A34A]" aria-hidden="true" />
            <div className="min-w-0">
              <h2 className="text-sm font-bold text-[#0F172A] truncate">
                {copy.headerTitle}
              </h2>
              <p className="text-[11px] text-[#64748B] truncate">
                {activeRoleMeta.subtitle}
              </p>
            </div>
          </div>

          <div className="flex items-center gap-1.5 shrink-0">
            <button
              type="button"
              onClick={handleClearConversation}
              title={copy.newChatTitle}
              className="rounded-md p-1.5 text-[#64748B] transition-colors hover:bg-slate-100 hover:text-[#0F172A] cursor-pointer"
            >
              <RotateCcw className="h-4 w-4" aria-hidden="true" />
            </button>
            <button
              type="button"
              onClick={onClose}
              aria-label={copy.closeAria}
              className="rounded-md p-1.5 text-[#64748B] transition-colors hover:bg-slate-100 hover:text-[#0F172A] cursor-pointer"
            >
              <X className="h-4 w-4" aria-hidden="true" />
            </button>
          </div>
        </div>

        {/* Role Selector + Context Controls */}
        <div className="border-b border-[#E2E8F0] bg-white px-5 py-3 space-y-2.5">
          <div className="grid grid-cols-3 gap-1 rounded-md border border-[#E2E8F0] bg-slate-100 p-1">
            {roleOptions.map((role) => (
              <button
                key={role.id}
                type="button"
                onClick={() => setSelectedRole(role.id)}
                className={`rounded px-2 py-1.5 text-[11px] font-semibold transition-all whitespace-nowrap truncate cursor-pointer ${
                  selectedRole === role.id
                    ? 'bg-white text-[#16A34A] shadow-sm'
                    : 'text-[#64748B] hover:text-[#0F172A]'
                }`}
              >
                {role.label}
              </button>
            ))}
          </div>

          <div className="flex items-center justify-between gap-2 text-[11px] text-[#64748B]">
            <label className="inline-flex items-center gap-1.5 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={includeCrmContext}
                onChange={(e) => setIncludeCrmContext(e.target.checked)}
                className="h-3.5 w-3.5 rounded border-[#CBD5E1] bg-white accent-[#16A34A]"
              />
              <span>{copy.attachSummary}</span>
            </label>
            <span className="font-mono text-[10px] font-medium text-[#16A34A]">
              {activeRoleMeta.modelHint}
            </span>
          </div>
        </div>

        {/* Scrollable Message Thread */}
        <div className="flex-1 overflow-y-auto px-5 py-4 space-y-4">
          {!isSignedIn ? (
            <div className="rounded-lg border border-[#E2E8F0] bg-white p-5 text-xs text-[#0F172A] shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
              <p className="font-semibold text-[#16A34A]">{copy.needLogin}</p>
              <p className="mt-2 text-[#64748B]">{copy.needLoginDesc}</p>
            </div>
          ) : isLoadingHistory ? (
            <div className="py-10 text-center text-xs text-[#64748B]">{copy.loadingHistory}</div>
          ) : messages.length === 0 ? (
            <div className="space-y-4 py-4">
              <div className="rounded-lg border border-[#E2E8F0] bg-white p-4 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
                <p className="text-xs font-semibold text-[#0F172A]">
                  {copy.selectedRolePrefix} {activeRoleMeta.label}
                </p>
                <p className="mt-1 text-xs leading-relaxed text-[#64748B]">
                  {copy.roleSessionDesc(activeRoleMeta.subtitle, Boolean(userEmail))}
                </p>
              </div>

              <div>
                <p className="mb-2 font-mono text-[11px] font-semibold text-[#64748B]">
                  {copy.quickStartLabel}
                </p>
                <div className="space-y-2">
                  {activeRoleMeta.prompts.map((promptText, idx) => (
                    <button
                      key={idx}
                      type="button"
                      onClick={() => handleSendMessage(promptText)}
                      className="w-full rounded-lg border border-[#E2E8F0] bg-white px-3.5 py-2.5 text-left text-xs text-[#0F172A] shadow-[0_1px_3px_rgba(0,0,0,0.05)] transition-colors hover:border-[#16A34A] hover:bg-emerald-50/40 cursor-pointer"
                    >
                      {promptText}
                    </button>
                  ))}
                </div>
              </div>
            </div>
          ) : (
            messages.map((m) => (
              <div
                key={m.id}
                className={`flex flex-col ${
                  m.role === 'user' ? 'items-end' : 'items-start'
                }`}
              >
                <div
                  className={`max-w-[88%] rounded-lg px-4 py-3 text-xs leading-relaxed whitespace-pre-wrap ${
                    m.role === 'user'
                      ? 'bg-[#16A34A] font-medium text-white shadow-sm'
                      : 'border border-[#E2E8F0] bg-white text-[#0F172A] shadow-[0_1px_3px_rgba(0,0,0,0.05)]'
                  }`}
                >
                  {m.content}
                </div>
                <div className="mt-1 flex items-center gap-2 px-1 font-mono text-[10px] text-[#64748B]">
                  <span>{m.role === 'user' ? 'USER' : m.modelUsed}</span>
                  {m.role === 'model' && (
                    <button
                      type="button"
                      onClick={() => handleCopyReply(m.content)}
                      className="inline-flex items-center gap-1 text-[#64748B] hover:text-[#16A34A] cursor-pointer"
                    >
                      <Copy className="h-3 w-3" aria-hidden="true" />
                      {copy.copyBtn}
                    </button>
                  )}
                </div>
              </div>
            ))
          )}

          {isSending && (
            <div className="flex items-start">
              <div className="rounded-lg border border-[#E2E8F0] bg-white px-4 py-3 font-mono text-xs text-[#16A34A] shadow-sm">
                {copy.typingMsg}
              </div>
            </div>
          )}

          {errorMsg && (
            <div className="rounded-lg border border-rose-200 bg-rose-50 px-4 py-3 text-xs text-rose-800">
              <p className="font-semibold text-rose-700">{copy.errorTitle}</p>
              <p className="mt-1">{errorMsg}</p>
            </div>
          )}

          <div ref={threadEndRef} />
        </div>

        {/* Message Composer */}
        <form
          onSubmit={(e) => {
            e.preventDefault();
            handleSendMessage();
          }}
          className="border-t border-[#E2E8F0] bg-white px-4 py-3"
        >
          <div className="flex items-end gap-2">
            <textarea
              rows={2}
              value={inputText}
              onChange={(e) => setInputText(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' && !e.shiftKey) {
                  e.preventDefault();
                  handleSendMessage();
                }
              }}
              disabled={!isSignedIn}
              placeholder={isSignedIn ? copy.placeholderSignedIn : copy.placeholderGuest}
              className="flex-1 resize-none rounded-md border border-[#CBD5E1] bg-[#F8FAFC] px-3 py-2 text-xs text-[#0F172A] placeholder:text-[#64748B] focus:border-[#16A34A] focus:bg-white focus:outline-none disabled:opacity-50"
            />
            <button
              type="submit"
              disabled={!isSignedIn || isSending || !inputText.trim()}
              aria-label={copy.sendAria}
              className="inline-flex h-10 items-center justify-center gap-1.5 rounded-md bg-[#16A34A] px-4 text-xs font-semibold text-white transition-colors hover:bg-[#15803D] disabled:cursor-not-allowed disabled:opacity-50 cursor-pointer"
            >
              <Send className="h-3.5 w-3.5" aria-hidden="true" />
              {copy.sendBtn}
            </button>
          </div>
          <p className="mt-1.5 font-mono text-[10px] text-[#64748B]">{copy.shortcutHint}</p>
        </form>
      </div>
    </div>
  );
}
