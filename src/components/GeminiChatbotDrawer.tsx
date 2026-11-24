import React, { useEffect, useRef, useState } from 'react';
import { Copy, MessageSquareText, RotateCcw, Send, X } from 'lucide-react';

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
  isOpen: boolean;
  onClose: () => void;
  getIdToken: () => Promise<string | null>;
  userEmail: string | null;
  isSignedIn: boolean;
  crmSummary: string;
  onCopySuccess?: (msg: string) => void;
}

const ROLE_OPTIONS: Array<{
  id: ChatbotRoleId;
  label: string;
  subtitle: string;
  modelHint: string;
  prompts: string[];
}> = [
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
];

export function GeminiChatbotDrawer({
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

  const activeRoleMeta =
    ROLE_OPTIONS.find((r) => r.id === selectedRole) ?? ROLE_OPTIONS[0];

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
      setErrorMsg('Vui lòng đăng nhập Google để sử dụng Trợ lý AI Gemini.');
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
            : `Lỗi kết nối AI (HTTP ${resp.status})`
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
      const msg =
        err instanceof Error
          ? err.message
          : 'Không thể nhận phản hồi từ Gemini. Vui lòng kiểm tra GEMINI_API_KEY trong Settings > Secrets.';
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
      onCopySuccess?.('Đã sao chép phản hồi AI');
    } catch {
      onCopySuccess?.('Không thể truy cập clipboard trên trình duyệt này');
    }
  };

  if (!isOpen) return null;

  return (
    <div
      className="fixed inset-0 z-50 flex justify-end bg-black/70 backdrop-blur-[2px]"
      role="dialog"
      aria-modal="true"
      aria-label="Trợ lý AI Gemini"
    >
      <div className="flex h-full w-full max-w-lg flex-col border-l border-[#27272a] bg-[#0c0c0e] text-[#edeeef] shadow-2xl">
        {/* Drawer Header */}
        <div className="flex items-center justify-between border-b border-[#27272a] bg-[#141417] px-5 py-3.5">
          <div className="flex items-center gap-2.5 min-w-0">
            <MessageSquareText className="h-5 w-5 shrink-0 text-[#22c55e]" aria-hidden="true" />
            <div className="min-w-0">
              <h2 className="text-sm font-bold text-[#edeeef] truncate">
                Trợ lý AI Gemini · Tokutei CRM
              </h2>
              <p className="text-[11px] text-[#94949e] truncate">
                {activeRoleMeta.subtitle}
              </p>
            </div>
          </div>

          <div className="flex items-center gap-1.5 shrink-0">
            <button
              type="button"
              onClick={handleClearConversation}
              title="Bắt đầu hội thoại mới"
              className="rounded-md p-1.5 text-[#94949e] transition-colors hover:bg-[#27272a] hover:text-[#edeeef] cursor-pointer"
            >
              <RotateCcw className="h-4 w-4" aria-hidden="true" />
            </button>
            <button
              type="button"
              onClick={onClose}
              aria-label="Đóng Trợ lý AI"
              className="rounded-md p-1.5 text-[#94949e] transition-colors hover:bg-[#27272a] hover:text-[#edeeef] cursor-pointer"
            >
              <X className="h-4 w-4" aria-hidden="true" />
            </button>
          </div>
        </div>

        {/* Role Selector + Context Controls */}
        <div className="border-b border-[#27272a] bg-[#141417] px-5 py-3 space-y-2.5">
          <div className="grid grid-cols-3 gap-1 rounded-md border border-[#27272a] bg-[#0c0c0e] p-1">
            {ROLE_OPTIONS.map((role) => (
              <button
                key={role.id}
                type="button"
                onClick={() => setSelectedRole(role.id)}
                className={`rounded px-2 py-1.5 text-[11px] font-semibold transition-colors whitespace-nowrap truncate cursor-pointer ${
                  selectedRole === role.id
                    ? 'bg-[#22c55e] text-[#0c0c0e]'
                    : 'text-[#94949e] hover:text-[#edeeef]'
                }`}
              >
                {role.label}
              </button>
            ))}
          </div>

          <div className="flex items-center justify-between gap-2 text-[11px] text-[#94949e]">
            <label className="inline-flex items-center gap-1.5 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={includeCrmContext}
                onChange={(e) => setIncludeCrmContext(e.target.checked)}
                className="h-3.5 w-3.5 rounded border-[#27272a] bg-[#0c0c0e] accent-[#22c55e]"
              />
              <span>Đính kèm số liệu CRM hiện tại</span>
            </label>
            <span className="font-mono text-[10px] text-[#22c55e]">
              {activeRoleMeta.modelHint}
            </span>
          </div>
        </div>

        {/* Scrollable Message Thread */}
        <div className="flex-1 overflow-y-auto px-5 py-4 space-y-4">
          {!isSignedIn ? (
            <div className="rounded-lg border border-[#27272a] bg-[#141417] p-5 text-xs text-[#edeeef]">
              <p className="font-semibold text-[#22c55e]">
                Vui lòng đăng nhập tài khoản Google để sử dụng Trợ lý AI Gemini.
              </p>
              <p className="mt-2 text-[#94949e]">
                Sau khi đăng nhập, bạn có thể hỏi đáp với AI theo ngữ cảnh hồ sơ ứng viên trong CRM và lưu lịch sử trò chuyện.
              </p>
            </div>
          ) : isLoadingHistory ? (
            <div className="py-10 text-center text-xs text-[#94949e]">
              Đang tải lịch sử hội thoại từ Cloud SQL...
            </div>
          ) : messages.length === 0 ? (
            <div className="space-y-4 py-4">
              <div className="rounded-lg border border-[#27272a] bg-[#141417] p-4">
                <p className="text-xs font-semibold text-[#edeeef]">
                  Vai trò đang chọn: {activeRoleMeta.label}
                </p>
                <p className="mt-1 text-xs leading-relaxed text-[#94949e]">
                  {activeRoleMeta.subtitle}. Lịch sử trò chuyện nhiều lượt được duy trì liên tục trong phiên làm việc{userEmail ? ' và tự động lưu vào Cloud SQL' : ''}.
                </p>
              </div>

              <div>
                <p className="mb-2 font-mono text-[11px] font-semibold text-[#94949e]">
                  Gợi ý bắt đầu nhanh:
                </p>
                <div className="space-y-2">
                  {activeRoleMeta.prompts.map((promptText, idx) => (
                    <button
                      key={idx}
                      type="button"
                      onClick={() => handleSendMessage(promptText)}
                      className="w-full rounded-lg border border-[#27272a] bg-[#141417] px-3.5 py-2.5 text-left text-xs text-[#edeeef] transition-colors hover:border-[#22c55e] cursor-pointer"
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
                      ? 'bg-[#22c55e] font-medium text-[#0c0c0e]'
                      : 'border border-[#27272a] bg-[#141417] text-[#edeeef]'
                  }`}
                >
                  {m.content}
                </div>
                <div className="mt-1 flex items-center gap-2 px-1 font-mono text-[10px] text-[#94949e]">
                  <span>{m.role === 'user' ? 'USER' : m.modelUsed}</span>
                  {m.role === 'model' && (
                    <button
                      type="button"
                      onClick={() => handleCopyReply(m.content)}
                      className="inline-flex items-center gap-1 text-[#94949e] hover:text-[#22c55e] cursor-pointer"
                    >
                      <Copy className="h-3 w-3" aria-hidden="true" />
                      Sao chép
                    </button>
                  )}
                </div>
              </div>
            ))
          )}

          {isSending && (
            <div className="flex items-start">
              <div className="rounded-lg border border-[#27272a] bg-[#141417] px-4 py-3 font-mono text-xs text-[#22c55e]">
                Đang soạn câu trả lời...
              </div>
            </div>
          )}

          {errorMsg && (
            <div className="rounded-lg border border-[#ef4444]/40 bg-[#ef4444]/10 px-4 py-3 text-xs text-rose-300">
              <p className="font-semibold text-[#ef4444]">Không thể kết nối Trợ lý AI</p>
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
          className="border-t border-[#27272a] bg-[#141417] px-4 py-3"
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
              placeholder={
                isSignedIn
                  ? 'Nhập câu hỏi về ứng viên, ghép đơn hoặc nhờ soạn tin nhắn...'
                  : 'Đăng nhập Google để gửi câu hỏi...'
              }
              className="flex-1 resize-none rounded-md border border-[#27272a] bg-[#0c0c0e] px-3 py-2 text-xs text-[#edeeef] placeholder:text-[#94949e] focus:border-[#22c55e] focus:outline-none disabled:opacity-50"
            />
            <button
              type="submit"
              disabled={!isSignedIn || isSending || !inputText.trim()}
              aria-label="Gửi tin nhắn"
              className="inline-flex h-10 items-center justify-center gap-1.5 rounded-md bg-[#22c55e] px-4 text-xs font-semibold text-[#0c0c0e] transition-opacity hover:opacity-90 disabled:cursor-not-allowed disabled:opacity-50 cursor-pointer"
            >
              <Send className="h-3.5 w-3.5" aria-hidden="true" />
              Gửi
            </button>
          </div>
          <p className="mt-1.5 font-mono text-[10px] text-[#94949e]">
            Enter để gửi · Shift + Enter xuống dòng
          </p>
        </form>
      </div>
    </div>
  );
}
