import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import { onAuthStateChanged, signOut, type User } from 'firebase/auth';
import {
  Calendar,
  CheckSquare,
  ClipboardPaste,
  Copy,
  ExternalLink,
  FileText,
  Image as ImageIcon,
  Link2,
  MessageCircle,
  MessageSquareText,
  RefreshCw,
  ShieldCheck,
  Sparkles,
  Square,
  Upload,
  UserPlus,
} from 'lucide-react';
import {
  auth,
  clearCachedCalendarAccessToken,
  getCachedCalendarAccessToken,
  mapGoogleOAuthErrorToVietnamese,
  signInWithGooglePopup,
} from './lib/firebase';
import { HomePage } from './components/HomePage';
import { GeminiChatbotDrawer } from './components/GeminiChatbotDrawer';
import { ResignationTab } from './components/ResignationTab';
import {
  STORAGE_KEYS,
  SECTORS,
  EMPLOYMENT_TYPES,
  INDUSTRIES,
  KAIGO_FACILITY_TYPES,
  PROGRESS_STATUSES,
  JLPT_LEVELS,
  INITIAL_CANDIDATES,
  INITIAL_SCHOOL_STUDENTS,
  INITIAL_KYUJINS,
  INITIAL_PT_PARTNERS,
  INITIAL_RESIGNATIONS,
} from './data/seedData';
import type {
  Candidate,
  CareFilterType,
  ChannelType,
  DailyFocusCompletionMap,
  DailyFocusItem,
  KyujinCompany,
  KyujinSlot,
  MainTabId,
  PtPartner,
  ReplyScenario,
  ResignationRecord,
  RirekishoDraft,
} from './types/crm';
import {
  buildAnonymousJobPostText,
  buildRirekishoFallbackFromText,
  buildRirekishoStandaloneHtml,
  buildSmartReplyMessages,
  calcJapaneseAgeFromBirthDate,
  copyTextSafely,
  createInitialRirekishoDraft,
  detectEmploymentTypeFromText,
  detectIndustryFromText,
  detectSectorFromIndustryOrText,
  doesAreaMatchFilter,
  evaluateCandidateKyujinMatches,
  extractCandidateViaServer,
  extractDeclineReason,
  extractKyujinRawTextViaServer,
  extractKyujinViaServer,
  extractNormalizedMendanInfo,
  extractRirekishoTextFromLocalFile,
  extractRirekishoViaServer,
  formatPublicArea,
  getEffectiveProposedCompany,
  getMissingInfoForCandidate,
  getUrgentFollowUpReason,
  isActualProposedRecord,
  normalizeBonusInfo,
  normalizeHousingDisplay,
  normalizeKyujinItem,
  parseChatIntoUpdates,
  parseExcelFileToKyujinEntries,
  parseFlexibleDateToTimestamp,
  parseRawKyujinTextFallback,
  sanitizeJobDescriptionForCandidate,
} from './utils/crmHelpers';
import {
  buildCandidateCalendarEventDraft,
  buildResignationCalendarEventDraft,
  getSafeFacebookMessengerUrl,
  isValidGoogleCalendarUrl,
  normalizeFacebookPageTargetUrl,
  type CalendarEventDraft,
} from './utils/calendarHelpers';
import {
  compressImageFileToBase64,
  extractFirstImageFromClipboard,
} from './utils/imageCompression';
import {
  applyManualProtectionOnResignationSave,
  assessResignationDeadline,
  buildResignationReminderMessages,
  convertResignationToCandidate,
  getJstTodayYmd,
  mergeResignationWithGasProtection,
  normalizeResignationNationality,
} from './utils/resignationHelpers';

const MESSENGER_WEBHOOK_URL = 'https://fb-crm.ai.studio/api/webhooks/messenger';
const FB_PAGE_LINK_STORAGE_KEY = 'tokutei_fb_page_link_v1';

function readFallbackCandidates(): Candidate[] {
  try {
    const saved = localStorage.getItem(STORAGE_KEYS.CANDIDATES);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length > 0) return parsed;
    }
  } catch {
    // ignore storage parse errors
  }
  return INITIAL_CANDIDATES;
}

function readFallbackSchoolStudents(): Candidate[] {
  try {
    const saved = localStorage.getItem(STORAGE_KEYS.SCHOOL_STUDENTS);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed)) return parsed;
    }
  } catch {
    // ignore storage parse errors
  }
  return INITIAL_SCHOOL_STUDENTS;
}

function readFallbackKyujins(): KyujinCompany[] {
  try {
    const saved = localStorage.getItem(STORAGE_KEYS.KYUJINS);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed.map(normalizeKyujinItem);
      }
    }
  } catch {
    // ignore storage parse errors
  }
  return INITIAL_KYUJINS.map(normalizeKyujinItem);
}

function readFallbackPtPartners(): PtPartner[] {
  try {
    const saved = localStorage.getItem(STORAGE_KEYS.PT_PARTNERS);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed)) return parsed;
    }
  } catch {
    // ignore storage parse errors
  }
  return INITIAL_PT_PARTNERS;
}

function readFallbackResignations(): ResignationRecord[] {
  try {
    const saved = localStorage.getItem(STORAGE_KEYS.RESIGNATIONS);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed)) {
        return parsed.map((item) => ({
          ...item,
          nationality: normalizeResignationNationality(item?.nationality),
        }));
      }
    }
  } catch {
    // ignore storage parse errors
  }
  return INITIAL_RESIGNATIONS;
}

function readFallbackDailyFocus(): DailyFocusCompletionMap {
  try {
    const saved = localStorage.getItem(STORAGE_KEYS.DAILY_FOCUS);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (parsed && typeof parsed === 'object') return parsed;
    }
  } catch {
    // ignore storage parse errors
  }
  return {};
}

export default function App() {
  const [candidates, setCandidates] = useState<Candidate[]>(() => readFallbackCandidates());
  const [schoolStudents, setSchoolStudents] = useState<Candidate[]>(() =>
    readFallbackSchoolStudents()
  );
  const [kyujins, setKyujins] = useState<KyujinCompany[]>(() => readFallbackKyujins());
  const [ptPartners, setPtPartners] = useState<PtPartner[]>(() => readFallbackPtPartners());
  const [resignations, setResignations] = useState<ResignationRecord[]>(() =>
    readFallbackResignations()
  );
  const [dailyFocusCompleted, setDailyFocusCompleted] = useState<DailyFocusCompletionMap>(() =>
    readFallbackDailyFocus()
  );

  const [gasWriteLocked, setGasWriteLocked] = useState(true);
  const [gasWebAppUrl, setGasWebAppUrl] = useState(() => {
    try {
      return localStorage.getItem(STORAGE_KEYS.GAS_URL) || '';
    } catch {
      return '';
    }
  });
  const [isSyncingGas, setIsSyncingGas] = useState(false);

  // Navigation & Filters
  const [activeTab, setActiveTab] = useState<MainTabId>('all');
  const [searchQuery, setSearchQuery] = useState('');
  const [careFilter, setCareFilter] = useState<CareFilterType>('all');
  const [statusFilter, setStatusFilter] = useState('all');
  const [proposedMonthFilter, setProposedMonthFilter] = useState('all');
  const [proposedStageFilter, setProposedStageFilter] = useState('all');
  const [declineReasonQuery, setDeclineReasonQuery] = useState('');
  const [kyujinSearch, setKyujinSearch] = useState('');
  const [kyujinStatusFilter, setKyujinStatusFilter] = useState<'all' | 'open' | 'closed'>('all');
  const [kyujinBonusOnly, setKyujinBonusOnly] = useState(false);
  const [kyujinSectorFilter, setKyujinSectorFilter] = useState('all');
  const [kyujinIndustryFilter, setKyujinIndustryFilter] = useState('all');
  const [kyujinEmploymentTypeFilter, setKyujinEmploymentTypeFilter] = useState('all');
  const [kyujinFacilityFilter, setKyujinFacilityFilter] = useState('all');
  const [selectedSlotIds, setSelectedSlotIds] = useState<string[]>([]);
  const [isImportingKyujin, setIsImportingKyujin] = useState(false);
  const [kyujinImportProgress, setKyujinImportProgress] = useState<{
    current: number;
    total: number;
    label: string;
  } | null>(null);
  const [kyujinRawModalOpen, setKyujinRawModalOpen] = useState(false);
  const [kyujinRawText, setKyujinRawText] = useState('');

  // Modals & AI Extraction States
  const [editingPerson, setEditingPerson] = useState<{
    data: Candidate;
    isSchool: boolean;
    isNew: boolean;
  } | null>(null);
  const [editChatSnippet, setEditChatSnippet] = useState('');
  const [isExtractingCandidateAi, setIsExtractingCandidateAi] = useState(false);
  const [fbChatModalOpen, setFbChatModalOpen] = useState(false);
  const [fbChatMode, setFbChatMode] = useState<'new' | 'update'>('new');
  const [fbChatChannel, setFbChatChannel] = useState<ChannelType>('personal');
  const [fbChatTargetId, setFbChatTargetId] = useState('');
  const [fbChatRawText, setFbChatRawText] = useState('');
  const [isExtractingFbChatAi, setIsExtractingFbChatAi] = useState(false);
  const [resignationReminderModal, setResignationReminderModal] =
    useState<ResignationRecord | null>(null);
  const [syncingResignationCalIds, setSyncingResignationCalIds] = useState<
    Record<string, boolean>
  >({});
  const [isBulkSyncingResignationsCal, setIsBulkSyncingResignationsCal] = useState(false);

  const [rirekishoModal, setRirekishoModal] = useState<{
    person: Candidate;
    draft: RirekishoDraft;
  } | null>(null);
  const [isParsingRirekisho, setIsParsingRirekisho] = useState(false);
  const [rirekishoStatusBanner, setRirekishoStatusBanner] = useState<{
    type: 'loading' | 'success' | 'error';
    text: string;
  } | null>(null);
  const [rirekishoChatInput, setRirekishoChatInput] = useState('');

  const [replyModal, setReplyModal] = useState<{
    person: Candidate;
    scenario: ReplyScenario;
  } | null>(null);
  const [previewModal, setPreviewModal] = useState<{
    company: KyujinCompany;
    slot: KyujinSlot;
  } | null>(null);
  const [confirmDialog, setConfirmDialog] = useState<{
    title: string;
    message: string;
    confirmLabel?: string;
    cancelLabel?: string;
    tone?: 'danger' | 'accent';
    onConfirm: () => void;
    onCancel?: () => void;
  } | null>(null);
  const [ptModal, setPtModal] = useState<{
    isNew: boolean;
    data: PtPartner;
  } | null>(null);

  // App View Mode, Auth, Webhook Status & Chatbot State
  const [viewMode, setViewMode] = useState<'home' | 'workspace'>('home');
  const [chatbotOpen, setChatbotOpen] = useState(false);
  const [authUser, setAuthUser] = useState<User | null>(null);
  const [authLoading, setAuthLoading] = useState(false);
  const [authStatus, setAuthStatus] = useState<
    'signed_out' | 'checking' | 'signed_in'
  >('signed_out');
  const [webhookStatus, setWebhookStatus] = useState<{
    configured: boolean;
    hasVerifyToken: boolean;
    hasAppSecret: boolean;
    verifyTokenEnforced?: boolean;
    signatureEnforced?: boolean;
    webhookUrl: string;
    sharedWebhookUrl?: string;
  } | null>(null);
  const [messengerLinkModalOpen, setMessengerLinkModalOpen] = useState(false);
  const [fbPageLink, setFbPageLink] = useState(() => {
    try {
      return localStorage.getItem(FB_PAGE_LINK_STORAGE_KEY) || '';
    } catch {
      return '';
    }
  });
  const [fbPageLinkDraft, setFbPageLinkDraft] = useState('');
  const [isTestingWebhook, setIsTestingWebhook] = useState(false);
  const [isRefreshingSql, setIsRefreshingSql] = useState(false);
  const [hasCalendarToken, setHasCalendarToken] = useState(() =>
    Boolean(getCachedCalendarAccessToken())
  );
  const [syncingCalendarId, setSyncingCalendarId] = useState<string | null>(null);
  const [isBulkSyncingCalendar, setIsBulkSyncingCalendar] = useState(false);

  const sqlHydratedForUidRef = useRef<string | null>(null);
  const sqlSyncTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const skipNextSqlSaveRef = useRef(false);

  // Toast Notification
  const [toastMsg, setToastMsg] = useState('');
  const toastTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const backupInputRef = useRef<HTMLInputElement | null>(null);
  const kyujinFileInputRef = useRef<HTMLInputElement | null>(null);
  const rirekishoFileInputRef = useRef<HTMLInputElement | null>(null);

  const showToast = useCallback((msg: string) => {
    setToastMsg(msg);
    if (toastTimerRef.current) clearTimeout(toastTimerRef.current);
    toastTimerRef.current = setTimeout(() => setToastMsg(''), 3400);
  }, []);

  const getFirebaseIdToken = useCallback(async () => {
    if (!auth.currentUser) return null;
    try {
      return await auth.currentUser.getIdToken();
    } catch {
      return null;
    }
  }, []);

  const ensureGoogleCalendarToken = useCallback(
    async (forceReauth = false): Promise<string | null> => {
      if (!forceReauth) {
        const existing = getCachedCalendarAccessToken();
        if (existing) {
          setHasCalendarToken(true);
          return existing;
        }
      } else {
        clearCachedCalendarAccessToken();
        setHasCalendarToken(false);
      }

      try {
        const { calendarAccessToken } = await signInWithGooglePopup(forceReauth);
        setHasCalendarToken(Boolean(calendarAccessToken));
        if (!calendarAccessToken) {
          showToast(
            '⚠️ Bạn chưa cấp quyền Google Calendar (calendar.events). Vui lòng tích chọn cấp quyền Calendar khi đăng nhập Google.'
          );
          return null;
        }
        return calendarAccessToken;
      } catch (err) {
        showToast(`⚠️ ${mapGoogleOAuthErrorToVietnamese(err)}`);
        return null;
      }
    },
    [showToast]
  );

  const applyCalendarEventResultToState = useCallback(
    (
      entityId: string,
      eventId: string | undefined,
      htmlLink: string | undefined,
      entityType: 'candidate' | 'resignation' = 'candidate'
    ) => {
      if (entityType === 'resignation') {
        setResignations((prev) =>
          prev.map((item) =>
            item.id === entityId
              ? {
                  ...item,
                  googleCalendarEventId: eventId,
                  googleCalendarEventLink: htmlLink,
                }
              : item
          )
        );
        return;
      }
      const updateList = (list: Candidate[]) =>
        list.map((item) =>
          item.id === entityId
            ? {
                ...item,
                googleCalendarEventId: eventId,
                googleCalendarEventLink: htmlLink,
              }
            : item
        );
      setCandidates(updateList);
      setSchoolStudents(updateList);
    },
    []
  );

  const executeCalendarSyncForDraft = useCallback(
    async (draft: CalendarEventDraft) => {
      const entityType = draft.entityType || 'candidate';
      if (entityType === 'resignation') {
        setSyncingResignationCalIds((prev) => ({ ...prev, [draft.candidateId]: true }));
      } else {
        setSyncingCalendarId(draft.candidateId);
      }
      try {
        // Request OAuth token FIRST inside user gesture to avoid popup-blocker issues
        let googleToken = await ensureGoogleCalendarToken(false);
        if (!googleToken) return;

        let firebaseToken = await getFirebaseIdToken();
        if (!firebaseToken) {
          googleToken = await ensureGoogleCalendarToken(true);
          if (!googleToken) return;
          firebaseToken = await getFirebaseIdToken();
          if (!firebaseToken) {
            showToast('⚠️ Vui lòng đăng nhập Google để đồng bộ Google Calendar');
            return;
          }
        }

        const callSyncApi = async (gToken: string, fbToken: string) =>
          fetch('/api/calendar/sync-candidate', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              Authorization: `Bearer ${fbToken}`,
              'X-Google-Access-Token': gToken,
            },
            body: JSON.stringify({
              candidateId: draft.candidateId,
              entityType,
              eventTitle: draft.eventTitle,
              description: draft.description,
              dateIso: draft.dateIso,
              startIso: draft.startIso,
              endIso: draft.endIso,
              existingEventId: draft.existingEventId,
              reminderMinutes: draft.reminderMinutes,
            }),
          });

        let resp = await callSyncApi(googleToken, firebaseToken);
        if (resp.status === 401) {
          clearCachedCalendarAccessToken();
          setHasCalendarToken(false);
          const refreshedGoogle = await ensureGoogleCalendarToken(true);
          if (!refreshedGoogle) return;
          const refreshedFb = (await getFirebaseIdToken()) || firebaseToken;
          resp = await callSyncApi(refreshedGoogle, refreshedFb);
        }

        const data = await resp.json().catch(() => ({}));
        if (!resp.ok) {
          throw new Error(data?.error || `HTTP ${resp.status}`);
        }

        applyCalendarEventResultToState(
          draft.candidateId,
          data.eventId,
          data.htmlLink,
          entityType
        );
        showToast(
          data.action === 'updated'
            ? `✅ Đã cập nhật sự kiện Google Calendar cho ${draft.candidateName}`
            : `✅ Đã tạo sự kiện Google Calendar cho ${draft.candidateName}`
        );
      } catch (err) {
        const msg = err instanceof Error ? err.message : 'Đồng bộ thất bại';
        showToast(`❌ Lỗi Google Calendar: ${msg}`);
      } finally {
        if (entityType === 'resignation') {
          setSyncingResignationCalIds((prev) => ({ ...prev, [draft.candidateId]: false }));
        } else {
          setSyncingCalendarId(null);
        }
      }
    },
    [
      applyCalendarEventResultToState,
      ensureGoogleCalendarToken,
      getFirebaseIdToken,
      showToast,
    ]
  );

  const executeCalendarDelete = useCallback(
    async (
      entityId: string,
      eventId: string,
      alsoUpdateState = true,
      entityType: 'candidate' | 'resignation' = 'candidate'
    ) => {
      if (entityType === 'resignation') {
        setSyncingResignationCalIds((prev) => ({ ...prev, [entityId]: true }));
      } else {
        setSyncingCalendarId(entityId);
      }
      try {
        const firebaseToken = await getFirebaseIdToken();
        if (!firebaseToken) return;
        let googleToken = getCachedCalendarAccessToken();
        if (!googleToken) {
          googleToken = await ensureGoogleCalendarToken();
        }
        if (!googleToken) return;

        const callDelApi = (gToken: string) =>
          fetch(
            `/api/calendar/events/${encodeURIComponent(
              eventId
            )}?candidateId=${encodeURIComponent(entityId)}&entityType=${encodeURIComponent(
              entityType
            )}`,
            {
              method: 'DELETE',
              headers: {
                Authorization: `Bearer ${firebaseToken}`,
                'X-Google-Access-Token': gToken,
              },
            }
          );

        let resp = await callDelApi(googleToken);
        if (resp.status === 401) {
          clearCachedCalendarAccessToken();
          setHasCalendarToken(false);
          const refreshed = await ensureGoogleCalendarToken();
          if (!refreshed) return;
          resp = await callDelApi(refreshed);
        }

        const data = await resp.json().catch(() => ({}));
        if (!resp.ok) {
          throw new Error(data?.error || `HTTP ${resp.status}`);
        }

        if (alsoUpdateState) {
          applyCalendarEventResultToState(entityId, undefined, undefined, entityType);
        }
        showToast('🗑️ Đã xóa sự kiện tương ứng trên Google Calendar');
      } catch (err) {
        const msg = err instanceof Error ? err.message : 'Xóa lịch thất bại';
        showToast(`⚠️ Không thể xóa sự kiện Calendar: ${msg}`);
      } finally {
        if (entityType === 'resignation') {
          setSyncingResignationCalIds((prev) => ({ ...prev, [entityId]: false }));
        } else {
          setSyncingCalendarId(null);
        }
      }
    },
    [
      applyCalendarEventResultToState,
      ensureGoogleCalendarToken,
      getFirebaseIdToken,
      showToast,
    ]
  );

  const requestSyncCandidateCalendar = useCallback(
    (person: Candidate) => {
      const draft = buildCandidateCalendarEventDraft(person);
      if (!draft) {
        showToast(
          '⚠️ Ứng viên chưa có ngày giờ Mendan hoặc Phỏng vấn hợp lệ (VD: 2026-11-25 14:00)'
        );
        return;
      }

      const isUpdate = Boolean(draft.existingEventId);
      setConfirmDialog({
        title: isUpdate
          ? 'Cập nhật sự kiện Google Calendar'
          : 'Đồng bộ lên Google Calendar',
        message: `${
          isUpdate
            ? 'Cập nhật lại sự kiện hiện có'
            : 'Tạo mới sự kiện trên Google Calendar của bạn'
        }:\n• Tiêu đề: ${draft.eventTitle}\n• Thời gian: ${
          draft.displayDateTime
        }\n• Nhắc nhở tự động: trước 30 phút & 10 phút.`,
        confirmLabel: isUpdate ? 'Cập nhật Calendar' : 'Đồng bộ ngay',
        cancelLabel: 'Hủy',
        tone: 'accent',
        onConfirm: () => {
          setConfirmDialog(null);
          void executeCalendarSyncForDraft(draft);
        },
      });
    },
    [executeCalendarSyncForDraft, showToast]
  );

  const requestBulkSyncUpcomingCalendar = useCallback(() => {
    const allPeople = [...candidates, ...schoolStudents];
    const drafts = allPeople
      .map((p) => buildCandidateCalendarEventDraft(p))
      .filter((d): d is CalendarEventDraft => Boolean(d));

    if (drafts.length === 0) {
      showToast('⚠️ Không có ứng viên nào có lịch Mendan / Phỏng vấn hợp lệ');
      return;
    }

    const previewNames = drafts
      .slice(0, 6)
      .map((d) => `• ${d.eventTitle} (${d.displayDateTime})`)
      .join('\n');
    const extraCount = drafts.length > 6 ? `\n...và ${drafts.length - 6} lịch khác.` : '';

    setConfirmDialog({
      title: `Đồng bộ ${drafts.length} lịch hẹn lên Google Calendar`,
      message: `Xác nhận tạo mới hoặc cập nhật ${drafts.length} sự kiện trên Google Calendar chính của bạn (nhắc nhở trước 30 phút & 10 phút):\n${previewNames}${extraCount}`,
      confirmLabel: `Đồng bộ ${drafts.length} lịch`,
      cancelLabel: 'Hủy',
      tone: 'accent',
      onConfirm: async () => {
        setConfirmDialog(null);
        setIsBulkSyncingCalendar(true);
        try {
          let googleToken = await ensureGoogleCalendarToken(false);
          if (!googleToken) return;

          let firebaseToken = await getFirebaseIdToken();
          if (!firebaseToken) {
            googleToken = await ensureGoogleCalendarToken(true);
            if (!googleToken) return;
            firebaseToken = await getFirebaseIdToken();
            if (!firebaseToken) {
              showToast('⚠️ Vui lòng đăng nhập Google để đồng bộ Google Calendar');
              return;
            }
          }

          const callBulkApi = (gToken: string, fbToken: string) =>
            fetch('/api/calendar/sync-bulk', {
              method: 'POST',
              headers: {
                'Content-Type': 'application/json',
                Authorization: `Bearer ${fbToken}`,
                'X-Google-Access-Token': gToken,
              },
              body: JSON.stringify({ items: drafts }),
            });

          let resp = await callBulkApi(googleToken, firebaseToken);
          if (resp.status === 401) {
            clearCachedCalendarAccessToken();
            setHasCalendarToken(false);
            const refreshedGoogle = await ensureGoogleCalendarToken(true);
            if (!refreshedGoogle) return;
            const refreshedFb = (await getFirebaseIdToken()) || firebaseToken;
            resp = await callBulkApi(refreshedGoogle, refreshedFb);
          }

          const data = (await resp.json().catch(() => ({}))) as {
            synced?: Array<{
              candidateId: string;
              eventId: string;
              htmlLink: string;
            }>;
            failed?: Array<{ candidateId: string; error: string }>;
            error?: string;
          };

          if (!resp.ok) {
            throw new Error(data.error || `HTTP ${resp.status}`);
          }

          const syncedList = Array.isArray(data.synced) ? data.synced : [];
          const failedList = Array.isArray(data.failed) ? data.failed : [];
          const syncedMap = new Map(syncedList.map((s) => [s.candidateId, s]));

          if (syncedList.length > 0) {
            const applyBulk = (list: Candidate[]) =>
              list.map((c) => {
                const hit = syncedMap.get(c.id);
                return hit
                  ? {
                      ...c,
                      googleCalendarEventId: hit.eventId,
                      googleCalendarEventLink: hit.htmlLink,
                    }
                  : c;
              });
            setCandidates(applyBulk);
            setSchoolStudents(applyBulk);
          }

          if (syncedList.length === 0 && failedList.length > 0) {
            showToast(`❌ Lỗi đồng bộ hàng loạt: ${failedList[0].error}`);
          } else if (failedList.length > 0) {
            showToast(
              `⚠️ Đã đồng bộ ${syncedList.length} lịch (${failedList.length} lịch lỗi: ${failedList[0].error})`
            );
          } else {
            showToast(
              `✅ Đã đồng bộ toàn bộ ${syncedList.length} lịch hẹn lên Google Calendar`
            );
          }
        } catch (err) {
          const msg = err instanceof Error ? err.message : 'Đồng bộ thất bại';
          showToast(`❌ Lỗi đồng bộ hàng loạt: ${msg}`);
        } finally {
          setIsBulkSyncingCalendar(false);
        }
      },
    });
  }, [
    candidates,
    schoolStudents,
    ensureGoogleCalendarToken,
    getFirebaseIdToken,
    showToast,
  ]);

  const liveMessengerWebhookUrl = useMemo(() => {
    if (
      webhookStatus?.webhookUrl &&
      webhookStatus.webhookUrl !== MESSENGER_WEBHOOK_URL
    ) {
      return webhookStatus.webhookUrl;
    }
    if (typeof window !== 'undefined' && window.location?.origin) {
      return `${window.location.origin}/api/webhooks/messenger`;
    }
    return webhookStatus?.webhookUrl || MESSENGER_WEBHOOK_URL;
  }, [webhookStatus]);

  const handleCopyWebhookUrl = useCallback(() => {
    copyTextSafely(liveMessengerWebhookUrl);
    showToast(`Đã sao chép Webhook URL: ${liveMessengerWebhookUrl}`);
  }, [liveMessengerWebhookUrl, showToast]);

  const openMessengerLinkModal = useCallback(() => {
    setFbPageLinkDraft(fbPageLink);
    setMessengerLinkModalOpen(true);
  }, [fbPageLink]);

  const handleSaveFbPageLink = useCallback(() => {
    const trimmed = fbPageLinkDraft.trim();
    if (!trimmed) {
      setFbPageLink('');
      try {
        localStorage.removeItem(FB_PAGE_LINK_STORAGE_KEY);
      } catch {
        // ignore storage error
      }
      showToast('Đã xóa liên kết Fanpage Messenger');
      return;
    }
    const normalized = normalizeFacebookPageTargetUrl(trimmed);
    if (!normalized) {
      showToast(
        '⚠️ Vui lòng nhập username Fanpage (VD: mypage) hoặc link https://m.me/... / https://facebook.com/...'
      );
      return;
    }
    setFbPageLink(normalized);
    setFbPageLinkDraft(normalized);
    try {
      localStorage.setItem(FB_PAGE_LINK_STORAGE_KEY, normalized);
    } catch {
      // ignore storage error
    }
    showToast('✅ Đã lưu liên kết Facebook Messenger Fanpage');
  }, [fbPageLinkDraft, showToast]);

  const sharedMessengerWebhookUrl = useMemo(() => {
    if (webhookStatus?.sharedWebhookUrl) return webhookStatus.sharedWebhookUrl;
    if (liveMessengerWebhookUrl.includes('://ais-dev-')) {
      return liveMessengerWebhookUrl.replace('://ais-dev-', '://ais-pre-');
    }
    return '';
  }, [webhookStatus, liveMessengerWebhookUrl]);

  const handleSyncResignationCalendarSingle = useCallback(
    async (record: ResignationRecord) => {
      if (record.paperworkStatus === 'Hoàn tất') {
        showToast('ℹ️ Hồ sơ này đã hoàn tất giấy tờ nên không cần đồng bộ lịch nhắc.');
        return;
      }
      const draft = buildResignationCalendarEventDraft(record);
      if (!draft) {
        showToast(
          '⚠️ Vui lòng cập nhật Ngày nghỉ việc hoặc Hạn giấy tờ trước khi đồng bộ lịch.'
        );
        return;
      }

      const isUpdate = Boolean(draft.existingEventId);
      setConfirmDialog({
        title: isUpdate
          ? 'Cập nhật lịch nhắc giấy tờ trên Google Calendar'
          : 'Đồng bộ lịch nhắc giấy tờ lên Google Calendar',
        message: `${
          isUpdate
            ? 'Cập nhật lại sự kiện hiện có (tránh tạo trùng)'
            : 'Tạo mới sự kiện nhắc hạn trên Google Calendar'
        }:\n• Tiêu đề: ${draft.eventTitle}\n• Thời gian (Asia/Tokyo): ${
          draft.displayDateTime
        }\n• Nhắc nhở tự động: 3 ngày, 1 ngày và 2 giờ trước hạn.`,
        confirmLabel: isUpdate ? 'Cập nhật Calendar' : 'Đồng bộ ngay',
        cancelLabel: 'Hủy',
        tone: 'accent',
        onConfirm: () => {
          setConfirmDialog(null);
          void executeCalendarSyncForDraft(draft);
        },
      });
    },
    [executeCalendarSyncForDraft, showToast]
  );

  const handleBulkSyncResignationsCalendar = useCallback(async () => {
    const activeRecords = resignations.filter((r) => r.paperworkStatus !== 'Hoàn tất');
    if (activeRecords.length === 0) {
      showToast('ℹ️ Không có hồ sơ nghỉ việc nào đang chờ hạn nộp giấy tờ.');
      return;
    }

    const drafts = activeRecords
      .map((r) => buildResignationCalendarEventDraft(r))
      .filter((d): d is CalendarEventDraft => Boolean(d));

    if (drafts.length === 0) {
      showToast(
        '⚠️ Vui lòng cập nhật Ngày nghỉ việc hoặc Hạn giấy tờ trước khi đồng bộ lịch.'
      );
      return;
    }

    const skippedCount = activeRecords.length - drafts.length;
    const previewNames = drafts
      .slice(0, 6)
      .map((d) => `• ${d.eventTitle} (${d.displayDateTime})`)
      .join('\n');
    const extraCount =
      drafts.length > 6 ? `\n...và ${drafts.length - 6} hồ sơ khác.` : '';
    const skippedNote =
      skippedCount > 0
        ? `\n(Bỏ qua ${skippedCount} hồ sơ chưa có Ngày nghỉ việc và Hạn giấy tờ)`
        : '';

    setConfirmDialog({
      title: `Đồng bộ ${drafts.length} lịch nghỉ việc lên Google Calendar`,
      message: `Xác nhận tạo mới hoặc cập nhật ${drafts.length} sự kiện (múi giờ Asia/Tokyo, nhắc trước 3 ngày / 1 ngày / 2 giờ):\n${previewNames}${extraCount}${skippedNote}`,
      confirmLabel: `Đồng bộ ${drafts.length} lịch`,
      cancelLabel: 'Hủy',
      tone: 'accent',
      onConfirm: async () => {
        setConfirmDialog(null);
        setIsBulkSyncingResignationsCal(true);
        try {
          let googleToken = await ensureGoogleCalendarToken(false);
          if (!googleToken) return;

          let firebaseToken = await getFirebaseIdToken();
          if (!firebaseToken) {
            googleToken = await ensureGoogleCalendarToken(true);
            if (!googleToken) return;
            firebaseToken = await getFirebaseIdToken();
            if (!firebaseToken) {
              showToast('⚠️ Vui lòng đăng nhập Google để đồng bộ Google Calendar');
              return;
            }
          }

          const callBulk = (gToken: string, fbToken: string) =>
            fetch('/api/calendar/sync-bulk', {
              method: 'POST',
              headers: {
                'Content-Type': 'application/json',
                Authorization: `Bearer ${fbToken}`,
                'X-Google-Access-Token': gToken,
              },
              body: JSON.stringify({ items: drafts }),
            });

          let resp = await callBulk(googleToken, firebaseToken);
          if (resp.status === 401) {
            clearCachedCalendarAccessToken();
            setHasCalendarToken(false);
            const refreshedGoogle = await ensureGoogleCalendarToken(true);
            if (!refreshedGoogle) return;
            const refreshedFb = (await getFirebaseIdToken()) || firebaseToken;
            resp = await callBulk(refreshedGoogle, refreshedFb);
          }

          const data = (await resp.json().catch(() => ({}))) as {
            synced?: Array<{ candidateId: string; eventId: string; htmlLink: string }>;
            failed?: Array<{ candidateId: string; error: string }>;
            error?: string;
          };
          if (!resp.ok) throw new Error(data.error || `HTTP ${resp.status}`);

          const syncedList = Array.isArray(data.synced) ? data.synced : [];
          const failedList = Array.isArray(data.failed) ? data.failed : [];
          const syncedMap = new Map(syncedList.map((s) => [s.candidateId, s]));

          if (syncedList.length > 0) {
            setResignations((prev) =>
              prev.map((r) => {
                const hit = syncedMap.get(r.id);
                return hit
                  ? {
                      ...r,
                      googleCalendarEventId: hit.eventId,
                      googleCalendarEventLink: hit.htmlLink,
                    }
                  : r;
              })
            );
          }

          if (syncedList.length === 0 && failedList.length > 0) {
            showToast(`❌ Lỗi đồng bộ Calendar nghỉ việc: ${failedList[0].error}`);
          } else if (failedList.length > 0) {
            showToast(
              `⚠️ Đã đồng bộ ${syncedList.length} hồ sơ (${failedList.length} hồ sơ lỗi: ${failedList[0].error})`
            );
          } else {
            showToast(
              `✅ Đã đồng bộ ${syncedList.length} hạn giấy tờ nghỉ việc lên Google Calendar (nhắc 3 ngày / 1 ngày / 2 giờ)`
            );
          }
        } catch (err) {
          const msg = err instanceof Error ? err.message : 'Đồng bộ thất bại';
          showToast(`❌ Lỗi đồng bộ Calendar nghỉ việc: ${msg}`);
        } finally {
          setIsBulkSyncingResignationsCal(false);
        }
      },
    });
  }, [resignations, ensureGoogleCalendarToken, getFirebaseIdToken, showToast]);

  const handleSaveResignation = useCallback(
    async (record: ResignationRecord, isNew: boolean) => {
      let protectedRec: ResignationRecord = record;
      setResignations((prev) => {
        const existing = prev.find((r) => r.id === record.id) || null;
        protectedRec = applyManualProtectionOnResignationSave(existing, {
          ...record,
          updatedAt: new Date().toISOString(),
        });
        return isNew
          ? [protectedRec, ...prev]
          : prev.map((r) => (r.id === protectedRec.id ? protectedRec : r));
      });
      showToast(
        isNew
          ? '✅ Đã thêm hồ sơ nghỉ việc mới'
          : `✅ Đã cập nhật hồ sơ (đã bật bảo vệ chống ghi đè GAS): ${protectedRec.fullName}`
      );

      if (protectedRec.paperworkStatus === 'Hoàn tất' && protectedRec.googleCalendarEventId) {
        void executeCalendarDelete(
          protectedRec.id,
          protectedRec.googleCalendarEventId,
          true,
          'resignation'
        );
      } else if (
        (protectedRec.paperworkDeadline || protectedRec.resignationDate) &&
        protectedRec.paperworkStatus !== 'Hoàn tất' &&
        auth.currentUser &&
        getCachedCalendarAccessToken()
      ) {
        const draft = buildResignationCalendarEventDraft(protectedRec);
        if (draft) {
          void executeCalendarSyncForDraft(draft);
        }
      }
    },
    [executeCalendarDelete, executeCalendarSyncForDraft, showToast]
  );

  const handleResignationStatusChange = useCallback(
    async (record: ResignationRecord, nextStatus: string) => {
      const updated: ResignationRecord = {
        ...record,
        paperworkStatus: nextStatus as ResignationRecord['paperworkStatus'],
        updatedAt: new Date().toISOString(),
      };
      setResignations((prev) => prev.map((r) => (r.id === record.id ? updated : r)));
      showToast(`✅ Đã chuyển trạng thái giấy tờ sang "${nextStatus}"`);

      if (nextStatus === 'Hoàn tất' && record.googleCalendarEventId) {
        void executeCalendarDelete(record.id, record.googleCalendarEventId, true, 'resignation');
      }
    },
    [executeCalendarDelete, showToast]
  );

  const handleDeleteResignation = useCallback(
    async (record: ResignationRecord) => {
      setResignations((prev) => prev.filter((r) => r.id !== record.id));
      showToast(`🗑️ Đã xóa hồ sơ nghỉ việc: ${record.fullName}`);
      if (record.googleCalendarEventId && auth.currentUser) {
        void executeCalendarDelete(record.id, record.googleCalendarEventId, false, 'resignation');
      }
    },
    [executeCalendarDelete, showToast]
  );

  const handleBulkImportResignations = useCallback(
    (rows: ResignationRecord[]) => {
      if (!rows.length) return;
      setResignations((prev) => {
        const nextList = [...prev];
        const newEntries: ResignationRecord[] = [];
        rows.forEach((rawRow) => {
          const rowWithNat: ResignationRecord = {
            ...rawRow,
            nationality: normalizeResignationNationality(rawRow.nationality),
          };
          const matchIdx = nextList.findIndex(
            (item) =>
              item.id === rowWithNat.id ||
              (item.fullName.trim().toLowerCase() === rowWithNat.fullName.trim().toLowerCase() &&
                item.previousCompany.trim().toLowerCase() ===
                  rowWithNat.previousCompany.trim().toLowerCase())
          );
          if (matchIdx >= 0) {
            nextList[matchIdx] = mergeResignationWithGasProtection(
              nextList[matchIdx],
              rowWithNat as ResignationRecord & Record<string, unknown>,
              gasWriteLocked
            );
          } else {
            newEntries.push(rowWithNat);
          }
        });
        return [...newEntries, ...nextList];
      });
      showToast(`✅ Đã nhập ${rows.length} hồ sơ nghỉ việc từ Excel/CSV`);
    },
    [gasWriteLocked, showToast]
  );

  const handleConvertResignationToCandidate = useCallback(
    (record: ResignationRecord) => {
      const existing = candidates.find(
        (c) =>
          (record.convertedCandidateId && c.id === record.convertedCandidateId) ||
          (c.name.toLowerCase() === record.fullName.toLowerCase() &&
            c.memo?.includes(record.previousCompany))
      );
      if (existing) {
        setActiveTab('all');
        setSearchQuery(existing.name);
        showToast(`ℹ️ ${record.fullName} đã có trong danh sách Ứng viên tìm việc mới`);
        return;
      }
      const newCand = convertResignationToCandidate(record);
      setCandidates((prev) => [newCand, ...prev]);
      setResignations((prev) =>
        prev.map((r) => (r.id === record.id ? { ...r, convertedCandidateId: newCand.id } : r))
      );
      setActiveTab('all');
      setEditingPerson({ data: newCand, isSchool: false, isNew: false });
      showToast(`✅ Đã chuyển "${record.fullName}" sang Ứng viên tìm việc mới`);
    },
    [candidates, showToast]
  );

  const handleTestMessengerWebhook = useCallback(async () => {
    setIsTestingWebhook(true);
    try {
      const token = await getFirebaseIdToken();
      const headers: Record<string, string> = { 'Content-Type': 'application/json' };
      if (token) headers.Authorization = `Bearer ${token}`;

      const resp = await fetch('/api/webhooks/messenger/test', {
        method: 'POST',
        headers,
        body: JSON.stringify({}),
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok || !data?.candidate) {
        throw new Error(data?.error || `HTTP ${resp.status}`);
      }
      const testCand = data.candidate as Candidate;
      setCandidates((prev) => [
        testCand,
        ...prev.filter((c) => c.id !== testCand.id),
      ]);
      setActiveTab('company');
      setMessengerLinkModalOpen(false);
      showToast(
        `✅ Đã nhận thử nghiệm từ Messenger và lưu vào FB会社公式: ${testCand.name}`
      );
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Test thất bại';
      showToast(`❌ Lỗi thử nghiệm Messenger: ${msg}`);
    } finally {
      setIsTestingWebhook(false);
    }
  }, [getFirebaseIdToken, showToast]);

  const handleSignInGoogle = async () => {
    setAuthLoading(true);
    try {
      const { calendarAccessToken } = await signInWithGooglePopup();
      setHasCalendarToken(Boolean(calendarAccessToken));
      showToast('Đã kết nối tài khoản Google, đang đồng bộ Cloud SQL...');
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Đăng nhập thất bại';
      showToast(`Không thể đăng nhập Google: ${msg}`);
    } finally {
      setAuthLoading(false);
    }
  };

  const handleSignOut = async () => {
    try {
      await signOut(auth);
      clearCachedCalendarAccessToken();
      setHasCalendarToken(false);
      sqlHydratedForUidRef.current = null;
      setAuthStatus('signed_out');
      showToast('Đã đăng xuất Google (dữ liệu nội bộ vẫn được giữ nguyên)');
    } catch {
      showToast('Không thể đăng xuất lúc này');
    }
  };

  const fetchWebhookStatus = useCallback(async (token?: string | null) => {
    try {
      const headers: Record<string, string> = {};
      if (token) headers.Authorization = `Bearer ${token}`;
      const resp = await fetch('/api/webhooks/messenger/status', { headers });
      if (!resp.ok) return;
      const statusData = await resp.json();
      setWebhookStatus(statusData);
    } catch {
      // ignore transient webhook status error
    }
  }, []);

  useEffect(() => {
    void fetchWebhookStatus(null);
    const unsub = onAuthStateChanged(auth, (user) => {
      setAuthUser(user);
      if (!user) {
        sqlHydratedForUidRef.current = null;
        setAuthStatus('signed_out');
      }
    });
    return () => unsub();
  }, [fetchWebhookStatus]);

  const mergePendingWebhookCandidates = useCallback(
    (baseList: Candidate[], incomingList?: Candidate[]): Candidate[] => {
      if (!Array.isArray(incomingList) || incomingList.length === 0) {
        return baseList;
      }
      const next = [...baseList];
      for (const incoming of incomingList) {
        const idx = next.findIndex(
          (c) =>
            c.id === incoming.id ||
            (incoming.fbSenderId && c.fbSenderId === incoming.fbSenderId)
        );
        if (idx >= 0) {
          next[idx] = {
            ...next[idx],
            ...incoming,
            memo:
              incoming.memo && incoming.memo !== next[idx].memo
                ? incoming.memo
                : next[idx].memo,
          };
        } else {
          next.unshift(incoming);
        }
      }
      return next;
    },
    []
  );

  const refreshCrmFromSql = useCallback(
    async (silent = false) => {
      if (!auth.currentUser || authStatus !== 'signed_in') {
        if (!silent) {
          await fetchWebhookStatus(null);
          showToast('ℹ️ Đang dùng chế độ làm việc mở (localStorage). Đăng nhập Google để đồng bộ Cloud SQL.');
        }
        return;
      }
      if (!silent) setIsRefreshingSql(true);
      try {
        const token = await auth.currentUser.getIdToken();
        if (!token) return;
        const resp = await fetch('/api/crm/state', {
          headers: { Authorization: `Bearer ${token}` },
        });
        if (!resp.ok) return;

        const data = (await resp.json()) as {
          candidates?: Candidate[];
          schoolStudents?: Candidate[];
          kyujins?: KyujinCompany[];
          ptPartners?: PtPartner[];
          resignations?: ResignationRecord[];
          pendingWebhookCandidates?: Candidate[];
          storageMode?: string;
        };

        if (
          data.storageMode === 'cloudsql' &&
          Array.isArray(data.candidates)
        ) {
          setCandidates((prev) => {
            const merged = mergePendingWebhookCandidates(
              data.candidates!,
              data.pendingWebhookCandidates
            );
            if (JSON.stringify(prev) === JSON.stringify(merged)) return prev;
            skipNextSqlSaveRef.current = true;
            return merged;
          });
        } else if (
          Array.isArray(data.pendingWebhookCandidates) &&
          data.pendingWebhookCandidates.length > 0
        ) {
          setCandidates((prev) =>
            mergePendingWebhookCandidates(prev, data.pendingWebhookCandidates)
          );
        }

        if (
          data.storageMode === 'cloudsql' &&
          Array.isArray(data.schoolStudents)
        ) {
          setSchoolStudents((prev) => {
            const nextJson = JSON.stringify(data.schoolStudents);
            if (JSON.stringify(prev) === nextJson) return prev;
            skipNextSqlSaveRef.current = true;
            return data.schoolStudents!;
          });
        }

        if (
          data.storageMode === 'cloudsql' &&
          Array.isArray(data.resignations)
        ) {
          setResignations((prev) => {
            const prevMap = new Map(prev.map((item) => [item.id, item]));
            const merged = data.resignations!.map((incoming) => {
              const existing = prevMap.get(incoming.id);
              if (
                existing &&
                (existing.gasProtected ||
                  (existing.manualProtectedFields && existing.manualProtectedFields.length > 0))
              ) {
                const existingTs = existing.updatedAt ? Date.parse(existing.updatedAt) : 0;
                const incomingTs = incoming.updatedAt ? Date.parse(incoming.updatedAt) : 0;
                if (existingTs > incomingTs) {
                  return mergeResignationWithGasProtection(
                    existing,
                    incoming as ResignationRecord & Record<string, unknown>,
                    true
                  );
                }
              }
              return {
                ...incoming,
                nationality: normalizeResignationNationality(incoming.nationality),
              };
            });
            const nextJson = JSON.stringify(merged);
            if (JSON.stringify(prev) === nextJson) return prev;
            skipNextSqlSaveRef.current = true;
            return merged;
          });
        }

        if (!silent) {
          await fetchWebhookStatus(token);
          showToast('Đã đồng bộ dữ liệu mới nhất từ hệ thống & Messenger');
        }
      } catch {
        if (!silent) {
          showToast('Không thể làm mới dữ liệu lúc này');
        }
      } finally {
        if (!silent) setIsRefreshingSql(false);
      }
    },
    [authStatus, fetchWebhookStatus, mergePendingWebhookCandidates, showToast]
  );

  // Load CRM snapshot & webhook status upon Google sign-in
  useEffect(() => {
    if (!authUser) return;
    if (sqlHydratedForUidRef.current === authUser.uid) return;

    let cancelled = false;
    const syncInitialFromSql = async () => {
      setAuthStatus('checking');

      try {
        const token = await authUser.getIdToken();
        if (!token || cancelled) return;

        const resp = await fetch('/api/crm/state', {
          headers: { Authorization: `Bearer ${token}` },
        });

        if (cancelled) return;

        if (resp.status === 401) {
          setAuthStatus('signed_out');
          return;
        }

        if (!resp.ok) {
          throw new Error(`HTTP ${resp.status}`);
        }

        const data = (await resp.json()) as {
          candidates?: Candidate[];
          schoolStudents?: Candidate[];
          kyujins?: KyujinCompany[];
          ptPartners?: PtPartner[];
          resignations?: ResignationRecord[];
          pendingWebhookCandidates?: Candidate[];
          hasExistingData?: boolean;
        };

        if (cancelled) return;

        setAuthStatus('signed_in');
        void fetchWebhookStatus(token);

        if (data.hasExistingData) {
          skipNextSqlSaveRef.current = true;
          const loadedCandidates = Array.isArray(data.candidates)
            ? data.candidates
            : [];
          setCandidates(
            mergePendingWebhookCandidates(
              loadedCandidates,
              data.pendingWebhookCandidates
            )
          );
          setSchoolStudents(
            Array.isArray(data.schoolStudents) ? data.schoolStudents : []
          );
          setKyujins(
            Array.isArray(data.kyujins)
              ? data.kyujins.map(normalizeKyujinItem)
              : []
          );
          setPtPartners(Array.isArray(data.ptPartners) ? data.ptPartners : []);
          if (Array.isArray(data.resignations) && data.resignations.length > 0) {
            setResignations(
              data.resignations.map((r) => ({
                ...r,
                nationality: normalizeResignationNationality(r.nationality),
              }))
            );
          }
        } else {
          const seedCandidates = mergePendingWebhookCandidates(
            readFallbackCandidates(),
            data.pendingWebhookCandidates
          );
          const seedSchools = readFallbackSchoolStudents();
          const seedKyujins = readFallbackKyujins();
          const seedPt = readFallbackPtPartners();
          const seedResignations = readFallbackResignations();

          skipNextSqlSaveRef.current = true;
          setCandidates(seedCandidates);
          setSchoolStudents(seedSchools);
          setKyujins(seedKyujins);
          setPtPartners(seedPt);
          setResignations(seedResignations);

          await fetch('/api/crm/state', {
            method: 'PUT',
            headers: {
              'Content-Type': 'application/json',
              Authorization: `Bearer ${token}`,
            },
            body: JSON.stringify({
              candidates: seedCandidates,
              schoolStudents: seedSchools,
              kyujins: seedKyujins,
              ptPartners: seedPt,
              resignations: seedResignations,
            }),
          });
        }
        sqlHydratedForUidRef.current = authUser.uid;
      } catch {
        if (!cancelled) {
          setAuthStatus('signed_in');
          sqlHydratedForUidRef.current = authUser.uid;
        }
      }
    };

    syncInitialFromSql();
    return () => {
      cancelled = true;
    };
  }, [authUser, fetchWebhookStatus, mergePendingWebhookCandidates]);

  // Background poll every 20s for newly arrived Messenger Webhook candidates
  useEffect(() => {
    if (authStatus !== 'signed_in' || viewMode !== 'workspace') return;
    const timer = setInterval(() => {
      if (!sqlSyncTimerRef.current) {
        void refreshCrmFromSql(true);
      }
    }, 20000);
    return () => clearInterval(timer);
  }, [authStatus, viewMode, refreshCrmFromSql]);

  // Debounced auto-save to PostgreSQL when signed in
  useEffect(() => {
    if (
      authStatus !== 'signed_in' ||
      !authUser ||
      sqlHydratedForUidRef.current !== authUser.uid
    ) {
      return;
    }

    if (skipNextSqlSaveRef.current) {
      skipNextSqlSaveRef.current = false;
      return;
    }

    if (sqlSyncTimerRef.current) clearTimeout(sqlSyncTimerRef.current);
    sqlSyncTimerRef.current = setTimeout(async () => {
      sqlSyncTimerRef.current = null;
      try {
        const token = await authUser.getIdToken();
        if (!token) return;
        await fetch('/api/crm/state', {
          method: 'PUT',
          headers: {
            'Content-Type': 'application/json',
            Authorization: `Bearer ${token}`,
          },
          body: JSON.stringify({
            candidates,
            schoolStudents,
            kyujins,
            ptPartners,
            resignations,
          }),
        });
      } catch {
        // Silent retry on next mutation
      }
    }, 800);

    return () => {
      if (sqlSyncTimerRef.current) {
        clearTimeout(sqlSyncTimerRef.current);
        sqlSyncTimerRef.current = null;
      }
    };
  }, [authStatus, authUser, candidates, schoolStudents, kyujins, ptPartners, resignations]);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEYS.CANDIDATES, JSON.stringify(candidates));
    } catch {
      // ignore quota errors
    }
  }, [candidates]);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEYS.SCHOOL_STUDENTS, JSON.stringify(schoolStudents));
    } catch {
      // ignore quota errors
    }
  }, [schoolStudents]);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEYS.KYUJINS, JSON.stringify(kyujins));
    } catch {
      // ignore quota errors
    }
  }, [kyujins]);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEYS.PT_PARTNERS, JSON.stringify(ptPartners));
    } catch {
      // ignore quota errors
    }
  }, [ptPartners]);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEYS.RESIGNATIONS, JSON.stringify(resignations));
    } catch {
      // ignore quota errors
    }
  }, [resignations]);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEYS.DAILY_FOCUS, JSON.stringify(dailyFocusCompleted));
    } catch {
      // ignore quota errors
    }
  }, [dailyFocusCompleted]);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEYS.GAS_URL, gasWebAppUrl);
    } catch {
      // ignore quota errors
    }
  }, [gasWebAppUrl]);

  // ==========================================================================
  // RIREKISHO FILE UPLOAD & AI CHAT AUTO-FILL PARSER (VIA SERVER PROXY)
  // ==========================================================================
  const applyExtractedRirekishoToDraft = (extracted: Record<string, unknown> = {}) => {
    if (!extracted || typeof extracted !== 'object') return 0;
    let filledCount = 0;

    setRirekishoModal((prev) => {
      if (!prev) return prev;
      const d: RirekishoDraft = { ...prev.draft };

      const mapScalar = (srcKeys: string[], targetKey: keyof RirekishoDraft) => {
        for (const k of srcKeys) {
          const v = extracted[k];
          if (typeof v === 'string' && v.trim()) {
            (d[targetKey] as string) = v.trim();
            filledCount += 1;
            return;
          }
        }
      };

      mapScalar(['furigana'], 'furigana');
      mapScalar(['fullName', 'name'], 'fullName');
      mapScalar(['birthDate'], 'birthDate');
      mapScalar(['gender'], 'gender');
      mapScalar(['postalCode'], 'postalCode');
      mapScalar(['address'], 'address');
      mapScalar(['phone'], 'phone');
      mapScalar(['email'], 'email');
      mapScalar(['nationality'], 'nationality');
      mapScalar(['visaExpiry'], 'visaExpiry');
      mapScalar(['visaType', 'visaStatus'], 'visaStatus');
      mapScalar(['nativeLanguage'], 'nativeLanguage');
      mapScalar(['spokenLanguages', 'availableLanguages'], 'availableLanguages');
      mapScalar(['speakingLevel', 'caSpeaking'], 'caSpeaking');
      mapScalar(['listeningLevel', 'caListening'], 'caListening');
      mapScalar(['motivation'], 'motivation');
      mapScalar(['selfPr'], 'selfPr');
      mapScalar(['allergies', 'allergy'], 'allergy');
      mapScalar(['dependentsCount'], 'dependentsCount');
      mapScalar(['hasSpouse', 'spouse'], 'spouse');
      mapScalar(['spouseSupportObligation', 'spouseSupport'], 'spouseSupport');
      mapScalar(['wishes', 'personalRequests'], 'personalRequests');

      if (d.birthDate) {
        const autoAge = calcJapaneseAgeFromBirthDate(d.birthDate);
        if (autoAge) d.age = autoAge;
      }

      const mapRows = (
        arr: unknown,
        targetKey: 'educationRows' | 'workRows' | 'licenseRows'
      ) => {
        if (!Array.isArray(arr)) return;
        const cleanRows = arr
          .map((item) => {
            const obj = (item || {}) as Record<string, unknown>;
            return {
              year: String(obj.year || '').trim(),
              month: String(obj.month || '').trim(),
              content: String(obj.content || '').trim(),
            };
          })
          .filter((r) => r.year || r.month || r.content);
        if (cleanRows.length > 0) {
          d[targetKey] = cleanRows;
          filledCount += cleanRows.length;
        }
      };

      mapRows(extracted.educationList, 'educationRows');
      mapRows(extracted.workHistoryList, 'workRows');
      mapRows(extracted.qualificationsList, 'licenseRows');

      return { ...prev, draft: d };
    });

    return filledCount;
  };

  const handleImportExistingRirekishoFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file || !rirekishoModal) return;
    setIsParsingRirekisho(true);
    setRirekishoStatusBanner({
      type: 'loading',
      text: `⏳ Đang đọc và phân tích "${file.name}" để điền vào mẫu Rirekisho 2 cột...`,
    });

    let extractedTextFallback = '';
    try {
      const token = await getFirebaseIdToken();
      const partsPayload = await extractRirekishoTextFromLocalFile(file);
      extractedTextFallback = partsPayload
        .map((p) => p.text || '')
        .filter(Boolean)
        .join('\n');

      try {
        const extracted = await extractRirekishoViaServer(partsPayload, token);
        const filledCount = applyExtractedRirekishoToDraft(extracted);
        if (filledCount > 0) {
          const msg = `✅ Đã bóc tách ${filledCount} mục từ "${file.name}" và điền tự động vào Sơ yếu lý lịch 2 cột!`;
          setRirekishoStatusBanner({ type: 'success', text: msg });
          showToast(msg);
          return;
        }
      } catch (aiErr) {
        if (extractedTextFallback.trim()) {
          const localParsed = buildRirekishoFallbackFromText(
            extractedTextFallback,
            rirekishoModal.person
          );
          const localFilled = applyExtractedRirekishoToDraft(
            localParsed as Record<string, unknown>
          );
          const reason =
            aiErr instanceof Error && aiErr.message.includes('GEMINI_API_KEY')
              ? 'Chưa cấu hình GEMINI_API_KEY trên máy chủ'
              : 'AI tạm gián đoạn';
          const fallbackMsg = `ℹ️ ${reason} — Đã tự động điền ${localFilled} mục cơ bản từ nội dung văn bản của "${file.name}". Vui lòng kiểm tra và bổ sung thêm nếu cần.`;
          setRirekishoStatusBanner({ type: 'success', text: fallbackMsg });
          showToast(fallbackMsg);
          return;
        }
        throw aiErr;
      }

      if (extractedTextFallback.trim()) {
        const localParsed = buildRirekishoFallbackFromText(
          extractedTextFallback,
          rirekishoModal.person
        );
        const localFilled = applyExtractedRirekishoToDraft(
          localParsed as Record<string, unknown>
        );
        if (localFilled > 0) {
          const msg = `✅ Đã trích xuất ${localFilled} mục từ văn bản "${file.name}" vào Rirekisho!`;
          setRirekishoStatusBanner({ type: 'success', text: msg });
          showToast(msg);
          return;
        }
      }

      const warnMsg = `⚠️ Không tìm thấy mục Rirekisho rõ ràng trong "${file.name}". Bạn có thể kiểm tra và chỉnh sửa trực tiếp bên dưới.`;
      setRirekishoStatusBanner({ type: 'error', text: warnMsg });
      showToast(warnMsg);
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Vui lòng kiểm tra lại định dạng file';
      const errMsg = `❌ Lỗi đọc file: ${message}`;
      setRirekishoStatusBanner({ type: 'error', text: errMsg });
      showToast(errMsg);
    } finally {
      setIsParsingRirekisho(false);
      e.target.value = '';
    }
  };

  const handleGenerateRirekishoFromChat = async () => {
    if (!rirekishoModal) return;
    const text = String(rirekishoChatInput || '').trim();
    if (!text) {
      const warn = '⚠️ Vui lòng dán nội dung chat / ghi chú phỏng vấn của ứng viên trước khi tạo';
      setRirekishoStatusBanner({ type: 'error', text: warn });
      showToast(warn);
      return;
    }

    setIsParsingRirekisho(true);
    setRirekishoStatusBanner({
      type: 'loading',
      text: '⏳ Đang phân tích lịch sử chat & hồ sơ để điền Rirekisho 2 cột...',
    });

    const p = rirekishoModal.person || {};
    const contextHeader = `【Hồ sơ CRM hiện tại】
Tên: ${p.name || ''} / FB: ${p.fbName || ''} / Quốc tịch: ${p.nationality || ''} / JLPT: ${p.jlpt || ''} / Ngành: ${p.desiredIndustry || ''} / Khu vực: ${p.desiredArea || ''} / Visa: ${p.visaStatus || ''} (${p.visaExpiry || ''}) / Kinh nghiệm: ${p.japanExperience || ''} / Chứng chỉ: ${p.passedExam || ''} / Mục tiêu: ${p.conversionGoal || ''}`;

    try {
      const token = await getFirebaseIdToken();
      const extracted = await extractRirekishoViaServer(
        [
          {
            text: `${contextHeader}\n\n【Nội dung chat & nhật ký trao đổi】\n${text}\n\nHãy bóc tách toàn bộ thông tin trên để điền vào các mục của Sơ yếu lý lịch chuẩn Nhật 2 cột (学歴・職歴・免許資格・志望動機・自己PR・本人希望欄・基本情報).`,
          },
        ],
        token
      );

      const filledCount = applyExtractedRirekishoToDraft(extracted);
      if (filledCount > 0) {
        const okMsg = `✨ AI đã trích xuất ${filledCount} mục từ nội dung chat vào Sơ yếu lý lịch 2 cột!`;
        setRirekishoStatusBanner({ type: 'success', text: okMsg });
        showToast(okMsg);
      } else {
        const localParsed = buildRirekishoFallbackFromText(text, p);
        const localFilled = applyExtractedRirekishoToDraft(localParsed as Record<string, unknown>);
        const okMsg = `✅ Đã điền ${localFilled} mục từ hồ sơ & nội dung chat bằng bộ tách nhanh!`;
        setRirekishoStatusBanner({ type: 'success', text: okMsg });
        showToast(okMsg);
      }
    } catch (err) {
      const localParsed = buildRirekishoFallbackFromText(text, p);
      const localFilled = applyExtractedRirekishoToDraft(localParsed as Record<string, unknown>);
      const reason =
        err instanceof Error && err.message.includes('GEMINI_API_KEY')
          ? 'Chưa cấu hình GEMINI_API_KEY trên máy chủ'
          : 'Kết nối AI tạm gián đoạn';
      const fallbackMsg = `ℹ️ ${reason} — Đã tự động điền ${localFilled} mục từ hồ sơ & đoạn chat bằng bộ trích xuất cục bộ. Bạn có thể chỉnh sửa trực tiếp ngay bên dưới.`;
      setRirekishoStatusBanner({ type: 'success', text: fallbackMsg });
      showToast(fallbackMsg);
    } finally {
      setIsParsingRirekisho(false);
    }
  };

  const handleDownloadRirekishoFile = () => {
    if (!rirekishoModal?.draft) return;
    try {
      const draft = rirekishoModal.draft;
      const htmlContent = buildRirekishoStandaloneHtml(draft);
      const safeName = String(draft.fullName || '応募者')
        .replace(/[\\/:*?"<>|\s]+/g, '_')
        .trim();
      const dateTag = new Date().toISOString().slice(0, 10);
      const blob = new Blob(['\uFEFF' + htmlContent], {
        type: 'text/html;charset=utf-8',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `履歴書_${safeName}_${dateTag}.html`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      showToast(`📥 「履歴書_${safeName}_${dateTag}.html」をダウンロードしました`);
    } catch {
      showToast('⚠️ 履歴書ファイルのダウンロードに失敗しました');
    }
  };

  const handlePrintRirekishoSafely = () => {
    if (!rirekishoModal?.draft) return;
    const htmlContent = buildRirekishoStandaloneHtml(rirekishoModal.draft);
    const iframe = document.createElement('iframe');
    iframe.style.position = 'fixed';
    iframe.style.right = '0';
    iframe.style.bottom = '0';
    iframe.style.width = '0';
    iframe.style.height = '0';
    iframe.style.border = '0';
    iframe.srcdoc = htmlContent;
    document.body.appendChild(iframe);
    iframe.onload = () => {
      try {
        iframe.contentWindow?.focus();
        iframe.contentWindow?.print();
      } catch {
        handleDownloadRirekishoFile();
      }
      setTimeout(() => {
        if (document.body.contains(iframe)) document.body.removeChild(iframe);
      }, 2000);
    };
  };

  const handleSyncToGas = async () => {
    if (gasWriteLocked) {
      showToast('🔒 GAS上書き保護中です。同期するには保護ボタンを解除してください');
      return;
    }
    if (!gasWebAppUrl.trim()) {
      showToast('⚠️ GAS WebアプリURLを入力してください');
      return;
    }
    setIsSyncingGas(true);
    try {
      await fetch(gasWebAppUrl.trim(), {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'syncAll',
          respectManualProtection: true,
          candidates,
          schoolStudents,
          kyujins,
          ptPartners,
          resignations: resignations.map((r) => ({
            ...r,
            nationality: normalizeResignationNationality(r.nationality),
          })),
        }),
      });
      showToast('☁️ Google Apps Scriptへデータを送信しました');
    } catch {
      showToast('⚠️ GAS送信エラーが発生しました');
    } finally {
      setIsSyncingGas(false);
    }
  };

  // ==========================================================================
  // JSON BACKUP EXPORT & RESTORE HANDLERS
  // ==========================================================================
  const handleExportBackupJson = () => {
    try {
      const payload = {
        exportedAt: new Date().toISOString(),
        candidates,
        schoolStudents,
        kyujins,
        ptPartners,
        resignations,
      };
      const blob = new Blob([JSON.stringify(payload, null, 2)], {
        type: 'application/json',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      const dateStamp = new Date().toISOString().slice(0, 10);
      a.href = url;
      a.download = `tokutei_crm_backup_${dateStamp}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      showToast('💾 バックアップJSONファイルを保存しました');
    } catch {
      showToast('⚠️ バックアップ保存中にエラーが発生しました');
    }
  };

  const handleImportBackupJson = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      const parsed = JSON.parse(text);

      const rawCandidates: Candidate[] = Array.isArray(parsed)
        ? parsed
        : Array.isArray(parsed.candidates)
        ? parsed.candidates
        : [];
      const explicitStudents: Candidate[] = Array.isArray(parsed.schoolStudents)
        ? parsed.schoolStudents
        : [];

      const separatedStudents = [...explicitStudents];
      const regularCandidates: Candidate[] = [];
      rawCandidates.forEach((item) => {
        if (item?.channel === 'school' || item?.schoolName) {
          separatedStudents.push({ ...item, channel: 'school' });
        } else {
          regularCandidates.push(item);
        }
      });

      if (regularCandidates.length > 0 || Array.isArray(parsed.candidates)) {
        setCandidates(regularCandidates);
      }
      if (separatedStudents.length > 0 || Array.isArray(parsed.schoolStudents)) {
        setSchoolStudents(separatedStudents);
      }
      if (Array.isArray(parsed.kyujins)) {
        setKyujins(parsed.kyujins.map(normalizeKyujinItem));
      }
      const importedPt = Array.isArray(parsed.ptPartners)
        ? parsed.ptPartners
        : Array.isArray(parsed.partners)
        ? parsed.partners
        : Array.isArray(parsed.ptList)
        ? parsed.ptList
        : null;
      if (importedPt) {
        setPtPartners(importedPt);
      }
      if (Array.isArray(parsed.resignations)) {
        setResignations((prev) => {
          const prevById = new Map(prev.map((item) => [item.id, item]));
          return parsed.resignations.map(
            (rawRow: Partial<ResignationRecord> & Record<string, unknown>) =>
              mergeResignationWithGasProtection(
                rawRow.id ? prevById.get(String(rawRow.id)) : undefined,
                rawRow,
                gasWriteLocked
              )
          );
        });
      }

      setGasWriteLocked(true);
      showToast(
        `📂 データ復元完了：相談者 ${regularCandidates.length}名 ／ 留学生 ${separatedStudents.length}名 ／ 求人 ${
          Array.isArray(parsed.kyujins) ? parsed.kyujins.length : kyujins.length
        }社`
      );
    } catch {
      showToast('⚠️ JSONファイルの読み込みに失敗しました');
    } finally {
      e.target.value = '';
    }
  };

  // ==========================================================================
  // DERIVED COLLECTIONS
  // ==========================================================================
  const personalCandidates = useMemo(
    () => candidates.filter((c) => c.channel !== 'company' && c.channel !== 'school'),
    [candidates]
  );

  const companyCandidates = useMemo(
    () => candidates.filter((c) => c.channel === 'company'),
    [candidates]
  );

  const allPeopleCombined = useMemo<Candidate[]>(
    () => [
      ...candidates.map((c) => ({ ...c, _isSchool: false })),
      ...schoolStudents.map((s) => ({ ...s, channel: 'school', _isSchool: true })),
    ],
    [candidates, schoolStudents]
  );

  const mendanPeopleSorted = useMemo(() => {
    return allPeopleCombined
      .map((person) => ({ person, mendan: extractNormalizedMendanInfo(person) }))
      .filter((item) => item.mendan.hasMendan)
      .sort((a, b) => b.mendan.timestamp - a.mendan.timestamp)
      .map((item) => item.person);
  }, [allPeopleCombined]);

  const allProposedPeople = useMemo(
    () => allPeopleCombined.filter((p) => isActualProposedRecord(p)),
    [allPeopleCombined]
  );

  const proposedStageCounts = useMemo(() => {
    const counts = {
      all: allProposedPeople.length,
      企業提案中: 0,
      '1次面接予定': 0,
      '2次面接予定': 0,
      内定: 0,
      '辞退・終了': 0,
    };
    allProposedPeople.forEach((p) => {
      const st = String(p.progressStatus || '').trim();
      if (st === '企業提案中' || st === '面接調整中') counts['企業提案中'] += 1;
      else if (st === '1次面接予定') counts['1次面接予定'] += 1;
      else if (st === '2次面接予定') counts['2次面接予定'] += 1;
      else if (st === '内定' || st === '入社手続き中') counts['内定'] += 1;
      else if (st.includes('辞退') || extractDeclineReason(p)) counts['辞退・終了'] += 1;
    });
    return counts;
  }, [allProposedPeople]);

  const declinedPeopleCount = useMemo(
    () => allPeopleCombined.filter((p) => Boolean(extractDeclineReason(p))).length,
    [allPeopleCombined]
  );

  const proposedMonthOptions = useMemo(() => {
    const months = new Set<string>();
    allProposedPeople.forEach((p) => {
      const mInfo = extractNormalizedMendanInfo(p);
      const dateStr = mInfo.dateLabel || p.lastContactDate || '';
      const ym = String(dateStr).slice(0, 7);
      if (/^\d{4}-\d{2}$/.test(ym)) months.add(ym);
    });
    return Array.from(months).sort((a, b) => b.localeCompare(a));
  }, [allProposedPeople]);

  const upcomingMendanAlerts = useMemo(() => {
    const now = new Date();
    const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
    return allPeopleCombined
      .map((person) => ({ person, mendan: extractNormalizedMendanInfo(person) }))
      .filter(({ person, mendan }) => {
        const st = String(person.progressStatus || '').trim();
        if (['内定', '入社手続き中', '辞退・終了'].includes(st)) return false;
        if (mendan.timestamp > 0) {
          const diffDays = Math.round((mendan.timestamp - todayStart) / 86400000);
          return diffDays >= -1 && diffDays <= 14;
        }
        return ['1次面接予定', '2次面接予定', '面接調整中'].includes(st);
      })
      .sort((a, b) => (a.mendan.timestamp || Infinity) - (b.mendan.timestamp || Infinity))
      .slice(0, 8);
  }, [allPeopleCombined]);

  const urgentFollowUpList = useMemo(() => {
    return allPeopleCombined
      .map((person) => ({ person, reason: getUrgentFollowUpReason(person) }))
      .filter((item) => Boolean(item.reason))
      .slice(0, 10);
  }, [allPeopleCombined]);

  const incompleteProfileAlerts = useMemo(() => {
    return allPeopleCombined
      .map((person) => ({ person, missing: getMissingInfoForCandidate(person) }))
      .filter(
        (item) =>
          item.missing.length > 0 &&
          !String(item.person.progressStatus || '').includes('辞退')
      )
      .slice(0, 6);
  }, [allPeopleCombined]);

  const resignationAlerts = useMemo(() => {
    return resignations
      .map((record) => ({
        record,
        assessment: assessResignationDeadline(
          record.paperworkDeadline,
          record.paperworkStatus
        ),
      }))
      .filter(
        ({ assessment }) =>
          assessment.urgency === 'overdue' ||
          assessment.urgency === 'urgent3' ||
          assessment.urgency === 'warn7'
      )
      .sort((a, b) => (a.assessment.daysLeft ?? 999) - (b.assessment.daysLeft ?? 999));
  }, [resignations]);

  const dailyFocusItems = useMemo<DailyFocusItem[]>(() => {
    const items: DailyFocusItem[] = [];
    const todayYmd = getJstTodayYmd();
    const [ty, tm, td] = todayYmd.split('-').map(Number);
    const todayUtc = Date.UTC(ty, tm - 1, td);

    // 1. Mendan / Interview today or tomorrow
    allPeopleCombined.forEach((p) => {
      const mInfo = extractNormalizedMendanInfo(p);
      const targetTs =
        parseFlexibleDateToTimestamp(p.interviewDate, ty) || mInfo.timestamp;
      if (targetTs > 0) {
        const d = new Date(targetTs + 9 * 60 * 60 * 1000);
        const ymd = d.toISOString().slice(0, 10);
        const [py, pm, pd] = ymd.split('-').map(Number);
        const diffDays = Math.round((Date.UTC(py, pm - 1, pd) - todayUtc) / 86400000);
        if (diffDays === 0 || diffDays === 1) {
          items.push({
            id: `focus_mendan_${p.id}_${ymd}`,
            category: 'mendan',
            title: `${diffDays === 0 ? 'Hôm nay' : 'Ngày mai'}: ${p.name}`,
            subtitle: p.interviewDate || mInfo.dateLabel || p.progressStatus || 'Lịch Mendan',
            urgency: diffDays === 0 ? 'critical' : 'warning',
            targetTab: p._isSchool ? 'school' : 'mendan',
            entityId: p.id,
          });
        }
      }
    });

    // 2. Resignation paperwork deadlines <= 7 days
    resignationAlerts.forEach(({ record, assessment }) => {
      items.push({
        id: `focus_resig_${record.id}`,
        category: 'resignation',
        title: `Giấy tờ nghỉ việc: ${record.fullName}`,
        subtitle: `${assessment.label} · ${record.previousCompany}`,
        urgency:
          assessment.urgency === 'overdue' || assessment.urgency === 'urgent3'
            ? 'critical'
            : 'warning',
        targetTab: 'resignation',
        entityId: record.id,
      });
    });

    // 3. Uncontacted > 3 days
    urgentFollowUpList.slice(0, 5).forEach(({ person, reason }) => {
      items.push({
        id: `focus_uncontacted_${person.id}`,
        category: 'uncontacted',
        title: `Follow-up: ${person.name}`,
        subtitle: reason || 'Chưa liên hệ > 3 ngày',
        urgency: 'warning',
        targetTab: person._isSchool ? 'school' : 'all',
        entityId: person.id,
      });
    });

    // 4. Residence card expiry < 90 days
    allPeopleCombined.forEach((p) => {
      if (!p.visaExpiry) return;
      const expTs = parseFlexibleDateToTimestamp(p.visaExpiry, ty);
      if (expTs > 0) {
        const diffDays = Math.round((expTs - todayUtc) / 86400000);
        if (diffDays >= -30 && diffDays <= 90) {
          items.push({
            id: `focus_visa_${p.id}`,
            category: 'visa_expiry',
            title: `Hạn thẻ ngoại kiều: ${p.name}`,
            subtitle:
              diffDays < 0
                ? `Đã hết hạn ${Math.abs(diffDays)} ngày (${p.visaExpiry})`
                : `Còn ${diffDays} ngày (${p.visaExpiry})`,
            urgency: diffDays <= 30 ? 'critical' : 'warning',
            targetTab: p._isSchool ? 'school' : 'all',
            entityId: p.id,
          });
        }
      }
    });

    return items.slice(0, 14);
  }, [allPeopleCombined, resignationAlerts, urgentFollowUpList]);

  const toggleDailyFocusItem = useCallback((id: string) => {
    setDailyFocusCompleted((prev) => ({
      ...prev,
      [id]: !prev[id],
    }));
  }, []);

  const filteredCandidates = useMemo(() => {
    let baseList: Candidate[] = [];
    if (activeTab === 'personal') {
      baseList = personalCandidates.map((c) => ({ ...c, _isSchool: false }));
    } else if (activeTab === 'company') {
      baseList = companyCandidates.map((c) => ({ ...c, _isSchool: false }));
    } else if (activeTab === 'school') {
      baseList = schoolStudents.map((s) => ({ ...s, channel: 'school', _isSchool: true }));
    } else if (activeTab === 'mendan') {
      baseList = mendanPeopleSorted;
    } else if (activeTab === 'proposed') {
      baseList = allProposedPeople;
    } else {
      baseList = candidates.map((c) => ({ ...c, _isSchool: false }));
    }

    const q = searchQuery.trim().toLowerCase();
    const now = new Date();
    const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();

    let result = baseList.filter((p) => {
      if (q) {
        const hay = [
          p.name,
          p.fbName,
          p.nationality,
          p.desiredIndustry,
          p.desiredArea,
          p.jlpt,
          p.visaStatus,
          p.visaExpiry,
          p.passedExam,
          p.japanExperience,
          p.conversionGoal,
          p.salaryExpectation,
          p.proposedCompany,
          getEffectiveProposedCompany(p),
          p.schoolName,
          p.negotiationLog,
          p.memo,
        ]
          .filter(Boolean)
          .join(' ')
          .toLowerCase();
        if (!hay.includes(q) && !doesAreaMatchFilter(p.desiredArea || '', q)) {
          return false;
        }
      }

      if (statusFilter !== 'all' && (p.progressStatus || '新規相談') !== statusFilter) {
        return false;
      }

      if (activeTab === 'proposed' && proposedMonthFilter !== 'all') {
        const mInfo = extractNormalizedMendanInfo(p);
        const ym = String(mInfo.dateLabel || p.lastContactDate || '').slice(0, 7);
        if (ym !== proposedMonthFilter) return false;
      }

      if (activeTab === 'proposed' && proposedStageFilter !== 'all') {
        const st = String(p.progressStatus || '').trim();
        if (proposedStageFilter === '企業提案中') {
          if (st !== '企業提案中' && st !== '面接調整中') return false;
        } else if (proposedStageFilter === '内定') {
          if (st !== '内定' && st !== '入社手続き中') return false;
        } else if (proposedStageFilter === '辞退・終了') {
          if (!st.includes('辞退') && !extractDeclineReason(p)) return false;
        } else if (st !== proposedStageFilter) {
          return false;
        }
      }

      if (careFilter === 'recent7') {
        const lastTs = parseFlexibleDateToTimestamp(p.lastContactDate, 2026);
        if (!lastTs || todayStart - lastTs > 7 * 86400000) return false;
      } else if (careFilter === 'scheduled') {
        const st = String(p.progressStatus || '').trim();
        const mInfo = extractNormalizedMendanInfo(p);
        const hasUpcoming =
          ['面談調整中', '1次面接予定', '2次面接予定', '面接調整中'].includes(st) ||
          mInfo.timestamp >= todayStart;
        if (!hasUpcoming) return false;
      } else if (careFilter === 'recentMendan') {
        const mInfo = extractNormalizedMendanInfo(p);
        if (!mInfo.hasMendan) return false;
      } else if (careFilter === 'urgent') {
        if (!getUrgentFollowUpReason(p)) return false;
      } else if (careFilter === 'declined') {
        const reason = extractDeclineReason(p);
        if (!reason) return false;
        const rq = declineReasonQuery.trim().toLowerCase();
        if (rq && !`${reason} ${p.memo || ''} ${p.name || ''}`.toLowerCase().includes(rq)) {
          return false;
        }
      }

      return true;
    });

    if (careFilter === 'recentMendan' || activeTab === 'mendan') {
      result = [...result].sort(
        (a, b) =>
          extractNormalizedMendanInfo(b).timestamp -
          extractNormalizedMendanInfo(a).timestamp
      );
    }

    return result;
  }, [
    activeTab,
    candidates,
    personalCandidates,
    companyCandidates,
    schoolStudents,
    mendanPeopleSorted,
    allProposedPeople,
    searchQuery,
    statusFilter,
    proposedMonthFilter,
    proposedStageFilter,
    careFilter,
    declineReasonQuery,
  ]);

  const kyujinStats = useMemo(() => {
    let totalSlots = 0;
    let openSlots = 0;
    let closedSlots = 0;
    let bonusSlots = 0;
    let baitoSlots = 0;
    kyujins.forEach((c) => {
      (c.areaBreakdown || []).forEach((sl) => {
        totalSlots += 1;
        if (sl.recruitmentStatus === '募集終了') closedSlots += 1;
        else openSlots += 1;
        if (sl.bonus === '賞与あり') bonusSlots += 1;
        if (/baito|アルバイト|パート/i.test(sl.employmentType || c.employmentType || '')) {
          baitoSlots += 1;
        }
      });
    });
    return {
      totalCompanies: kyujins.length,
      totalSlots,
      openSlots,
      closedSlots,
      bonusSlots,
      baitoSlots,
    };
  }, [kyujins]);

  const filteredKyujinCompanies = useMemo(() => {
    const q = kyujinSearch.trim().toLowerCase();
    return kyujins
      .map((comp) => {
        const matchedSlots = (comp.areaBreakdown || []).filter((sl) => {
          if (kyujinStatusFilter === 'open' && sl.recruitmentStatus === '募集終了') return false;
          if (kyujinStatusFilter === 'closed' && sl.recruitmentStatus !== '募集終了') return false;
          if (kyujinBonusOnly && sl.bonus !== '賞与あり') return false;

          const slotInd = sl.industry || comp.industry || '';
          const slotSector =
            sl.sector || comp.sector || detectSectorFromIndustryOrText(slotInd, sl.jobDescription || '');
          const slotEmp =
            sl.employmentType ||
            comp.employmentType ||
            detectEmploymentTypeFromText(`${sl.salary || ''} ${sl.hourlyWage || ''}`);

          if (kyujinSectorFilter !== 'all' && slotSector !== kyujinSectorFilter) return false;
          if (kyujinIndustryFilter !== 'all' && slotInd !== kyujinIndustryFilter) return false;
          if (
            kyujinEmploymentTypeFilter !== 'all' &&
            slotEmp !== kyujinEmploymentTypeFilter
          ) {
            return false;
          }
          if (kyujinFacilityFilter !== 'all' && sl.facilityType !== kyujinFacilityFilter) {
            return false;
          }

          if (q) {
            const hay = [
              comp.companyName,
              slotSector,
              slotInd,
              slotEmp,
              sl.jobCode,
              sl.area,
              sl.shiftTime,
              sl.salary,
              sl.hourlyWage,
              sl.estimatedNetSalary,
              sl.nightShiftAllowance,
              sl.minJlpt,
              sl.facilityType,
              sl.experienceRequirement,
              sl.experienceDetail,
              sl.jobDescription,
            ]
              .filter(Boolean)
              .join(' ')
              .toLowerCase();
            if (!hay.includes(q)) return false;
          }
          return true;
        });

        return { ...comp, areaBreakdown: matchedSlots };
      })
      .filter((comp) => comp.areaBreakdown.length > 0);
  }, [
    kyujins,
    kyujinSearch,
    kyujinStatusFilter,
    kyujinBonusOnly,
    kyujinSectorFilter,
    kyujinIndustryFilter,
    kyujinEmploymentTypeFilter,
    kyujinFacilityFilter,
  ]);

  // ==========================================================================
  // CANDIDATE CRUD & CHAT AUTO-EXTRACTION
  // ==========================================================================
  const handleMarkContactedToday = (person: Candidate, isSchool: boolean) => {
    const now = new Date();
    const todayIso = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(
      now.getDate()
    ).padStart(2, '0')}`;
    const updater = (list: Candidate[]) =>
      list.map((p) => (p.id === person.id ? { ...p, lastContactDate: todayIso } : p));
    if (isSchool) setSchoolStudents(updater);
    else setCandidates(updater);
    showToast(`✅ ${person.name} さんの最終連絡日を本日（${todayIso}）に更新しました`);
  };

  const handleQuickStatusChange = (personId: string, isSchool: boolean, newStatus: string) => {
    const todayIso = new Date().toISOString().slice(0, 10);
    const updater = (list: Candidate[]) =>
      list.map((p) =>
        p.id === personId ? { ...p, progressStatus: newStatus, lastContactDate: todayIso } : p
      );
    if (isSchool) setSchoolStudents(updater);
    else setCandidates(updater);
    showToast(`✅ ステータスを「${newStatus}」に更新しました`);
  };

  const handleProposeCompanyToCandidate = (
    person: Candidate,
    isSchool: boolean,
    companyName: string
  ) => {
    const todayIso = new Date().toISOString().slice(0, 10);
    const updater = (list: Candidate[]) =>
      list.map((p) =>
        p.id === person.id
          ? {
              ...p,
              proposedCompany: companyName,
              progressStatus: '企業提案中',
              lastContactDate: todayIso,
              memo: `${p.memo ? p.memo + '\n' : ''}[${todayIso}] ${companyName} へ提案`,
            }
          : p
      );
    if (isSchool) setSchoolStudents(updater);
    else setCandidates(updater);
    showToast(`📨 ${person.name} さんに「${companyName}」を提案済みに設定しました`);
  };

  const mergeExtractedCandidateData = (
    base: Candidate,
    extracted: Partial<Candidate> & { memoLine?: string; fbUrl?: string },
    rawSourceText?: string
  ): Candidate => {
    const mergedExp = extracted.japanExperience
      ? base.japanExperience && !base.japanExperience.includes(extracted.japanExperience)
        ? `${base.japanExperience} / ${extracted.japanExperience}`
        : extracted.japanExperience
      : base.japanExperience;
    const mergedExam = extracted.passedExam
      ? base.passedExam && !base.passedExam.includes(extracted.passedExam)
        ? `${base.passedExam} / ${extracted.passedExam}`
        : extracted.passedExam
      : base.passedExam;
    const todayIso = getJstTodayYmd();
    const rawTrimmed = String(rawSourceText || '').trim();
    const rawArchive = rawTrimmed ? `[${todayIso} Nguyên văn đoạn chat]\n${rawTrimmed}` : '';

    // Guarantee 100% verbatim preservation in BOTH negotiationLog and memo
    let incomingLog = String(extracted.negotiationLog || '').trim();
    if (rawTrimmed) {
      if (!incomingLog) {
        incomingLog = rawArchive;
      } else if (!incomingLog.includes(rawTrimmed)) {
        incomingLog = `${incomingLog}\n\n${rawArchive}`;
      }
    }

    const incomingMemo = rawArchive || String(extracted.memoLine || incomingLog || '').trim();

    const shouldReplaceName =
      Boolean(extracted.name) &&
      (!base.name || base.name === 'NEW CANDIDATE' || base.name === base.fbName);

    return {
      ...base,
      name: shouldReplaceName ? extracted.name! : base.name,
      fbName: extracted.fbName || base.fbName,
      fbProfileUrl: extracted.fbProfileUrl || extracted.fbUrl || base.fbProfileUrl,
      nationality: extracted.nationality || base.nationality,
      gender: extracted.gender || base.gender,
      birthDate: extracted.birthDate || base.birthDate,
      age: extracted.age || base.age,
      email: extracted.email || base.email,
      address: extracted.address || base.address,
      jlpt: extracted.jlpt || base.jlpt,
      phone: extracted.phone || base.phone,
      visaStatus: extracted.visaStatus || base.visaStatus,
      visaExpiry: extracted.visaExpiry || base.visaExpiry,
      desiredIndustry: extracted.desiredIndustry || base.desiredIndustry,
      desiredArea: extracted.desiredArea || base.desiredArea,
      passedExam: mergedExam,
      japanExperience: mergedExp,
      conversionGoal: extracted.conversionGoal || base.conversionGoal,
      salaryExpectation: extracted.salaryExpectation || base.salaryExpectation,
      startDate: extracted.startDate || base.startDate,
      graduationDate: extracted.graduationDate || base.graduationDate,
      mendanDate: extracted.mendanDate || base.mendanDate,
      interviewDate: extracted.interviewDate || base.interviewDate,
      progressStatus: extracted.progressStatus || base.progressStatus,
      proposedCompany: extracted.proposedCompany || base.proposedCompany,
      lastContactDate: extracted.lastContactDate || base.lastContactDate || todayIso,
      negotiationLog: incomingLog
        ? base.negotiationLog && !base.negotiationLog.includes(incomingLog)
          ? `${base.negotiationLog}\n\n${incomingLog}`
          : base.negotiationLog || incomingLog
        : base.negotiationLog,
      memo: incomingMemo
        ? base.memo && (!rawTrimmed || !base.memo.includes(rawTrimmed))
          ? `${base.memo}\n\n${incomingMemo}`
          : base.memo || incomingMemo
        : base.memo,
    };
  };

  const handleApplyChatToEditingPerson = async () => {
    if (!editingPerson || !editChatSnippet.trim()) {
      showToast('⚠️ Vui lòng dán đoạn chat hoặc ghi chú trao đổi trước');
      return;
    }
    const rawSnippet = editChatSnippet.trim();
    setIsExtractingCandidateAi(true);
    let usedAi = false;
    try {
      let extracted: (Partial<Candidate> & { memoLine?: string; fbUrl?: string }) | null = null;
      try {
        const token = await getFirebaseIdToken();
        const aiData = await extractCandidateViaServer({ text: rawSnippet }, token);
        if (aiData && Object.keys(aiData).length > 0) {
          extracted = aiData;
          usedAi = true;
        }
      } catch {
        // Fallback smoothly to deterministic zero-truncation local parser
      }
      const fallback = parseChatIntoUpdates(rawSnippet);
      const mergedExtract: Partial<Candidate> & { memoLine?: string; fbUrl?: string } = {
        ...fallback,
        ...(extracted || {}),
        name: extracted?.name || fallback.name || '',
        fbName: extracted?.fbName || fallback.fbName || '',
        fbProfileUrl: extracted?.fbProfileUrl || fallback.fbProfileUrl || '',
        visaStatus: extracted?.visaStatus || fallback.visaStatus || '',
        visaExpiry: extracted?.visaExpiry || fallback.visaExpiry || '',
        jlpt: extracted?.jlpt || fallback.jlpt || '',
        desiredIndustry: extracted?.desiredIndustry || fallback.desiredIndustry || '',
        desiredArea: extracted?.desiredArea || fallback.desiredArea || '',
        passedExam: extracted?.passedExam || fallback.passedExam || '',
        japanExperience: extracted?.japanExperience || fallback.japanExperience || '',
        conversionGoal: extracted?.conversionGoal || fallback.conversionGoal || '',
        salaryExpectation: extracted?.salaryExpectation || fallback.salaryExpectation || '',
        startDate: extracted?.startDate || fallback.startDate || '',
        graduationDate: extracted?.graduationDate || fallback.graduationDate || '',
        phone: extracted?.phone || fallback.phone || '',
        email: extracted?.email || fallback.email || '',
        birthDate: extracted?.birthDate || fallback.birthDate || '',
        age: extracted?.age || fallback.age || '',
        nationality: extracted?.nationality || fallback.nationality || '',
        proposedCompany: extracted?.proposedCompany || fallback.proposedCompany || '',
        mendanDate: extracted?.mendanDate || fallback.mendanDate || '',
        interviewDate: extracted?.interviewDate || fallback.interviewDate || '',
        progressStatus: extracted?.progressStatus || fallback.progressStatus || '',
        negotiationLog: extracted?.negotiationLog || fallback.negotiationLog,
        memoLine: fallback.memoLine,
      };
      setEditingPerson((prev) =>
        prev
          ? { ...prev, data: mergeExtractedCandidateData(prev.data, mergedExtract, rawSnippet) }
          : prev
      );
      setEditChatSnippet('');
      showToast(
        usedAi
          ? '✨ AI đã bóc tách & bảo toàn 100% nguyên văn vào hồ sơ!'
          : '✨ Đã bóc tách bằng bộ lọc dự phòng & lưu 100% nguyên văn!'
      );
    } finally {
      setIsExtractingCandidateAi(false);
    }
  };

  const handleExtractCandidateScreenshot = async (
    file: File,
    target: 'edit' | 'fbChat'
  ) => {
    if (!file.type.startsWith('image/')) {
      showToast('⚠️ Vui lòng chọn hoặc dán ảnh chụp màn hình hợp lệ');
      return;
    }
    if (target === 'edit') setIsExtractingCandidateAi(true);
    else setIsExtractingFbChatAi(true);

    try {
      const compressed = await compressImageFileToBase64(file, 1600, 0.78);
      const token = await getFirebaseIdToken();
      const extracted = await extractCandidateViaServer(
        {
          parts: [
            {
              inlineData: {
                mimeType: compressed.mimeType,
                data: compressed.base64Data,
              },
            },
          ],
        },
        token
      );

      if (target === 'edit' && editingPerson) {
        setEditingPerson((prev) =>
          prev
            ? {
                ...prev,
                data: mergeExtractedCandidateData(
                  prev.data,
                  extracted,
                  extracted.negotiationLog || extracted.memo || ''
                ),
              }
            : prev
        );
        showToast('✨ Đã đọc ảnh chụp Messenger / Thẻ ngoại kiều và điền vào hồ sơ!');
      } else if (target === 'fbChat') {
        const formattedSummary = [
          extracted.name ? `Họ tên khai sinh: ${extracted.name}` : '',
          extracted.fbName ? `FB: ${extracted.fbName}` : '',
          extracted.fbProfileUrl ? `Link FB: ${extracted.fbProfileUrl}` : '',
          extracted.nationality ? `Quốc tịch: ${extracted.nationality}` : '',
          extracted.visaStatus ? `Visa: ${extracted.visaStatus}` : '',
          extracted.visaExpiry ? `Hạn thẻ: ${extracted.visaExpiry}` : '',
          extracted.jlpt ? `JLPT: ${extracted.jlpt}` : '',
          extracted.desiredIndustry ? `Ngành: ${extracted.desiredIndustry}` : '',
          extracted.desiredArea ? `Khu vực: ${extracted.desiredArea}` : '',
          extracted.startDate ? `Nhập công ty: ${extracted.startDate}` : '',
          extracted.japanExperience ? `Kinh nghiệm: ${extracted.japanExperience}` : '',
          extracted.passedExam ? `Chứng chỉ: ${extracted.passedExam}` : '',
          extracted.phone ? `SĐT: ${extracted.phone}` : '',
          extracted.email ? `Email: ${extracted.email}` : '',
          extracted.mendanDate ? `Mendan: ${extracted.mendanDate}` : '',
          extracted.interviewDate ? `Phỏng vấn cty: ${extracted.interviewDate}` : '',
          extracted.conversionGoal ? `Mục tiêu: ${extracted.conversionGoal}` : '',
          extracted.salaryExpectation ? `Lương kỳ vọng: ${extracted.salaryExpectation}` : '',
          extracted.negotiationLog || extracted.memo || '',
        ]
          .filter(Boolean)
          .join('\n');
        setFbChatRawText((prev) =>
          prev.trim() ? `${prev.trim()}\n\n${formattedSummary}` : formattedSummary
        );
        showToast('✨ Đã bóc tách nội dung từ ảnh chụp vào khung đăng ký nhanh!');
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Không thể đọc ảnh chụp lúc này';
      showToast(`⚠️ Lỗi đọc ảnh: ${msg}`);
    } finally {
      if (target === 'edit') setIsExtractingCandidateAi(false);
      else setIsExtractingFbChatAi(false);
    }
  };

  const handleSubmitFbChatModal = async () => {
    const raw = fbChatRawText.trim();
    if (!raw) {
      showToast('⚠️ Vui lòng dán nội dung chat hoặc ảnh chụp trước');
      return;
    }
    setIsExtractingFbChatAi(true);
    let usedAi = false;
    try {
      let aiExtracted: Partial<Candidate> | null = null;
      try {
        const token = await getFirebaseIdToken();
        const aiRes = await extractCandidateViaServer({ text: raw }, token);
        if (aiRes && Object.keys(aiRes).length > 0) {
          aiExtracted = aiRes;
          usedAi = true;
        }
      } catch {
        // Fallback to zero-truncation parser
      }
      const fallback = parseChatIntoUpdates(raw);
      const extracted: Partial<Candidate> & { memoLine?: string } = {
        ...fallback,
        ...(aiExtracted || {}),
        name: aiExtracted?.name || fallback.name || '',
        fbName: aiExtracted?.fbName || fallback.fbName || '',
        fbProfileUrl: aiExtracted?.fbProfileUrl || fallback.fbProfileUrl || '',
        visaStatus: aiExtracted?.visaStatus || fallback.visaStatus || '',
        visaExpiry: aiExtracted?.visaExpiry || fallback.visaExpiry || '',
        jlpt: aiExtracted?.jlpt || fallback.jlpt || '',
        desiredIndustry: aiExtracted?.desiredIndustry || fallback.desiredIndustry || '',
        desiredArea: aiExtracted?.desiredArea || fallback.desiredArea || '',
        passedExam: aiExtracted?.passedExam || fallback.passedExam || '',
        japanExperience: aiExtracted?.japanExperience || fallback.japanExperience || '',
        conversionGoal: aiExtracted?.conversionGoal || fallback.conversionGoal || '',
        salaryExpectation: aiExtracted?.salaryExpectation || fallback.salaryExpectation || '',
        startDate: aiExtracted?.startDate || fallback.startDate || '',
        graduationDate: aiExtracted?.graduationDate || fallback.graduationDate || '',
        phone: aiExtracted?.phone || fallback.phone || '',
        email: aiExtracted?.email || fallback.email || '',
        birthDate: aiExtracted?.birthDate || fallback.birthDate || '',
        age: aiExtracted?.age || fallback.age || '',
        nationality: aiExtracted?.nationality || fallback.nationality || '',
        gender: aiExtracted?.gender || fallback.gender || '',
        address: aiExtracted?.address || fallback.address || '',
        proposedCompany: aiExtracted?.proposedCompany || fallback.proposedCompany || '',
        mendanDate: aiExtracted?.mendanDate || fallback.mendanDate || '',
        interviewDate: aiExtracted?.interviewDate || fallback.interviewDate || '',
        progressStatus: aiExtracted?.progressStatus || fallback.progressStatus || '',
        negotiationLog: aiExtracted?.negotiationLog || fallback.negotiationLog,
        memoLine: fallback.memoLine,
      };

      if (fbChatMode === 'update' && fbChatTargetId) {
        const existingTarget = allPeopleCombined.find((p) => p.id === fbChatTargetId);
        if (!existingTarget) {
          showToast('⚠️ Không tìm thấy ứng viên cần cập nhật');
          return;
        }
        const mergedTarget = mergeExtractedCandidateData(
          { ...existingTarget, channel: fbChatChannel || existingTarget.channel },
          extracted,
          raw
        );
        const targetIsSchool = mergedTarget.channel === 'school';
        if (targetIsSchool) {
          setCandidates((prev) => prev.filter((p) => p.id !== fbChatTargetId));
          setSchoolStudents((prev) => {
            const exists = prev.some((p) => p.id === fbChatTargetId);
            return exists
              ? prev.map((p) => (p.id === fbChatTargetId ? { ...mergedTarget, _isSchool: true } : p))
              : [{ ...mergedTarget, _isSchool: true }, ...prev];
          });
        } else {
          setSchoolStudents((prev) => prev.filter((p) => p.id !== fbChatTargetId));
          setCandidates((prev) => {
            const exists = prev.some((p) => p.id === fbChatTargetId);
            return exists
              ? prev.map((p) => (p.id === fbChatTargetId ? { ...mergedTarget, _isSchool: false } : p))
              : [{ ...mergedTarget, _isSchool: false }, ...prev];
          });
        }
        setFbChatModalOpen(false);
        setFbChatRawText('');
        showToast(
          usedAi
            ? '🔄 AI đã cập nhật & bảo toàn 100% nguyên văn đoạn chat!'
            : '🔄 Đã cập nhật bằng bộ lọc dự phòng & bảo toàn 100% nguyên văn!'
        );
        return;
      }

      const todayIso = getJstTodayYmd();
      const missingNamePrefix =
        !extracted.name && extracted.fbName
          ? `[Chưa có họ tên khai sinh — tạm hiển thị theo tên FB: ${extracted.fbName}]\n`
          : '';
      const rawArchive = `[${todayIso} Nguyên văn đoạn chat]\n${missingNamePrefix}${raw}`;
      const fullNegLog =
        extracted.negotiationLog && !extracted.negotiationLog.includes(raw)
          ? `${extracted.negotiationLog}\n\n${rawArchive}`
          : extracted.negotiationLog || rawArchive;

      const newPerson: Candidate = {
        id: `cand_${Date.now()}`,
        name: extracted.name || extracted.fbName || 'Ứng viên FB (chưa rõ tên)',
        fbName: extracted.fbName || '',
        fbProfileUrl: extracted.fbProfileUrl || '',
        channel: fbChatChannel,
        nationality: extracted.nationality || '国籍不明',
        gender: extracted.gender || '',
        age: extracted.age || '',
        birthDate: extracted.birthDate || '',
        jlpt: extracted.jlpt || '要確認',
        desiredIndustry: extracted.desiredIndustry || '',
        desiredArea: extracted.desiredArea || '',
        visaStatus: extracted.visaStatus || '',
        visaExpiry: extracted.visaExpiry || '',
        passedExam: extracted.passedExam || '',
        japanExperience: extracted.japanExperience || '',
        conversionGoal: extracted.conversionGoal || '',
        salaryExpectation: extracted.salaryExpectation || '',
        startDate: extracted.startDate || '',
        graduationDate: extracted.graduationDate || '',
        phone: extracted.phone || '',
        email: extracted.email || '',
        address: extracted.address || '',
        progressStatus: extracted.progressStatus || '新規相談',
        proposedCompany: extracted.proposedCompany || '',
        mendanDate: extracted.mendanDate || '',
        interviewDate: extracted.interviewDate || '',
        lastContactDate: extracted.lastContactDate || todayIso,
        negotiationLog: fullNegLog,
        memo: rawArchive,
      };

      if (fbChatChannel === 'school') {
        setSchoolStudents((prev) => [newPerson, ...prev]);
        setActiveTab('school');
      } else {
        setCandidates((prev) => [newPerson, ...prev]);
        setActiveTab(fbChatChannel === 'company' ? 'company' : 'personal');
      }

      setFbChatModalOpen(false);
      setFbChatRawText('');
      showToast(
        `✅ Đã tạo hồ sơ mới tại ${
          fbChatChannel === 'company'
            ? 'FB会社公式'
            : fbChatChannel === 'school'
            ? 'Trường Nhật ngữ'
            : 'FB個人'
        }: ${newPerson.name}${usedAi ? '' : ' (bộ lọc dự phòng)'}`
      );
    } finally {
      setIsExtractingFbChatAi(false);
    }
  };

  const handleQuickChannelChange = (person: Candidate, nextChannel: ChannelType) => {
    if (person.channel === nextChannel) return;
    const isCurrentlySchool = Boolean(person._isSchool || person.channel === 'school');
    const willBeSchool = nextChannel === 'school';
    const updatedPerson: Candidate = {
      ...person,
      channel: nextChannel,
      _isSchool: willBeSchool,
    };

    if (isCurrentlySchool && !willBeSchool) {
      setSchoolStudents((prev) => prev.filter((s) => s.id !== person.id));
      setCandidates((prev) => [updatedPerson, ...prev.filter((c) => c.id !== person.id)]);
    } else if (!isCurrentlySchool && willBeSchool) {
      setCandidates((prev) => prev.filter((c) => c.id !== person.id));
      setSchoolStudents((prev) => [updatedPerson, ...prev.filter((s) => s.id !== person.id)]);
    } else if (willBeSchool) {
      setSchoolStudents((prev) => prev.map((s) => (s.id === person.id ? updatedPerson : s)));
    } else {
      setCandidates((prev) => prev.map((c) => (c.id === person.id ? updatedPerson : c)));
    }

    const label =
      nextChannel === 'company'
        ? 'FB会社公式'
        : nextChannel === 'school'
        ? '日本語学校・留学生'
        : 'FB個人';
    showToast(`🔄 Đã chuyển "${person.name}" sang kênh ${label} (bảo toàn 100% dữ liệu)`);
  };

  const handleSaveEditingPerson = () => {
    if (!editingPerson) return;
    const { data, isSchool, isNew } = editingPerson;
    const trimmedName = String(data.name || '').trim();
    const trimmedFb = String(data.fbName || '').trim();
    if (!trimmedName && !trimmedFb) {
      showToast('⚠️ Vui lòng nhập Họ tên khai sinh hoặc Tên Facebook');
      return;
    }

    const effectiveName = trimmedName || trimmedFb;
    const fallbackNote =
      !trimmedName && trimmedFb && !(data.memo || '').includes('tạm dùng tên FB')
        ? `[Chưa có họ tên khai sinh — tạm dùng tên FB: ${trimmedFb}]${data.memo ? '\n' + data.memo : ''}`
        : data.memo;

    const isSchoolTarget = Boolean(data.channel === 'school' || (isSchool && !data.channel));
    const prevPerson = isNew
      ? undefined
      : allPeopleCombined.find((x) => x.id === data.id);

    const savedPerson: Candidate = {
      ...data,
      name: effectiveName,
      memo: fallbackNote,
      channel: isSchoolTarget ? 'school' : data.channel || 'personal',
      _isSchool: isSchoolTarget,
    };

    if (isSchoolTarget) {
      if (!isNew) {
        setCandidates((prev) => prev.filter((c) => c.id !== savedPerson.id));
      }
      setSchoolStudents((prev) => {
        const exists = prev.some((s) => s.id === savedPerson.id);
        return exists
          ? prev.map((s) => (s.id === savedPerson.id ? savedPerson : s))
          : [savedPerson, ...prev];
      });
    } else {
      if (!isNew) {
        setSchoolStudents((prev) => prev.filter((s) => s.id !== savedPerson.id));
      }
      setCandidates((prev) => {
        const exists = prev.some((c) => c.id === savedPerson.id);
        return exists
          ? prev.map((c) => (c.id === savedPerson.id ? savedPerson : c))
          : [savedPerson, ...prev];
      });
    }

    setEditingPerson(null);
    showToast(isNew ? '✅ 新規登録しました' : '✅ プロフィールを更新しました');

    const scheduleChanged =
      !prevPerson ||
      (prevPerson.mendanDate || '').trim() !==
        (savedPerson.mendanDate || '').trim() ||
      (prevPerson.interviewDate || '').trim() !==
        (savedPerson.interviewDate || '').trim() ||
      (prevPerson.proposedCompany || '').trim() !==
        (savedPerson.proposedCompany || '').trim();

    const nextDraft = buildCandidateCalendarEventDraft(savedPerson);

    if (nextDraft && scheduleChanged) {
      const hasExistingEvent = Boolean(savedPerson.googleCalendarEventId);
      setConfirmDialog({
        title: hasExistingEvent
          ? 'Cập nhật lịch trên Google Calendar?'
          : 'Tạo sự kiện mới trên Google Calendar?',
        message: `Đã lưu lịch hẹn của ${savedPerson.name} trong CRM.\n• Tiêu đề: ${nextDraft.eventTitle}\n• Thời gian: ${nextDraft.displayDateTime}\n• Nhắc nhở: 30 phút & 10 phút trước.\n\nBạn có muốn đồng bộ lên Google Calendar ngay bây giờ không?`,
        confirmLabel: hasExistingEvent ? 'Cập nhật Calendar' : 'Tạo lịch trên Calendar',
        cancelLabel: 'Chỉ lưu CRM',
        tone: 'accent',
        onConfirm: () => {
          setConfirmDialog(null);
          void executeCalendarSyncForDraft(nextDraft);
        },
      });
    } else if (!nextDraft && prevPerson?.googleCalendarEventId) {
      const staleEventId = prevPerson.googleCalendarEventId;
      setConfirmDialog({
        title: 'Hủy sự kiện trên Google Calendar?',
        message: `Bạn đã xóa ngày giờ lịch hẹn của ${savedPerson.name}. Bạn có muốn xóa sự kiện tương ứng khỏi Google Calendar không?`,
        confirmLabel: 'Xóa trên Google Calendar',
        cancelLabel: 'Giữ lại',
        tone: 'danger',
        onConfirm: () => {
          setConfirmDialog(null);
          void executeCalendarDelete(savedPerson.id, staleEventId, true);
        },
      });
    }
  };

  const handleDeletePerson = (personId: string, isSchool: boolean) => {
    const targetPerson = (isSchool ? schoolStudents : candidates).find(
      (p) => p.id === personId
    );
    const linkedEventId = targetPerson?.googleCalendarEventId;

    setConfirmDialog({
      title: '削除の確認',
      message: linkedEventId
        ? `この相談者データ（${targetPerson?.name || ''}）を削除しますか？関連付けられた Google Calendar の予定も同時に削除されます。`
        : 'この相談者データを削除してもよろしいですか？',
      confirmLabel: '削除を実行する',
      cancelLabel: 'キャンセル',
      tone: 'danger',
      onConfirm: () => {
        if (isSchool) {
          setSchoolStudents((prev) => prev.filter((s) => s.id !== personId));
        } else {
          setCandidates((prev) => prev.filter((c) => c.id !== personId));
        }
        setConfirmDialog(null);
        showToast('🗑️ 削除しました');
        if (linkedEventId) {
          void executeCalendarDelete(personId, linkedEventId, false);
        }
      },
    });
  };

  // ==========================================================================
  // KYUJIN IMPORT & BULK ACTIONS
  // ==========================================================================
  const handleQuickUpdateSlotField = (
    companyId: string,
    slotId: string,
    field: keyof KyujinSlot,
    value: string
  ) => {
    setKyujins((prev) =>
      prev.map((comp) => {
        if (comp.id !== companyId) return comp;
        const updatedSlots = (comp.areaBreakdown || []).map((sl) => {
          if (sl.slotId !== slotId) return sl;
          const nextSlot = { ...sl, [field]: value };
          if (field === 'bonus' && value !== '賞与あり') {
            nextSlot.bonusDetail = '';
          }
          if (field === 'industry' && value !== '介護') {
            nextSlot.facilityType = '';
            nextSlot.careQualification = '';
          }
          return nextSlot;
        });
        const nextCompIndustry = field === 'industry' ? value : comp.industry;
        return normalizeKyujinItem({
          ...comp,
          industry: nextCompIndustry,
          areaBreakdown: updatedSlots,
        });
      })
    );
    if (field === 'industry') {
      showToast(`🏭 分野を「${value}」に変更しました`);
    }
  };

  const toggleSelectSlot = (slotId: string) => {
    setSelectedSlotIds((prev) =>
      prev.includes(slotId) ? prev.filter((id) => id !== slotId) : [...prev, slotId]
    );
  };

  const handleDeleteSelectedSlots = () => {
    if (selectedSlotIds.length === 0) return;
    const count = selectedSlotIds.length;
    setConfirmDialog({
      title: '選択した求人枠の削除',
      message: `選択中の ${count} 枠を削除してもよろしいですか？`,
      onConfirm: () => {
        const idSet = new Set(selectedSlotIds);
        setKyujins((prev) =>
          prev
            .map((comp) => ({
              ...comp,
              areaBreakdown: (comp.areaBreakdown || []).filter((sl) => !idSet.has(sl.slotId)),
            }))
            .filter((comp) => comp.areaBreakdown.length > 0)
        );
        setSelectedSlotIds([]);
        setConfirmDialog(null);
        showToast(`🗑️ ${count} 件の求人枠を削除しました`);
      },
    });
  };

  const handleDeleteAllKyujins = () => {
    setConfirmDialog({
      title: '全求人一括削除の確認',
      message: '登録されているすべての求人票・企業データを一括削除しますか？（元に戻せません）',
      onConfirm: () => {
        setKyujins([]);
        setSelectedSlotIds([]);
        setConfirmDialog(null);
        showToast('🗑️ 全求人データを一括削除しました');
      },
    });
  };

  const handleDeleteSingleCompany = (companyId: string, companyName: string) => {
    setConfirmDialog({
      title: '企業ごと削除の確認',
      message: `「${companyName}」の求人データをすべて削除しますか？`,
      onConfirm: () => {
        setKyujins((prev) => prev.filter((c) => c.id !== companyId));
        setConfirmDialog(null);
        showToast(`🗑️ 「${companyName}」を削除しました`);
      },
    });
  };

  const mergeExtractedKyujinEntries = useCallback(
    (
      extractedEntries: Array<{
        companyName: string;
        sector?: string;
        industry: string;
        employmentType?: string;
        recruitmentStatus: string;
        slot: KyujinSlot;
      }>
    ) => {
      if (extractedEntries.length === 0) return;
      setKyujins((prev) => {
        const updated = [...prev];
        extractedEntries.forEach((entry) => {
          const existingIdx = updated.findIndex(
            (c) =>
              String(c.companyName || '').replace(/\s+/g, '') ===
              String(entry.companyName || '').replace(/\s+/g, '')
          );
          if (existingIdx >= 0) {
            const target = updated[existingIdx];
            updated[existingIdx] = normalizeKyujinItem({
              ...target,
              areaBreakdown: [...(target.areaBreakdown || []), entry.slot],
            });
          } else {
            updated.unshift(
              normalizeKyujinItem({
                id: `kj_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
                companyName: entry.companyName,
                sector: entry.sector || entry.slot.sector,
                industry: entry.industry,
                employmentType: entry.employmentType || entry.slot.employmentType,
                recruitmentStatus: '募集中',
                areaBreakdown: [entry.slot],
              })
            );
          }
        });
        return updated;
      });
    },
    []
  );

  const mapServerKyujinItemsToEntries = (
    items: Array<Record<string, string>>,
    sourceName: string
  ) => {
    const fallbackComp = sourceName
      .replace(/^求人票[_・\-\s]*/i, '')
      .replace(/様.*$/, '')
      .replace(/\.[^.]+$/, '')
      .trim();

    return items.map((parsed, idx) => {
      const cleanArea = formatPublicArea(parsed.area || '') || String(parsed.area || '').trim();
      const cleanDesc = sanitizeJobDescriptionForCandidate(parsed.jobDescription || '');
      const shiftLabel = String(parsed.shiftTime || '').trim();
      const industry = detectIndustryFromText(
        `${parsed.industry || ''} ${cleanDesc}`,
        `${sourceName} ${parsed.industry || ''}`
      );
      const sector =
        String(parsed.sector || '').trim() ||
        detectSectorFromIndustryOrText(industry, `${parsed.industry || ''} ${cleanDesc}`);
      const employmentType =
        String(parsed.employmentType || '').trim() ||
        detectEmploymentTypeFromText(
          `${parsed.salary || ''} ${parsed.hourlyWage || ''} ${cleanDesc}`,
          sourceName
        );
      const rawJlpt = String(parsed.minJlpt || '').trim();
      const normalizedJlpt = /^(N[1-5]|日常会話|不問)$/.test(rawJlpt)
        ? rawJlpt
        : '要確認';
      const derivedCode =
        String(parsed.jobCode || '').trim() ||
        cleanArea ||
        `${sourceName.replace(/\.[^.]+$/, '')}-${idx + 1}`;
      const bonusNorm = normalizeBonusInfo(parsed.bonus, parsed.bonusDetail, cleanDesc);

      return {
        companyName: String(parsed.companyName || fallbackComp || '新規取込企業')
          .replace(/\s*代表者.*$/, '')
          .trim(),
        sector,
        industry,
        employmentType,
        recruitmentStatus: '募集中',
        slot: {
          slotId: `slot_vis_${Date.now()}_${idx}_${Math.random().toString(36).slice(2, 7)}`,
          sector,
          industry,
          employmentType,
          jobCode: derivedCode,
          area: cleanArea || '勤務地要確認',
          shiftTime: shiftLabel,
          salary: String(
            parsed.salary || (parsed.hourlyWage ? `時給 ${parsed.hourlyWage}` : '給与要確認')
          ).trim(),
          nightShiftAllowance: String(parsed.nightShiftAllowance || '').trim(),
          estimatedNetSalary: String(parsed.estimatedNetSalary || '').trim(),
          hourlyWage: String(parsed.hourlyWage || '').trim(),
          commuteAllowance: String(parsed.commuteAllowance || '').trim(),
          shiftsPerWeek: String(parsed.shiftsPerWeek || '').trim(),
          bonus: bonusNorm.bonus,
          bonusDetail: bonusNorm.bonusDetail,
          minJlpt: normalizedJlpt,
          housingSupport: normalizeHousingDisplay(parsed.housingSupport),
          recruitmentStatus: '募集中',
          careQualification: parsed.careQualification || '',
          experienceRequirement: String(parsed.experienceRequirement || '').trim(),
          experienceDetail: String(parsed.experienceDetail || '').trim(),
          jobDescription: cleanDesc,
        } satisfies KyujinSlot,
      };
    });
  };

  const handleImportKyujinFileList = async (files: File[]) => {
    if (files.length === 0) return;
    setIsImportingKyujin(true);
    const extractedEntries: Array<{
      companyName: string;
      sector?: string;
      industry: string;
      employmentType?: string;
      recruitmentStatus: string;
      slot: KyujinSlot;
    }> = [];

    try {
      for (let fIdx = 0; fIdx < files.length; fIdx++) {
        const file = files[fIdx];
        setKyujinImportProgress({
          current: fIdx + 1,
          total: files.length,
          label: file.name,
        });
        // Yield briefly to keep UI responsive on large multi-file batches
        await new Promise((r) => setTimeout(r, 10));

        const lower = file.name.toLowerCase();
        if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) {
          const excelEntries = await parseExcelFileToKyujinEntries(file);
          extractedEntries.push(...excelEntries);
        } else {
          const token = await getFirebaseIdToken();
          try {
            const items = await extractKyujinViaServer(file, token);
            extractedEntries.push(...mapServerKyujinItemsToEntries(items, file.name));
          } catch (fileErr) {
            // Fallback for plain text / csv when Gemini key isn't set
            if (lower.endsWith('.txt') || lower.endsWith('.csv') || file.type.startsWith('text/')) {
              const rawText = await file.text();
              const fbItems = parseRawKyujinTextFallback(rawText, file.name);
              extractedEntries.push(...mapServerKyujinItemsToEntries(fbItems, file.name));
            } else {
              throw fileErr;
            }
          }
        }
      }

      if (extractedEntries.length > 0) {
        mergeExtractedKyujinEntries(extractedEntries);
        showToast(`✅ Đã đọc và nhập ${extractedEntries.length} đơn / vị trí tuyển dụng mới!`);
      } else {
        showToast('⚠️ Không tìm thấy dữ liệu đơn tuyển dụng hợp lệ trong file đã chọn');
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : 'Lỗi đọc phiếu tuyển dụng';
      showToast(`⚠️ ${message}`);
    } finally {
      setIsImportingKyujin(false);
      setKyujinImportProgress(null);
    }
  };

  const handleImportKyujinFiles = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []);
    e.target.value = '';
    if (files.length === 0) return;
    await handleImportKyujinFileList(files);
  };

  const handleImportKyujinRawText = async () => {
    const raw = kyujinRawText.trim();
    if (!raw) {
      showToast('⚠️ Vui lòng dán nội dung tin tuyển dụng hoặc danh sách đơn hàng');
      return;
    }
    setIsImportingKyujin(true);
    setKyujinImportProgress({ current: 1, total: 1, label: 'Phân tích văn bản Kyujin' });
    try {
      const token = await getFirebaseIdToken();
      let items: Array<Record<string, string>> = [];
      try {
        items = await extractKyujinRawTextViaServer(raw, token);
      } catch {
        items = parseRawKyujinTextFallback(raw, 'Đơn dán văn bản');
      }
      if (items.length === 0) {
        items = parseRawKyujinTextFallback(raw, 'Đơn dán văn bản');
      }
      const entries = mapServerKyujinItemsToEntries(items, 'Đơn dán văn bản');
      if (entries.length > 0) {
        mergeExtractedKyujinEntries(entries);
        setKyujinRawModalOpen(false);
        setKyujinRawText('');
        showToast(`✅ Đã bóc tách và lưu ${entries.length} đơn tuyển dụng từ văn bản!`);
      } else {
        showToast('⚠️ Không bóc tách được đơn tuyển dụng nào từ nội dung trên');
      }
    } finally {
      setIsImportingKyujin(false);
      setKyujinImportProgress(null);
    }
  };

  // ==========================================================================
  // CANDIDATE CARD RENDERER (HIGH DENSITY TECH - VARIATION 3)
  // ==========================================================================
  const renderCandidateCard = (person: Candidate) => {
    const isSchool = Boolean(person._isSchool || person.channel === 'school');
    const mendanInfo = extractNormalizedMendanInfo(person);
    const missingItems = getMissingInfoForCandidate(person);
    const matchEval = evaluateCandidateKyujinMatches(person, kyujins);
    const matchingKyujins = matchEval.matches;
    const proposedComp = getEffectiveProposedCompany(person);
    const declineReason = extractDeclineReason(person);
    const now = new Date();
    const todayIso = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(
      now.getDate()
    ).padStart(2, '0')}`;
    const isContactedToday = String(person.lastContactDate || '').trim() === todayIso;

    return (
      <div
        key={person.id}
        className="rounded-lg border border-[#27272a] bg-[#141417] p-4 transition-colors hover:border-[#3f3f46]"
      >
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div>
            <div className="flex flex-wrap items-center gap-2.5">
              <span className="text-base font-semibold tracking-tight text-[#edeeef]">
                {person.name}
              </span>
              {person.fbName && person.fbName !== person.name && (
                <span className="font-mono text-[11px] text-[#38bdf8]">
                  FB: {person.fbName}
                </span>
              )}
              <select
                value={isSchool ? 'school' : person.channel === 'company' ? 'company' : 'personal'}
                onChange={(e) =>
                  handleQuickChannelChange(person, e.target.value as ChannelType)
                }
                title="Chuyển nhanh kênh nguồn (FB個人 / FB会社公式 / 留学生)"
                className="rounded border border-[#27272a] bg-[#0c0c0e] px-2 py-0.5 font-mono text-[10px] uppercase tracking-wider text-[#94949e] hover:border-[#22c55e] hover:text-[#edeeef] focus:border-[#22c55e] focus:outline-none cursor-pointer"
              >
                <option value="personal" className="bg-[#141417] text-[#edeeef]">
                  FB個人
                </option>
                <option value="company" className="bg-[#141417] text-[#edeeef]">
                  FB会社公式
                </option>
                <option value="school" className="bg-[#141417] text-[#edeeef]">
                  留学生
                </option>
              </select>
              {isContactedToday ? (
                <span className="font-mono text-[10px] text-[#22c55e] tabular-nums">
                  CONTACTED {todayIso}
                </span>
              ) : (
                person.lastContactDate && (
                  <span className="font-mono text-[10px] text-[#94949e] tabular-nums">
                    LAST: {person.lastContactDate}
                  </span>
                )
              )}
              {proposedComp && (
                <span className="font-mono text-[10px] text-sky-400">
                  PROPOSED: {proposedComp}
                </span>
              )}
            </div>
            <div className="mt-1 font-mono text-[10px] uppercase tracking-wider text-[#22c55e]">
              STATUS: {person.progressStatus || '新規相談'}
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-1.5">
            <select
              value={person.progressStatus || '新規相談'}
              onChange={(e) => handleQuickStatusChange(person.id, isSchool, e.target.value)}
              title="進捗ステータスを1タップ変更"
              className="rounded-md border border-[#27272a] bg-[#0c0c0e] px-2.5 py-1.5 text-[11px] font-semibold text-[#edeeef] focus:border-[#22c55e] focus:outline-none whitespace-nowrap shrink-0"
            >
              {PROGRESS_STATUSES.map((st) => (
                <option key={st} value={st} className="bg-[#141417] text-[#edeeef]">
                  {st}
                </option>
              ))}
            </select>

            <button
              type="button"
              onClick={() => handleMarkContactedToday(person, isSchool)}
              className={`rounded-md border px-2.5 py-1.5 text-[11px] font-semibold transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                isContactedToday
                  ? 'border-[#22c55e]/50 bg-[#22c55e]/15 text-[#22c55e]'
                  : 'border-[#27272a] bg-transparent text-[#94949e] hover:border-[#94949e] hover:text-[#edeeef]'
              }`}
            >
              {isContactedToday ? '本日連絡済' : '今日連絡済み'}
            </button>

            <button
              type="button"
              onClick={() =>
                setReplyModal({
                  person,
                  scenario: missingItems.length > 0 ? 'first_mendan' : 'recommend_job',
                })
              }
              className="rounded-md bg-[#22c55e] px-3 py-1.5 text-[11px] font-semibold text-[#0c0c0e] transition-opacity hover:opacity-90 whitespace-nowrap shrink-0 cursor-pointer"
            >
              返信文提案
            </button>

            <button
              type="button"
              onClick={() => {
                setRirekishoStatusBanner(null);
                setRirekishoChatInput(person.memo || '');
                setRirekishoModal({
                  person,
                  draft: createInitialRirekishoDraft(person),
                });
              }}
              className="rounded-md border border-[#27272a] bg-transparent px-3 py-1.5 text-[11px] font-semibold text-[#edeeef] hover:border-[#94949e] whitespace-nowrap shrink-0 cursor-pointer"
            >
              履歴書作成
            </button>

            <button
              type="button"
              disabled={syncingCalendarId === person.id}
              onClick={() => requestSyncCandidateCalendar(person)}
              className={`inline-flex items-center gap-1 rounded-md border px-2.5 py-1.5 text-[11px] font-semibold transition-colors whitespace-nowrap shrink-0 cursor-pointer disabled:opacity-50 ${
                person.googleCalendarEventId
                  ? 'border-[#22c55e]/50 bg-[#22c55e]/10 text-[#22c55e] hover:bg-[#22c55e]/20'
                  : 'border-[#27272a] bg-transparent text-[#94949e] hover:border-[#22c55e] hover:text-[#22c55e]'
              }`}
              title={
                person.googleCalendarEventId
                  ? 'Đã đồng bộ Google Calendar · Bấm để cập nhật'
                  : 'Đồng bộ lịch Mendan / Phỏng vấn lên Google Calendar'
              }
            >
              <Calendar className="h-3 w-3" aria-hidden="true" />
              {syncingCalendarId === person.id
                ? 'Đang đồng bộ...'
                : person.googleCalendarEventId
                ? 'Đã đồng bộ Calendar'
                : 'Đồng bộ Google Calendar'}
            </button>

            {(() => {
              const fbMessengerUrl = getSafeFacebookMessengerUrl(person);
              if (fbMessengerUrl) {
                return (
                  <a
                    href={fbMessengerUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-1 rounded-md border border-[#38bdf8]/40 bg-[#38bdf8]/10 px-2.5 py-1.5 text-[11px] font-semibold text-[#38bdf8] hover:bg-[#38bdf8]/20 whitespace-nowrap shrink-0"
                    title="Mở trang cá nhân / Messenger của ứng viên"
                  >
                    <MessageCircle className="h-3 w-3" aria-hidden="true" />
                    Mở Messenger
                  </a>
                );
              }
              if (person.fbSenderId) {
                return (
                  <a
                    href="https://business.facebook.com/latest/inbox/all"
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-1 rounded-md border border-[#38bdf8]/40 bg-[#38bdf8]/10 px-2.5 py-1.5 text-[11px] font-semibold text-[#38bdf8] hover:bg-[#38bdf8]/20 whitespace-nowrap shrink-0"
                    title={`Mở hộp thư Meta Business (Sender ID: ${person.fbSenderId})`}
                  >
                    <MessageCircle className="h-3 w-3" aria-hidden="true" />
                    Hộp thư FB
                  </a>
                );
              }
              return null;
            })()}

            {person.googleCalendarEventId &&
              isValidGoogleCalendarUrl(person.googleCalendarEventLink) && (
                <a
                  href={person.googleCalendarEventLink}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="inline-flex items-center gap-1 rounded-md border border-[#27272a] px-2 py-1.5 text-[11px] font-semibold text-[#edeeef] hover:border-[#22c55e] hover:text-[#22c55e] whitespace-nowrap shrink-0"
                  title="Mở sự kiện trên Google Calendar"
                >
                  <ExternalLink className="h-3 w-3" aria-hidden="true" />
                  Mở Calendar
                </a>
              )}

            <button
              type="button"
              onClick={() =>
                setEditingPerson({
                  data: { ...person },
                  isSchool,
                  isNew: false,
                })
              }
              className="rounded-md border border-[#27272a] bg-transparent px-2.5 py-1.5 text-[11px] font-semibold text-[#94949e] hover:border-[#94949e] hover:text-[#edeeef] whitespace-nowrap shrink-0 cursor-pointer"
            >
              編集
            </button>

            <button
              type="button"
              onClick={() => handleDeletePerson(person.id, isSchool)}
              className="rounded-md border border-[#ef4444]/30 bg-transparent px-2.5 py-1.5 text-[11px] font-semibold text-[#ef4444] hover:bg-[#ef4444]/10 whitespace-nowrap shrink-0 cursor-pointer"
            >
              削除
            </button>
          </div>
        </div>

        {/* Structured Key-Value Grid */}
        <div className="mt-3 grid grid-cols-2 gap-3 border-t border-dashed border-[#27272a] pt-3 sm:grid-cols-3 lg:grid-cols-5">
          <div>
            <div className="font-mono text-[10px] uppercase tracking-wider text-[#94949e]">
              FB NAME
            </div>
            <div className="mt-0.5 break-words text-xs font-medium text-[#edeeef]">
              {person.fbName || person.name || '-'}
            </div>
          </div>
          <div>
            <div className="font-mono text-[10px] uppercase tracking-wider text-[#94949e]">
              NATIONALITY / VISA
            </div>
            <div className="mt-0.5 text-xs font-medium text-[#edeeef] tabular-nums">
              {person.nationality || '未設定'}
              {person.age ? ` / ${person.age}歳` : ''}
            </div>
            {person.visaStatus && (
              <div className="mt-0.5 text-[11px] text-[#94949e]">{person.visaStatus}</div>
            )}
          </div>
          <div>
            <div className="font-mono text-[10px] uppercase tracking-wider text-[#94949e]">
              LANGUAGE &amp; CERT
            </div>
            <div className="mt-0.5 font-mono text-xs font-semibold text-[#22c55e]">
              {person.jlpt || '要確認'}
            </div>
            {(person.passedExam || person.japanExperience) && (
              <div className="mt-0.5 text-[11px] text-[#94949e]">
                {[person.japanExperience, person.passedExam].filter(Boolean).join(' · ')}
              </div>
            )}
          </div>
          <div>
            <div className="font-mono text-[10px] uppercase tracking-wider text-[#94949e]">
              FIELD / DESIRED AREAS
            </div>
            <div className="mt-0.5 break-words text-xs font-medium text-[#edeeef]">
              {person.desiredIndustry || '未設定'}
              {person.desiredArea ? ` (${person.desiredArea})` : ''}
            </div>
          </div>
          <div>
            <div className="font-mono text-[10px] uppercase tracking-wider text-[#94949e]">
              {person.visaExpiry ? 'VISA EXP' : mendanInfo.hasMendan ? 'MENDAN' : 'SCHOOL'}
            </div>
            <div
              className={`mt-0.5 font-mono text-xs font-medium tabular-nums ${
                person.visaExpiry
                  ? 'text-[#ef4444]'
                  : mendanInfo.hasMendan
                  ? 'text-[#f59e0b]'
                  : 'text-[#94949e]'
              }`}
            >
              {person.visaExpiry || mendanInfo.dateLabel || person.schoolName || '-'}
            </div>
          </div>
        </div>

        {(person.conversionGoal || person.salaryExpectation) && (
          <div className="mt-2.5 grid grid-cols-1 gap-2 rounded-md border border-[#27272a] bg-[#0c0c0e]/70 px-3 py-2 text-xs sm:grid-cols-2">
            {person.conversionGoal && (
              <div>
                <span className="font-mono text-[10px] uppercase text-[#38bdf8]">MỤC TIÊU: </span>
                <span className="text-[#edeeef]">{person.conversionGoal}</span>
              </div>
            )}
            {person.salaryExpectation && (
              <div>
                <span className="font-mono text-[10px] uppercase text-[#22c55e]">LƯƠNG KỲ VỌNG: </span>
                <span className="text-[#edeeef]">{person.salaryExpectation}</span>
              </div>
            )}
          </div>
        )}

        {(person.negotiationLog || person.memo) && (
          <div className="mt-2.5 whitespace-pre-line rounded-md border border-[#27272a] bg-[#0c0c0e] px-3 py-2.5 text-xs leading-relaxed text-[#edeeef]/90">
            {person.negotiationLog && person.negotiationLog !== person.memo
              ? `${person.negotiationLog}${person.memo ? `\n\n--- Ghi chú thêm ---\n${person.memo}` : ''}`
              : person.memo || person.negotiationLog}
          </div>
        )}

        {/* Next Best Action Bar */}
        {(() => {
          const isInterviewPassed =
            Boolean(person.interviewDate && parseFlexibleDateToTimestamp(person.interviewDate) < Date.now()) ||
            /面接|内定/.test(String(person.progressStatus || ''));
          return (
            <div className="mt-2.5 flex flex-wrap items-center justify-between gap-2 rounded-md border border-[#27272a] bg-[#0c0c0e] px-3 py-2 text-xs">
              <div className="flex flex-wrap items-center gap-2">
                <span className="font-mono text-[10px] uppercase tracking-wider text-[#22c55e]">
                  Gợi ý tiếp theo:
                </span>
                {missingItems.length > 0 && (
                  <span className="text-[#f59e0b]">
                    Thiếu {missingItems.length} mục ({missingItems.map((m) => m.vi).join(', ')})
                  </span>
                )}
              </div>
              <div className="flex flex-wrap items-center gap-1.5">
                {missingItems.length > 0 && (
                  <button
                    type="button"
                    onClick={() =>
                      setReplyModal({ person, scenario: 'missing_info_request' })
                    }
                    className="rounded border border-[#f59e0b]/60 bg-[#f59e0b]/10 px-2.5 py-1 text-[11px] font-semibold text-[#f59e0b] hover:bg-[#f59e0b]/20 whitespace-nowrap cursor-pointer"
                  >
                    Xin thông tin thiếu
                  </button>
                )}
                <button
                  type="button"
                  onClick={() => setReplyModal({ person, scenario: 'recommend_job' })}
                  className="rounded border border-[#22c55e]/50 bg-[#22c55e]/10 px-2.5 py-1 text-[11px] font-semibold text-[#22c55e] hover:bg-[#22c55e]/20 whitespace-nowrap cursor-pointer"
                >
                  Gửi 3 Job phù hợp nhất
                </button>
                {isInterviewPassed && (
                  <button
                    type="button"
                    onClick={() =>
                      setReplyModal({ person, scenario: 'post_interview_followup' })
                    }
                    className="rounded border border-[#38bdf8]/50 bg-[#38bdf8]/10 px-2.5 py-1 text-[11px] font-semibold text-[#38bdf8] hover:bg-[#38bdf8]/20 whitespace-nowrap cursor-pointer"
                  >
                    Hỏi thăm sau phỏng vấn
                  </button>
                )}
              </div>
            </div>
          );
        })()}

        {declineReason && (
          <div className="mt-2.5 flex flex-wrap items-center justify-between gap-2 rounded-md border border-[#ef4444]/40 bg-[#ef4444]/10 px-3 py-2 text-xs">
            <span className="font-semibold text-[#ef4444]">
              辞退理由: <span className="font-normal text-[#edeeef]">{declineReason}</span>
            </span>
            <button
              type="button"
              onClick={() =>
                setEditingPerson({
                  data: { ...person },
                  isSchool,
                  isNew: false,
                })
              }
              className="rounded border border-[#ef4444]/40 px-2 py-0.5 text-[11px] font-semibold text-[#ef4444] hover:bg-[#ef4444]/15 whitespace-nowrap shrink-0 cursor-pointer"
            >
              理由を追記・修正
            </button>
          </div>
        )}

        {(matchingKyujins.length > 0 || Boolean(matchEval.noticeMessage)) && (
          <div className="mt-3 pt-2">
            <div className="mb-1.5 font-mono text-[10px] uppercase tracking-wider text-[#94949e]">
              Job Match Suggestions ({matchingKyujins.length})
            </div>
            {matchEval.noticeMessage && (
              <div
                className={`mb-2 rounded-md border px-3 py-2 text-[11px] ${
                  matchingKyujins.length > 0
                    ? 'border-[#f59e0b]/40 bg-[#f59e0b]/10 text-[#f59e0b]'
                    : 'border-[#27272a] bg-[#0c0c0e] text-[#94949e]'
                }`}
              >
                {matchEval.noticeMessage}
              </div>
            )}
            {matchingKyujins.length > 0 && (
              <div className="space-y-1.5">
                {matchingKyujins.map(({ company, slot, score, reasons }) => {
                const isBaito = /baito|アルバイト|パート/i.test(
                  slot.employmentType || company.employmentType || ''
                );
                return (
                  <div
                    key={slot.slotId}
                    className="flex flex-wrap items-center justify-between gap-2 rounded-md border border-[#27272a] bg-[#0c0c0e] px-3 py-2 text-xs"
                  >
                    <div className="min-w-0 flex-1 space-y-1">
                      <div className="flex flex-wrap items-center gap-2">
                        <span
                          className={`rounded px-1.5 py-0.5 font-mono text-[10px] font-bold tabular-nums ${
                            score >= 75
                              ? 'bg-[#22c55e]/20 text-[#22c55e]'
                              : score >= 55
                              ? 'bg-[#38bdf8]/20 text-[#38bdf8]'
                              : 'bg-[#f59e0b]/20 text-[#f59e0b]'
                          }`}
                        >
                          Khớp {score}%
                        </span>
                        <span className="font-semibold text-[#edeeef]">
                          {company.companyName}
                        </span>
                        <span
                          className={`rounded px-1.5 py-0.5 font-mono text-[10px] ${
                            isBaito
                              ? 'bg-[#38bdf8]/15 text-[#38bdf8]'
                              : 'bg-[#27272a] text-[#94949e]'
                          }`}
                        >
                          {isBaito ? 'Baito' : slot.employmentType || 'Tokutei'}
                        </span>
                        <span className="text-[#94949e]">{formatPublicArea(slot.area)}</span>
                        <span className="font-mono text-[#22c55e] tabular-nums">
                          {isBaito && slot.hourlyWage ? `時給 ${slot.hourlyWage}` : slot.salary}
                        </span>
                        <span className="font-mono text-[11px] text-[#94949e]">
                          {slot.minJlpt}
                        </span>
                      </div>
                      {reasons && reasons.length > 0 && (
                        <div className="flex flex-wrap gap-1.5">
                          {reasons.map((rs, rIdx) => (
                            <span
                              key={rIdx}
                              className="rounded border border-[#27272a] bg-[#141417] px-1.5 py-0.5 text-[10px] text-[#edeeef]/85"
                            >
                              {rs}
                            </span>
                          ))}
                        </div>
                      )}
                    </div>
                    <div className="inline-flex items-center gap-1.5 shrink-0">
                      <button
                        type="button"
                        onClick={async () => {
                          const post = buildAnonymousJobPostText(company, slot, 'vi');
                          const ok = await copyTextSafely(post);
                          showToast(
                            ok
                              ? `📋 Đã tạo & sao chép tin mời ứng tuyển ẩn danh (${score}% khớp)`
                              : '⚠️ Không thể sao chép tự động'
                          );
                        }}
                        className="rounded border border-[#38bdf8]/40 bg-[#38bdf8]/10 px-2.5 py-1 text-[11px] font-semibold text-[#38bdf8] hover:bg-[#38bdf8]/20 whitespace-nowrap cursor-pointer"
                      >
                        Tạo tin mời ứng tuyển
                      </button>
                      <button
                        type="button"
                        onClick={() => setPreviewModal({ company, slot })}
                        className="rounded border border-[#27272a] px-2.5 py-1 text-[11px] font-semibold text-[#94949e] hover:border-[#94949e] hover:text-[#edeeef] whitespace-nowrap cursor-pointer"
                      >
                        Xem ẩn danh
                      </button>
                      <button
                        type="button"
                        onClick={() =>
                          handleProposeCompanyToCandidate(
                            person,
                            isSchool,
                            company.companyName
                          )
                        }
                        className="rounded border border-[#22c55e]/50 bg-[#22c55e]/15 px-2.5 py-1 text-[11px] font-semibold text-[#22c55e] hover:bg-[#22c55e]/25 whitespace-nowrap cursor-pointer"
                      >
                        Đề xuất
                      </button>
                    </div>
                  </div>
                );
              })}
              </div>
            )}
          </div>
        )}
      </div>
    );
  };

  const crmSummaryForAi = useMemo(() => {
    const openCompanies = kyujins
      .slice(0, 8)
      .map(
        (c) =>
          `${c.companyName} (${c.industry || 'Chưa rõ ngành'}, ${c.areaBreakdown?.[0]?.area || 'Nhật Bản'})`
      )
      .join('; ');
    const upcomingList = upcomingMendanAlerts
      .slice(0, 5)
      .map(
        ({ person, mendan }) =>
          `${person.name} (${person.jlpt || 'Chưa rõ JLPT'}, lịch: ${mendan.dateLabel || person.progressStatus || 'Đang điều phối'})`
      )
      .join('; ');
    return [
      `Tổng ứng viên FB: ${candidates.length}, Du học sinh trường tiếng: ${schoolStudents.length}, Doanh nghiệp tuyển dụng: ${kyujinStats.totalCompanies} (${kyujinStats.totalSlots} suất, ${kyujinStats.openSlots} đang mở), Đối tác PT: ${ptPartners.length}.`,
      `Đang cần follow-up khẩn: ${urgentFollowUpList.length} hồ sơ. Lịch Mendan sắp tới (${upcomingMendanAlerts.length}): ${upcomingList || 'Chưa có'}.`,
      `Một số đơn hàng nổi bật: ${openCompanies || 'Chưa có'}.`,
    ].join('\n');
  }, [
    candidates.length,
    schoolStudents.length,
    kyujinStats,
    kyujins,
    ptPartners.length,
    urgentFollowUpList.length,
    upcomingMendanAlerts,
  ]);

  return (
    <div className="min-h-screen bg-[#0c0c0e] text-[#edeeef]">
      {toastMsg && (
        <div className="fixed bottom-20 right-5 z-[10000] rounded-lg border border-[#22c55e]/50 bg-[#141417] px-4 py-3 text-xs font-bold text-[#22c55e] shadow-xl">
          {toastMsg}
        </div>
      )}

      {/* Hidden File Inputs */}
      <input
        ref={backupInputRef}
        type="file"
        accept=".json,application/json"
        onChange={handleImportBackupJson}
        className="hidden"
      />
      <input
        ref={kyujinFileInputRef}
        type="file"
        multiple
        accept=".xlsx,.xls,.pdf,.docx,.csv,.txt,image/*"
        onChange={handleImportKyujinFiles}
        className="hidden"
      />
      <input
        ref={rirekishoFileInputRef}
        type="file"
        accept=".pdf,.xlsx,.xls,.csv,.doc,.docx,image/*"
        onChange={handleImportExistingRirekishoFile}
        className="hidden"
      />

      {viewMode === 'home' ? (
        <HomePage
          candidateCount={candidates.length}
          schoolCount={schoolStudents.length}
          companyCount={kyujinStats.totalCompanies}
          slotCount={kyujinStats.totalSlots}
          ptCount={ptPartners.length}
          upcomingMendanCount={upcomingMendanAlerts.length}
          userEmail={authUser?.email || null}
          authLoading={authLoading}
          onStartWorkspace={(targetTab) => {
            if (targetTab) setActiveTab(targetTab);
            setViewMode('workspace');
          }}
          onSignInGoogle={handleSignInGoogle}
          onSignOut={handleSignOut}
          onOpenChatbot={() => setChatbotOpen(true)}
        />
      ) : (
        <div className="flex min-h-screen flex-col lg:grid lg:h-screen lg:grid-cols-[240px_1fr_320px] lg:grid-rows-[64px_1fr] lg:overflow-hidden">
          {/* TOP HEADER BAR */}
          <header className="flex min-h-[64px] flex-wrap items-center justify-between gap-3 border-b border-[#27272a] bg-[#141417] px-5 py-2 lg:col-span-3">
            <div className="flex flex-wrap items-center gap-3">
              <button
                type="button"
                onClick={() => setViewMode('home')}
                className="flex items-center gap-2 text-sm font-bold tracking-tight text-[#edeeef] hover:text-[#22c55e] transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                title="Trang chủ"
              >
                <svg
                  width="18"
                  height="18"
                  viewBox="0 0 24 24"
                  fill="none"
                  stroke="currentColor"
                  strokeWidth="2.2"
                  aria-hidden="true"
                >
                  <rect x="3" y="3" width="18" height="18" rx="2" />
                  <path d="M9 3v18" />
                  <path d="M3 9h18" />
                </svg>
                特定技能FB相談者管理
              </button>

              <button
                type="button"
                onClick={() => {
                  const next = !gasWriteLocked;
                  setGasWriteLocked(next);
                  showToast(
                    next
                      ? '🔒 GAS上書き保護を有効にしました'
                      : '🔓 GAS上書き保護を解除しました（同期可能）'
                  );
                }}
                className={`rounded px-2 py-0.5 font-mono text-[10px] font-bold uppercase tracking-wider transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                  gasWriteLocked
                    ? 'bg-[#22c55e] text-[#0c0c0e]'
                    : 'border border-[#f59e0b] bg-[#f59e0b]/15 text-[#f59e0b]'
                }`}
              >
                {gasWriteLocked ? 'GAS上書き保護中' : 'GAS同期許可中'}
              </button>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <span className="hidden font-mono text-[10px] uppercase tracking-wider text-[#94949e] xl:inline tabular-nums">
                相談者 {candidates.length}名 · 留学生 {schoolStudents.length}名 · 求人{' '}
                {kyujinStats.totalCompanies}社 ({kyujinStats.totalSlots}枠)
              </span>

              <button
                type="button"
                onClick={() => backupInputRef.current?.click()}
                className="rounded-md border border-[#27272a] bg-transparent px-3 py-1.5 text-[11px] font-semibold text-[#94949e] hover:border-[#94949e] hover:text-[#edeeef] whitespace-nowrap shrink-0 cursor-pointer"
              >
                データ復元
              </button>

              <button
                type="button"
                onClick={handleExportBackupJson}
                className="rounded-md border border-[#27272a] bg-transparent px-3 py-1.5 text-[11px] font-semibold text-[#94949e] hover:border-[#94949e] hover:text-[#edeeef] whitespace-nowrap shrink-0 cursor-pointer"
              >
                バックアップ保存
              </button>

              <button
                type="button"
                onClick={openMessengerLinkModal}
                className={`inline-flex items-center gap-1.5 rounded-md border px-3 py-1.5 text-[11px] font-semibold transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                  webhookStatus?.configured || fbPageLink
                    ? 'border-[#38bdf8]/50 bg-[#38bdf8]/10 text-[#38bdf8] hover:bg-[#38bdf8]/20'
                    : 'border-[#27272a] bg-[#0c0c0e] text-[#edeeef] hover:border-[#38bdf8]'
                }`}
                title="Cấu hình liên kết Facebook Messenger Fanpage & Webhook"
              >
                <span
                  className={`inline-block h-1.5 w-1.5 rounded-full ${
                    webhookStatus?.configured
                      ? 'bg-[#22c55e]'
                      : fbPageLink
                      ? 'bg-[#38bdf8]'
                      : 'bg-[#f59e0b]'
                  }`}
                />
                <Link2 className="h-3.5 w-3.5" aria-hidden="true" />
                Liên kết Messenger FB
              </button>

              <button
                type="button"
                onClick={() => setChatbotOpen(true)}
                className="inline-flex items-center gap-1.5 rounded-md bg-[#22c55e] px-3 py-1.5 text-[11px] font-semibold text-[#0c0c0e] hover:opacity-90 whitespace-nowrap shrink-0 cursor-pointer"
              >
                <MessageSquareText className="h-3.5 w-3.5" aria-hidden="true" />
                Trợ lý AI
              </button>
            </div>
          </header>

          {/* MOBILE / TABLET TAB STRIP (< lg) */}
          <div className="flex items-center gap-1 overflow-x-auto border-b border-[#27272a] bg-[#0c0c0e] px-4 py-2 lg:hidden">
            {(
              [
                { id: 'all', label: `Ứng viên (${candidates.length})` },
                { id: 'personal', label: `FB cá nhân (${personalCandidates.length})` },
                { id: 'company', label: `Fanpage (${companyCandidates.length})` },
                { id: 'school', label: `Trường Nhật ngữ (${schoolStudents.length})` },
                { id: 'mendan', label: `Mendan/PV (${mendanPeopleSorted.length})` },
                { id: 'proposed', label: `Đã tiến cử (${allProposedPeople.length})` },
                { id: 'resignation', label: `Quản lý nghỉ việc (${resignations.length})` },
                { id: 'pt', label: `Đối tác PT (${ptPartners.length})` },
                { id: 'kyujin', label: `Tin tuyển dụng (${kyujinStats.totalCompanies}社)` },
                { id: 'report', label: 'Báo cáo' },
              ] as Array<{ id: MainTabId; label: string }>
            ).map((tab) => (
              <button
                key={tab.id}
                type="button"
                onClick={() => setActiveTab(tab.id)}
                className={`rounded-md px-2.5 py-1 text-xs font-semibold whitespace-nowrap shrink-0 tabular-nums cursor-pointer ${
                  activeTab === tab.id
                    ? 'bg-[#22c55e] text-[#0c0c0e]'
                    : 'text-[#94949e] hover:bg-[#141417] hover:text-[#edeeef]'
                }`}
              >
                {tab.label}
              </button>
            ))}
          </div>

          {/* LEFT SIDEBAR (DESKTOP) */}
          <aside className="hidden border-r border-[#27272a] bg-[#0c0c0e] py-4 lg:flex lg:flex-col lg:overflow-y-auto">
            <div className="mb-6">
              <div className="px-5 pb-2 font-mono text-[10px] uppercase tracking-widest text-[#94949e]">
                Database
              </div>
              {(
                [
                  { id: 'all', label: 'Tất cả ứng viên', count: `${candidates.length}` },
                  { id: 'personal', label: 'FB cá nhân', count: `${personalCandidates.length}` },
                  { id: 'company', label: 'Fanpage công ty', count: `${companyCandidates.length}` },
                  { id: 'school', label: 'Trường Nhật ngữ / DHS', count: `${schoolStudents.length}` },
                ] as Array<{ id: MainTabId; label: string; count: string }>
              ).map((item) => (
                <button
                  key={item.id}
                  type="button"
                  onClick={() => setActiveTab(item.id)}
                  className={`flex w-full items-center justify-between px-5 py-2.5 text-left text-xs transition-colors cursor-pointer ${
                    activeTab === item.id
                      ? 'border-r-2 border-[#22c55e] bg-[rgba(34,197,94,0.12)] font-semibold text-[#22c55e]'
                      : 'text-[#94949e] hover:bg-[#141417] hover:text-[#edeeef]'
                  }`}
                >
                  <span>{item.label}</span>
                  <span className="font-mono text-[11px] tabular-nums">{item.count}</span>
                </button>
              ))}
            </div>

            <div className="mb-6">
              <div className="px-5 pb-2 font-mono text-[10px] uppercase tracking-widest text-[#94949e]">
                Workflow
              </div>
              {(
                [
                  { id: 'mendan', label: 'Lịch Mendan / Phỏng vấn', count: `${mendanPeopleSorted.length}` },
                  { id: 'proposed', label: 'Đã tiến cử', count: `${allProposedPeople.length}` },
                  {
                    id: 'resignation',
                    label: 'Quản lý nghỉ việc',
                    count:
                      resignationAlerts.length > 0
                        ? `${resignationAlerts.length} gấp / ${resignations.length}`
                        : `${resignations.length}`,
                  },
                  { id: 'pt', label: 'Đối tác PT', count: `${ptPartners.length}` },
                  { id: 'kyujin', label: 'Tin tuyển dụng (Kyujin)', count: `${kyujinStats.totalCompanies}社` },
                  { id: 'report', label: 'Báo cáo phân tích', count: `${kyujinStats.totalSlots}枠` },
                ] as Array<{ id: MainTabId; label: string; count: string }>
              ).map((item) => (
                <button
                  key={item.id}
                  type="button"
                  onClick={() => setActiveTab(item.id)}
                  className={`flex w-full items-center justify-between px-5 py-2.5 text-left text-xs transition-colors cursor-pointer ${
                    activeTab === item.id
                      ? 'border-r-2 border-[#22c55e] bg-[rgba(34,197,94,0.12)] font-semibold text-[#22c55e]'
                      : 'text-[#94949e] hover:bg-[#141417] hover:text-[#edeeef]'
                  }`}
                >
                  <span>{item.label}</span>
                  <span className="font-mono text-[11px] tabular-nums">{item.count}</span>
                </button>
              ))}
            </div>

            <div className="mt-auto space-y-2 border-t border-[#27272a] p-4">
              {authUser ? (
                <button
                  type="button"
                  onClick={handleSignOut}
                  title={authUser.email || ''}
                  className="w-full rounded-md border border-[#27272a] bg-[#141417] px-3 py-2 text-[11px] font-semibold text-[#edeeef] hover:border-[#94949e] truncate cursor-pointer"
                >
                  {authUser.email ? `${authUser.email.split('@')[0]} · Đăng xuất` : 'Đăng xuất'}
                </button>
              ) : (
                <button
                  type="button"
                  disabled={authLoading}
                  onClick={handleSignInGoogle}
                  className="w-full rounded-md border border-[#27272a] bg-[#141417] px-3 py-2 text-[11px] font-semibold text-[#edeeef] hover:border-[#22c55e] disabled:opacity-60 cursor-pointer"
                >
                  {authLoading ? 'Đang đăng nhập...' : 'Đăng nhập Google'}
                </button>
              )}

              <button
                type="button"
                onClick={() => setViewMode('home')}
                className="w-full rounded-md border border-transparent px-3 py-1.5 text-[11px] font-medium text-[#94949e] hover:border-[#27272a] hover:text-[#edeeef] cursor-pointer"
              >
                Trang chủ
              </button>
            </div>
          </aside>

          {/* CENTER MAIN PANE */}
          <main className="min-w-0 bg-[#0c0c0e] p-5 lg:overflow-y-auto lg:border-r lg:border-[#27272a]">
            {['all', 'personal', 'company', 'school', 'mendan', 'proposed'].includes(activeTab) && (
              <div className="space-y-4">
                {/* Control Row + Filter Bar */}
                <div className="space-y-3 rounded-lg border border-[#27272a] bg-[#141417] p-3.5">
                  <div className="flex flex-wrap items-center gap-2">
                    <input
                      type="text"
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      placeholder="Quick Search (氏名・FB名・国籍・希望分野・企業名)..."
                      className="min-w-[220px] flex-1 rounded-md border border-[#27272a] bg-[#0c0c0e] px-3.5 py-2 text-xs text-[#edeeef] placeholder:text-[#94949e] focus:border-[#22c55e] focus:outline-none"
                    />

                    <select
                      value={statusFilter}
                      onChange={(e) => setStatusFilter(e.target.value)}
                      className="rounded-md border border-[#27272a] bg-[#0c0c0e] px-3 py-2 text-xs font-semibold text-[#edeeef] focus:border-[#22c55e] focus:outline-none"
                    >
                      <option value="all" className="bg-[#141417]">すべての進捗</option>
                      {PROGRESS_STATUSES.map((st) => (
                        <option key={st} value={st} className="bg-[#141417]">
                          {st}
                        </option>
                      ))}
                    </select>

                    {activeTab === 'proposed' && (
                      <select
                        value={proposedMonthFilter}
                        onChange={(e) => setProposedMonthFilter(e.target.value)}
                        className="rounded-md border border-[#22c55e]/50 bg-[#0c0c0e] px-3 py-2 text-xs font-semibold text-[#22c55e]"
                      >
                        <option value="all" className="bg-[#141417]">提案月（すべて）</option>
                        {proposedMonthOptions.map((m) => (
                          <option key={m} value={m} className="bg-[#141417]">
                            {m}
                          </option>
                        ))}
                      </select>
                    )}

                    <button
                      type="button"
                      onClick={() => {
                        setFbChatMode('new');
                        setFbChatChannel(
                          activeTab === 'company'
                            ? 'company'
                            : activeTab === 'school'
                            ? 'school'
                            : 'personal'
                        );
                        setFbChatRawText('');
                        setFbChatModalOpen(true);
                      }}
                      className="rounded-md border border-[#27272a] bg-transparent px-3 py-2 text-xs font-semibold text-[#edeeef] hover:border-[#94949e] whitespace-nowrap shrink-0 cursor-pointer"
                    >
                      FBチャット登録
                    </button>

                    <button
                      type="button"
                      onClick={() =>
                        setEditingPerson({
                          isSchool: activeTab === 'school',
                          isNew: true,
                          data: {
                            id: `cand_${Date.now()}`,
                            name: '',
                            fbName: '',
                            fbProfileUrl: '',
                            channel:
                              activeTab === 'company'
                                ? 'company'
                                : activeTab === 'school'
                                ? 'school'
                                : 'personal',
                            nationality: '',
                            gender: '',
                            age: '',
                            birthDate: '',
                            jlpt: '要確認',
                            desiredIndustry: '',
                            desiredArea: '',
                            visaStatus: '',
                            visaExpiry: '',
                            passedExam: '',
                            japanExperience: '',
                            conversionGoal: '',
                            salaryExpectation: '',
                            startDate: '',
                            phone: '',
                            email: '',
                            address: '',
                            progressStatus: '新規相談',
                            proposedCompany: '',
                            mendanDate: '',
                            lastContactDate: new Date().toISOString().slice(0, 10),
                            negotiationLog: '',
                            memo: '',
                          },
                        })
                      }
                      className="rounded-md bg-[#22c55e] px-3.5 py-2 text-xs font-semibold text-[#0c0c0e] hover:opacity-90 whitespace-nowrap shrink-0 cursor-pointer"
                    >
                      + New Candidate
                    </button>
                  </div>

                  {/* Quick Care Filter Row */}
                  <div className="flex flex-wrap items-center gap-1.5 border-t border-[#27272a] pt-2.5">
                    {(
                      [
                        { id: 'all', label: 'すべて表示' },
                        { id: 'recent7', label: '最近7日' },
                        { id: 'scheduled', label: '面談予定あり' },
                        { id: 'recentMendan', label: '直近面談順' },
                        { id: 'urgent', label: '要至急フォロー' },
                        { id: 'declined', label: `辞退者 (${declinedPeopleCount})` },
                      ] as Array<{ id: CareFilterType; label: string }>
                    ).map((pill) => (
                      <button
                        key={pill.id}
                        type="button"
                        onClick={() => setCareFilter(pill.id)}
                        className={`rounded px-2.5 py-1 text-[11px] font-semibold transition-colors whitespace-nowrap shrink-0 tabular-nums cursor-pointer ${
                          careFilter === pill.id
                            ? pill.id === 'declined'
                              ? 'bg-[#ef4444] text-white'
                              : 'bg-[#22c55e] text-[#0c0c0e]'
                            : 'border border-[#27272a] bg-[#0c0c0e] text-[#94949e] hover:text-[#edeeef]'
                        }`}
                      >
                        {pill.label}
                      </button>
                    ))}

                    {careFilter === 'declined' && (
                      <input
                        type="text"
                        value={declineReasonQuery}
                        onChange={(e) => setDeclineReasonQuery(e.target.value)}
                        placeholder="辞退理由で絞り込み..."
                        className="ml-1 w-60 rounded border border-[#ef4444]/50 bg-[#0c0c0e] px-2.5 py-1 text-xs text-[#edeeef] placeholder:text-[#94949e] focus:outline-none"
                      />
                    )}
                  </div>

                  {/* Proposed Tab Stage Sub-Filter Bar */}
                  {activeTab === 'proposed' && (
                    <div className="flex flex-wrap items-center gap-1.5 border-t border-[#27272a] pt-2.5">
                      <span className="mr-1 font-mono text-[10px] uppercase tracking-wider text-[#94949e]">
                        Stage:
                      </span>
                      {[
                        { id: 'all', label: `すべて (${proposedStageCounts.all})` },
                        { id: '企業提案中', label: `企業提案中 (${proposedStageCounts['企業提案中']})` },
                        { id: '1次面接予定', label: `1次面接 (${proposedStageCounts['1次面接予定']})` },
                        { id: '2次面接予定', label: `2次面接 (${proposedStageCounts['2次面接予定']})` },
                        { id: '内定', label: `内定 (${proposedStageCounts['内定']})` },
                        { id: '辞退・終了', label: `辞退 (${proposedStageCounts['辞退・終了']})` },
                      ].map((stPill) => (
                        <button
                          key={stPill.id}
                          type="button"
                          onClick={() => setProposedStageFilter(stPill.id)}
                          className={`rounded px-2.5 py-1 text-[11px] font-semibold transition-colors whitespace-nowrap shrink-0 tabular-nums cursor-pointer ${
                            proposedStageFilter === stPill.id
                              ? stPill.id === '辞退・終了'
                                ? 'bg-[#ef4444] text-white'
                                : 'bg-[#22c55e] text-[#0c0c0e]'
                              : 'border border-[#27272a] bg-[#0c0c0e] text-[#94949e] hover:text-[#edeeef]'
                          }`}
                        >
                          {stPill.label}
                        </button>
                      ))}
                    </div>
                  )}
                </div>

                {filteredCandidates.length === 0 ? (
                  <div className="rounded-lg border border-dashed border-[#27272a] bg-[#141417] p-10 text-center text-sm font-semibold text-[#94949e]">
                    条件に一致する相談者データがありません。
                  </div>
                ) : (
                  <div className="space-y-3">{filteredCandidates.map(renderCandidateCard)}</div>
                )}
              </div>
            )}

            {/* KYUJIN MANAGEMENT TAB */}
            {activeTab === 'kyujin' && (
              <div
                className="space-y-4"
                onPaste={(e) => {
                  const img = extractFirstImageFromClipboard(e.clipboardData);
                  if (img) {
                    e.preventDefault();
                    void handleImportKyujinFileList([img]);
                  }
                }}
                onDragOver={(e) => e.preventDefault()}
                onDrop={(e) => {
                  e.preventDefault();
                  const dropped = Array.from(e.dataTransfer.files || []);
                  if (dropped.length > 0) {
                    void handleImportKyujinFileList(dropped);
                  }
                }}
              >
                <div className="space-y-3 rounded-lg border border-[#27272a] bg-[#141417] p-4">
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <div className="flex flex-wrap items-center gap-2">
                      <button
                        type="button"
                        disabled={isImportingKyujin}
                        onClick={() => kyujinFileInputRef.current?.click()}
                        className="inline-flex items-center gap-1.5 rounded-md bg-[#22c55e] px-4 py-2 text-xs font-semibold text-[#0c0c0e] hover:opacity-90 disabled:opacity-60 whitespace-nowrap shrink-0 cursor-pointer"
                      >
                        <Upload className="h-3.5 w-3.5" aria-hidden="true" />
                        {isImportingKyujin
                          ? 'Đang phân tích Kyujin...'
                          : 'Nhập nhiều file (Excel / PDF / Word / CSV / Ảnh)'}
                      </button>

                      <button
                        type="button"
                        disabled={isImportingKyujin}
                        onClick={() => setKyujinRawModalOpen(true)}
                        className="inline-flex items-center gap-1.5 rounded-md border border-[#38bdf8]/50 bg-[#38bdf8]/10 px-3 py-2 text-xs font-semibold text-[#38bdf8] hover:bg-[#38bdf8]/20 whitespace-nowrap shrink-0 cursor-pointer"
                      >
                        <ClipboardPaste className="h-3.5 w-3.5" aria-hidden="true" />
                        Dán văn bản Kyujin / Baito
                      </button>

                      {selectedSlotIds.length > 0 && (
                        <button
                          type="button"
                          onClick={handleDeleteSelectedSlots}
                          className="rounded-md bg-[#ef4444] px-3 py-2 text-xs font-semibold text-white hover:opacity-90 whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          Xóa đã chọn ({selectedSlotIds.length})
                        </button>
                      )}

                      {kyujins.length > 0 && (
                        <button
                          type="button"
                          onClick={handleDeleteAllKyujins}
                          className="rounded-md border border-[#ef4444]/40 bg-transparent px-3 py-2 text-xs font-semibold text-[#ef4444] hover:bg-[#ef4444]/10 whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          Xóa tất cả
                        </button>
                      )}
                    </div>

                    <input
                      type="text"
                      value={kyujinSearch}
                      onChange={(e) => setKyujinSearch(e.target.value)}
                      placeholder="Tìm công ty, khu vực, lĩnh vực, baito, nội dung..."
                      className="w-72 rounded-md border border-[#27272a] bg-[#0c0c0e] px-3 py-1.5 text-xs text-[#edeeef] placeholder:text-[#94949e] focus:border-[#22c55e] focus:outline-none"
                    />
                  </div>

                  {/* Screenshot Drag/Paste hint & progress bar */}
                  <div className="flex flex-wrap items-center justify-between gap-2 rounded-md border border-dashed border-[#27272a] bg-[#0c0c0e] px-3 py-2 text-[11px] text-[#94949e]">
                    <span className="inline-flex items-center gap-1.5">
                      <ImageIcon className="h-3.5 w-3.5 text-[#38bdf8]" aria-hidden="true" />
                      Mẹo: Nhấn <strong className="text-[#edeeef]">Ctrl+V</strong> để dán trực tiếp ảnh chụp màn hình tin tuyển dụng / Baito hoặc kéo thả nhiều file vào đây.
                    </span>
                    {kyujinImportProgress && (
                      <span className="font-mono font-semibold text-[#22c55e] tabular-nums">
                        Đang xử lý {kyujinImportProgress.current}/{kyujinImportProgress.total}: {kyujinImportProgress.label}
                      </span>
                    )}
                  </div>

                  {kyujinImportProgress && (
                    <div className="h-1.5 w-full overflow-hidden rounded-full bg-[#0c0c0e]">
                      <div
                        className="h-full bg-[#22c55e] transition-all"
                        style={{
                          width: `${Math.round(
                            (kyujinImportProgress.current / Math.max(1, kyujinImportProgress.total)) * 100
                          )}%`,
                        }}
                      />
                    </div>
                  )}

                  <div className="flex flex-wrap items-center gap-2 border-t border-[#27272a] pt-3">
                    <div className="flex items-center gap-1 rounded-md border border-[#27272a] bg-[#0c0c0e] p-1">
                      <button
                        type="button"
                        onClick={() => setKyujinStatusFilter('all')}
                        className={`rounded px-2.5 py-1 text-xs font-semibold tabular-nums whitespace-nowrap shrink-0 cursor-pointer ${
                          kyujinStatusFilter === 'all'
                            ? 'bg-[#22c55e] text-[#0c0c0e]'
                            : 'text-[#94949e] hover:text-[#edeeef]'
                        }`}
                      >
                        Tất cả ({kyujinStats.totalSlots})
                      </button>
                      <button
                        type="button"
                        onClick={() => setKyujinStatusFilter('open')}
                        className={`rounded px-2.5 py-1 text-xs font-semibold tabular-nums whitespace-nowrap shrink-0 cursor-pointer ${
                          kyujinStatusFilter === 'open'
                            ? 'bg-[#22c55e] text-[#0c0c0e]'
                            : 'text-[#94949e] hover:text-[#edeeef]'
                        }`}
                      >
                        Đang tuyển ({kyujinStats.openSlots})
                      </button>
                      <button
                        type="button"
                        onClick={() => setKyujinStatusFilter('closed')}
                        className={`rounded px-2.5 py-1 text-xs font-semibold tabular-nums whitespace-nowrap shrink-0 cursor-pointer ${
                          kyujinStatusFilter === 'closed'
                            ? 'bg-[#ef4444] text-white'
                            : 'text-[#94949e] hover:text-[#edeeef]'
                        }`}
                      >
                        Đã đóng ({kyujinStats.closedSlots})
                      </button>
                    </div>

                    <button
                      type="button"
                      onClick={() => setKyujinBonusOnly((prev) => !prev)}
                      className={`rounded-md border px-3 py-1.5 text-xs font-semibold transition-colors tabular-nums whitespace-nowrap shrink-0 cursor-pointer ${
                        kyujinBonusOnly
                          ? 'border-[#f59e0b] bg-[#f59e0b]/15 text-[#f59e0b]'
                          : 'border-[#27272a] bg-[#0c0c0e] text-[#94949e] hover:text-[#edeeef]'
                      }`}
                    >
                      Có thưởng ({kyujinStats.bonusSlots})
                    </button>

                    <select
                      value={kyujinSectorFilter}
                      onChange={(e) => setKyujinSectorFilter(e.target.value)}
                      className="rounded-md border border-[#27272a] bg-[#0c0c0e] px-3 py-1.5 text-xs font-semibold text-[#edeeef]"
                    >
                      <option value="all" className="bg-[#141417]">Lĩnh vực lớn (Tất cả)</option>
                      {SECTORS.map((sec) => (
                        <option key={sec} value={sec} className="bg-[#141417]">
                          {sec}
                        </option>
                      ))}
                    </select>

                    <select
                      value={kyujinIndustryFilter}
                      onChange={(e) => setKyujinIndustryFilter(e.target.value)}
                      className="rounded-md border border-[#27272a] bg-[#0c0c0e] px-3 py-1.5 text-xs font-semibold text-[#edeeef]"
                    >
                      <option value="all" className="bg-[#141417]">Ngành nghề (Tất cả)</option>
                      {INDUSTRIES.map((ind) => (
                        <option key={ind} value={ind} className="bg-[#141417]">
                          {ind}
                        </option>
                      ))}
                    </select>

                    <select
                      value={kyujinEmploymentTypeFilter}
                      onChange={(e) => setKyujinEmploymentTypeFilter(e.target.value)}
                      className="rounded-md border border-[#27272a] bg-[#0c0c0e] px-3 py-1.5 text-xs font-semibold text-[#38bdf8]"
                    >
                      <option value="all" className="bg-[#141417]">
                        Hình thức tuyển dụng (Tất cả · {kyujinStats.baitoSlots} Baito)
                      </option>
                      {EMPLOYMENT_TYPES.map((emp) => (
                        <option key={emp} value={emp} className="bg-[#141417]">
                          {emp}
                        </option>
                      ))}
                    </select>

                    <select
                      value={kyujinFacilityFilter}
                      onChange={(e) => setKyujinFacilityFilter(e.target.value)}
                      className="rounded-md border border-[#27272a] bg-[#0c0c0e] px-3 py-1.5 text-xs font-semibold text-[#edeeef]"
                    >
                      <option value="all" className="bg-[#141417]">Loại hình Kaigo (Tất cả)</option>
                      {KAIGO_FACILITY_TYPES.map((ft) => (
                        <option key={ft} value={ft} className="bg-[#141417]">
                          {ft}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                {filteredKyujinCompanies.length === 0 ? (
                  <div className="rounded-lg border border-dashed border-[#27272a] bg-[#141417] p-10 text-center text-sm font-semibold text-[#94949e]">
                    Chưa có tin tuyển dụng khớp bộ lọc. Bạn có thể tải nhiều file Excel/PDF/Word/CSV/Ảnh, dán văn bản hoặc nhấn Ctrl+V ảnh chụp tin tuyển dụng.
                  </div>
                ) : (
                  filteredKyujinCompanies.map((comp) => (
                    <div
                      key={comp.id}
                      className="rounded-lg border border-[#27272a] bg-[#141417] p-4"
                    >
                      <div className="mb-3 flex flex-wrap items-center justify-between gap-2 border-b border-[#27272a] pb-2.5">
                        <div className="flex flex-wrap items-center gap-2.5">
                          <span className="text-base font-bold text-[#edeeef]">
                            {comp.companyName}
                          </span>
                          <span className="rounded bg-[#27272a] px-2 py-0.5 font-mono text-[10px] text-[#edeeef]">
                            {comp.sector || detectSectorFromIndustryOrText(comp.industry)}
                          </span>
                          <span className="font-mono text-xs font-semibold text-[#22c55e]">
                            {comp.industry}
                          </span>
                          <span className="font-mono text-xs text-[#94949e] tabular-nums">
                            {comp.areaBreakdown.length} vị trí
                          </span>
                        </div>

                        <button
                          type="button"
                          onClick={() => handleDeleteSingleCompany(comp.id, comp.companyName)}
                          className="rounded border border-[#ef4444]/40 bg-transparent px-2.5 py-1 text-xs font-semibold text-[#ef4444] hover:bg-[#ef4444]/10 whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          Xóa công ty
                        </button>
                      </div>

                      <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
                        {comp.areaBreakdown.map((sl) => {
                          const isChecked = selectedSlotIds.includes(sl.slotId);
                          const empType =
                            sl.employmentType || comp.employmentType || 'Tokutei Ginou (特定技能)';
                          const isBaito = /baito|アルバイト|パート/i.test(empType);
                          return (
                            <div
                              key={sl.slotId}
                              className={`rounded-md border p-3 transition-colors ${
                                sl.recruitmentStatus === '募集終了'
                                  ? 'border-[#ef4444]/30 bg-[#ef4444]/5 opacity-80'
                                  : isBaito
                                  ? 'border-[#38bdf8]/40 bg-[#0c0c0e]'
                                  : 'border-[#27272a] bg-[#0c0c0e]'
                              }`}
                            >
                              <div className="mb-2 flex items-start justify-between gap-2">
                                <label className="flex cursor-pointer flex-wrap items-center gap-2">
                                  <input
                                    type="checkbox"
                                    checked={isChecked}
                                    onChange={() => toggleSelectSlot(sl.slotId)}
                                    className="h-4 w-4 rounded border-[#27272a] bg-[#141417] accent-[#22c55e]"
                                  />
                                  <span
                                    className={`rounded px-1.5 py-0.5 font-mono text-[10px] font-bold ${
                                      isBaito
                                        ? 'bg-[#38bdf8]/20 text-[#38bdf8]'
                                        : 'bg-[#22c55e]/15 text-[#22c55e]'
                                    }`}
                                  >
                                    {empType}
                                  </span>
                                  <span className="font-mono text-xs font-bold text-[#edeeef]">
                                    #{sl.jobCode || 'Slot'}
                                  </span>
                                  <span className="text-xs font-semibold text-[#edeeef]/90">
                                    {formatPublicArea(sl.area) || sl.area || 'Nhật Bản'}
                                  </span>
                                  {sl.shiftTime && (
                                    <span className="text-[11px] font-semibold text-sky-400">
                                      · {sl.shiftTime}
                                    </span>
                                  )}
                                  {sl.facilityType && (
                                    <span className="text-[11px] font-semibold text-purple-400">
                                      · {sl.facilityType}
                                    </span>
                                  )}
                                </label>

                                <button
                                  type="button"
                                  onClick={() => setPreviewModal({ company: comp, slot: sl })}
                                  className="rounded bg-[#22c55e] px-2.5 py-1 text-xs font-semibold text-[#0c0c0e] hover:opacity-90 whitespace-nowrap shrink-0 cursor-pointer"
                                >
                                  Xuất ẩn danh
                                </button>
                              </div>

                              {/* Compensation & Experience Strip */}
                              <div className="mb-2 flex flex-wrap items-center gap-2.5 text-xs">
                                {isBaito ? (
                                  <>
                                    <span className="font-mono font-bold text-[#38bdf8] tabular-nums">
                                      時給: {sl.hourlyWage || sl.salary}
                                    </span>
                                    {sl.commuteAllowance && (
                                      <span className="text-[#94949e]">
                                        Đi lại: {sl.commuteAllowance}
                                      </span>
                                    )}
                                    {sl.shiftsPerWeek && (
                                      <span className="text-[#94949e]">
                                        Ca: {sl.shiftsPerWeek}
                                      </span>
                                    )}
                                  </>
                                ) : (
                                  <>
                                    <span className="font-mono font-bold text-[#22c55e] tabular-nums">
                                      Lương: {sl.salary}
                                    </span>
                                    {sl.estimatedNetSalary && (
                                      <span className="font-mono text-[11px] text-[#38bdf8]">
                                        Thực nhận: {sl.estimatedNetSalary}
                                      </span>
                                    )}
                                    {sl.nightShiftAllowance && (
                                      <span className="text-[11px] text-[#94949e]">
                                        Ca đêm: {sl.nightShiftAllowance}
                                      </span>
                                    )}
                                    {sl.bonus === '賞与あり' && (
                                      <span className="font-semibold text-[#f59e0b]">
                                        Thưởng{sl.bonusDetail ? ` (${sl.bonusDetail})` : ''}
                                      </span>
                                    )}
                                  </>
                                )}
                              </div>

                              {sl.experienceRequirement && (
                                <div className="mb-2 flex flex-wrap items-center gap-1.5 text-[11px]">
                                  <span
                                    className={`rounded px-1.5 py-0.5 font-semibold ${
                                      /未経験/.test(sl.experienceRequirement)
                                        ? 'bg-[#22c55e]/15 text-[#22c55e]'
                                        : 'bg-[#f59e0b]/15 text-[#f59e0b]'
                                    }`}
                                  >
                                    {sl.experienceRequirement}
                                  </span>
                                  {sl.experienceDetail && (
                                    <span className="text-[#94949e]">{sl.experienceDetail}</span>
                                  )}
                                </div>
                              )}

                              <div className="mb-2 flex flex-wrap items-center gap-1.5">
                                <select
                                  value={sl.recruitmentStatus}
                                  onChange={(e) =>
                                    handleQuickUpdateSlotField(
                                      comp.id,
                                      sl.slotId,
                                      'recruitmentStatus',
                                      e.target.value
                                    )
                                  }
                                  className={`rounded border px-2 py-1 text-[11px] font-semibold ${
                                    sl.recruitmentStatus === '募集終了'
                                      ? 'border-[#ef4444]/40 bg-[#ef4444]/10 text-[#ef4444]'
                                      : 'border-[#22c55e]/40 bg-[#22c55e]/10 text-[#22c55e]'
                                  }`}
                                >
                                  <option value="募集中" className="bg-[#141417] text-[#edeeef]">Đang tuyển</option>
                                  <option value="募集終了" className="bg-[#141417] text-[#edeeef]">Đã đóng</option>
                                </select>

                                <select
                                  value={empType}
                                  onChange={(e) =>
                                    handleQuickUpdateSlotField(
                                      comp.id,
                                      sl.slotId,
                                      'employmentType',
                                      e.target.value
                                    )
                                  }
                                  className="rounded border border-[#27272a] bg-[#141417] px-2 py-1 text-[11px] font-semibold text-[#38bdf8]"
                                >
                                  {EMPLOYMENT_TYPES.map((emp) => (
                                    <option key={emp} value={emp} className="bg-[#141417]">
                                      {emp}
                                    </option>
                                  ))}
                                </select>

                                <select
                                  value={sl.industry || comp.industry || '特定技能'}
                                  onChange={(e) =>
                                    handleQuickUpdateSlotField(
                                      comp.id,
                                      sl.slotId,
                                      'industry',
                                      e.target.value
                                    )
                                  }
                                  title="Đổi ngành nghề"
                                  className="rounded border border-[#27272a] bg-[#141417] px-2 py-1 text-[11px] font-semibold text-[#edeeef]"
                                >
                                  {INDUSTRIES.map((ind) => (
                                    <option key={ind} value={ind} className="bg-[#141417]">
                                      {ind}
                                    </option>
                                  ))}
                                  <option value="特定技能" className="bg-[#141417]">特定技能</option>
                                </select>

                                <select
                                  value={sl.minJlpt || '要確認'}
                                  onChange={(e) =>
                                    handleQuickUpdateSlotField(
                                      comp.id,
                                      sl.slotId,
                                      'minJlpt',
                                      e.target.value
                                    )
                                  }
                                  className="rounded border border-[#27272a] bg-[#141417] px-2 py-1 text-[11px] font-semibold text-[#edeeef]"
                                >
                                  {JLPT_LEVELS.map((lvl) => (
                                    <option key={lvl} value={lvl} className="bg-[#141417]">
                                      JLPT: {lvl}
                                    </option>
                                  ))}
                                </select>

                                <select
                                  value={sl.housingSupport || '⚠️ 要確認'}
                                  onChange={(e) =>
                                    handleQuickUpdateSlotField(
                                      comp.id,
                                      sl.slotId,
                                      'housingSupport',
                                      e.target.value
                                    )
                                  }
                                  className="rounded border border-[#27272a] bg-[#141417] px-2 py-1 text-[11px] font-semibold text-[#edeeef]"
                                >
                                  <option value="🏠 自社寮あり" className="bg-[#141417]">🏠 KTX công ty</option>
                                  <option value="🏠 借上げ宿舎あり" className="bg-[#141417]">🏠 Nhà thuê hỗ trợ</option>
                                  <option value="🏠 寮・社宅あり" className="bg-[#141417]">🏠 Có KTX/Nhà ở</option>
                                  <option value="❌ 提供無" className="bg-[#141417]">❌ Tự túc nhà</option>
                                  <option value="⚠️ 要確認" className="bg-[#141417]">⚠️ Nhà: Cần xác nhận</option>
                                </select>
                              </div>

                              {sl.jobDescription && (
                                <p className="text-xs leading-relaxed text-[#94949e]">
                                  {sl.jobDescription}
                                </p>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  ))
                )}
              </div>
            )}

            {/* RESIGNATION MANAGEMENT TAB */}
            {activeTab === 'resignation' && (
              <ResignationTab
                records={resignations}
                onSaveRecord={handleSaveResignation}
                onDeleteRecord={handleDeleteResignation}
                onStatusChange={handleResignationStatusChange}
                onBulkImport={handleBulkImportResignations}
                onSyncCalendarSingle={handleSyncResignationCalendarSingle}
                onSyncCalendarBulk={handleBulkSyncResignationsCalendar}
                onConvertToCandidate={handleConvertResignationToCandidate}
                syncingCalendarIds={syncingResignationCalIds}
                isBulkCalendarSyncing={isBulkSyncingResignationsCal}
                onNotify={(msg) => showToast(msg)}
              />
            )}

            {/* PT PARTNERSHIP TAB */}
            {activeTab === 'pt' && (
              <div className="space-y-4">
                <div className="flex flex-wrap items-center justify-between gap-2 rounded-lg border border-[#27272a] bg-[#141417] p-4">
                  <div>
                    <h2 className="text-base font-bold text-[#edeeef]">
                      PT提携・連携管理（パートナー機関・紹介連携）
                    </h2>
                    <p className="text-xs text-[#94949e]">
                      提携PTの連絡先・報酬条件・関連相談者や求人の連携状況を一元管理します。
                    </p>
                  </div>
                  <button
                    type="button"
                    onClick={() =>
                      setPtModal({
                        isNew: true,
                        data: {
                          id: `pt_${Date.now()}`,
                          partnerName: '',
                          contactPerson: '',
                          phone: '',
                          email: '',
                          status: '提携中（稼働）',
                          specialty: '',
                          commissionTerms: '',
                          memo: '',
                        },
                      })
                    }
                    className="rounded-md bg-[#22c55e] px-3.5 py-2 text-xs font-semibold text-[#0c0c0e] hover:opacity-90 whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    ＋ 新規PT提携先を追加
                  </button>
                </div>

                <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
                  {ptPartners.map((pt) => {
                    const relatedPeople = allPeopleCombined.filter((p) =>
                      `${p.memo || ''} ${p.proposedCompany || ''}`.includes(pt.partnerName)
                    );
                    return (
                      <div
                        key={pt.id}
                        className="rounded-lg border border-[#27272a] bg-[#141417] p-4"
                      >
                        <div className="mb-2 flex items-start justify-between gap-2">
                          <div>
                            <div className="flex items-center gap-2">
                              <span className="text-base font-bold text-[#edeeef]">
                                {pt.partnerName}
                              </span>
                              <span className="text-[#27272a]" aria-hidden="true">
                                ·
                              </span>
                              <span className="font-mono text-xs font-semibold text-[#22c55e]">
                                {pt.status}
                              </span>
                            </div>
                            <p className="mt-1 text-xs text-[#94949e]">
                              担当: {pt.contactPerson || '未設定'} / TEL: {pt.phone || '-'} / Email:{' '}
                              {pt.email || '-'}
                            </p>
                          </div>
                          <div className="flex gap-1.5">
                            <button
                              type="button"
                              onClick={() => setPtModal({ isNew: false, data: { ...pt } })}
                              className="rounded border border-[#27272a] px-2.5 py-1 text-xs font-semibold text-[#94949e] hover:border-[#94949e] hover:text-[#edeeef] whitespace-nowrap shrink-0 cursor-pointer"
                            >
                              編集
                            </button>
                            <button
                              type="button"
                              onClick={() =>
                                setConfirmDialog({
                                  title: 'PT提携先の削除',
                                  message: `「${pt.partnerName}」を削除しますか？`,
                                  onConfirm: () => {
                                    setPtPartners((prev) => prev.filter((x) => x.id !== pt.id));
                                    setConfirmDialog(null);
                                    showToast('🗑️ PT提携先を削除しました');
                                  },
                                })
                              }
                              className="rounded border border-[#ef4444]/40 px-2.5 py-1 text-xs font-semibold text-[#ef4444] hover:bg-[#ef4444]/10 whitespace-nowrap shrink-0 cursor-pointer"
                            >
                              削除
                            </button>
                          </div>
                        </div>

                        <div className="space-y-1.5 border-t border-[#27272a] pt-2.5 text-xs text-[#94949e]">
                          <div>
                            <span className="font-semibold text-[#edeeef]">得意分野: </span>
                            {pt.specialty || '全分野対応'}
                          </div>
                          <div>
                            <span className="font-semibold text-[#edeeef]">契約・シェア条件: </span>
                            {pt.commissionTerms || '別途協議'}
                          </div>
                          {pt.memo && <div className="text-[#94949e]">{pt.memo}</div>}
                        </div>

                        <div className="mt-2.5 font-mono text-xs font-semibold text-[#22c55e] tabular-nums">
                          連携中の相談者: {relatedPeople.length} 名
                          {relatedPeople.length > 0 && (
                            <span className="ml-1 font-sans font-normal text-[#94949e]">
                              （{relatedPeople.map((r) => r.name).join('、')}）
                            </span>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}

            {/* ANALYTICS & ENTERPRISE REPORT TAB */}
            {activeTab === 'report' &&
              (() => {
                const totalPeopleCount = allPeopleCombined.length || 1;
                const channelBreakdown = [
                  {
                    label: 'FB個人',
                    count: personalCandidates.length,
                    pct: Math.round((personalCandidates.length / totalPeopleCount) * 100),
                    barClass: 'bg-[#22c55e]',
                  },
                  {
                    label: 'FB会社公式',
                    count: companyCandidates.length,
                    pct: Math.round((companyCandidates.length / totalPeopleCount) * 100),
                    barClass: 'bg-sky-400',
                  },
                  {
                    label: '日本語学校・留学生',
                    count: schoolStudents.length,
                    pct: Math.round((schoolStudents.length / totalPeopleCount) * 100),
                    barClass: 'bg-purple-400',
                  },
                ];

                const statusCountsMap: Record<string, number> = {};
                allPeopleCombined.forEach((p) => {
                  const st = String(p.progressStatus || '新規相談').trim();
                  statusCountsMap[st] = (statusCountsMap[st] || 0) + 1;
                });
                const statusChartRows = Object.entries(statusCountsMap)
                  .map(([status, count]) => ({
                    status,
                    count,
                    pct: Math.round((count / totalPeopleCount) * 100),
                  }))
                  .sort((a, b) => b.count - a.count);

                const industryCountsMap: Record<string, number> = {};
                allPeopleCombined.forEach((p) => {
                  const ind = String(p.desiredIndustry || '未設定').trim() || '未設定';
                  industryCountsMap[ind] = (industryCountsMap[ind] || 0) + 1;
                });
                const industryChartRows = Object.entries(industryCountsMap)
                  .map(([industry, count]) => ({
                    industry,
                    count,
                    pct: Math.round((count / totalPeopleCount) * 100),
                  }))
                  .sort((a, b) => b.count - a.count);

                const jlptOrder = ['N1', 'N2', 'N3', 'N4', 'N5', '日常会話', '要確認'];
                const jlptCountsMap: Record<string, number> = {};
                allPeopleCombined.forEach((p) => {
                  const lv = String(p.jlpt || '要確認').trim() || '要確認';
                  jlptCountsMap[lv] = (jlptCountsMap[lv] || 0) + 1;
                });
                const jlptChartRows = Object.entries(jlptCountsMap)
                  .map(([level, count]) => ({
                    level,
                    count,
                    pct: Math.round((count / totalPeopleCount) * 100),
                  }))
                  .sort((a, b) => {
                    const ia = jlptOrder.indexOf(a.level);
                    const ib = jlptOrder.indexOf(b.level);
                    return (ia === -1 ? 99 : ia) - (ib === -1 ? 99 : ib);
                  });

                const nationalityCountsMap: Record<string, number> = {};
                allPeopleCombined.forEach((p) => {
                  const nat = String(p.nationality || '未設定').trim() || '未設定';
                  nationalityCountsMap[nat] = (nationalityCountsMap[nat] || 0) + 1;
                });
                const nationalityChartRows = Object.entries(nationalityCountsMap)
                  .map(([nat, count]) => ({
                    nat,
                    count,
                    pct: Math.round((count / totalPeopleCount) * 100),
                  }))
                  .sort((a, b) => b.count - a.count);

                const kyujinByIndustryMap: Record<
                  string,
                  { total: number; open: number; bonus: number }
                > = {};
                let housingAvailableSlots = 0;
                kyujins.forEach((c) => {
                  (c.areaBreakdown || []).forEach((sl) => {
                    const ind = String(sl.industry || c.industry || '特定技能').trim();
                    if (!kyujinByIndustryMap[ind]) {
                      kyujinByIndustryMap[ind] = { total: 0, open: 0, bonus: 0 };
                    }
                    kyujinByIndustryMap[ind].total += 1;
                    if (sl.recruitmentStatus !== '募集終了') kyujinByIndustryMap[ind].open += 1;
                    if (sl.bonus === '賞与あり') kyujinByIndustryMap[ind].bonus += 1;
                    if (String(sl.housingSupport || '').includes('あり')) housingAvailableSlots += 1;
                  });
                });
                const kyujinIndustryChartRows = Object.entries(kyujinByIndustryMap)
                  .map(([industry, stats]) => ({ industry, ...stats }))
                  .sort((a, b) => b.total - a.total);
                const totalSlotsSafe = kyujinStats.totalSlots || 1;

                return (
                  <div className="space-y-4">
                    <div className="grid grid-cols-2 gap-3 md:grid-cols-4">
                      <div className="rounded-lg border border-[#27272a] bg-[#141417] p-4">
                        <div className="font-mono text-[10px] uppercase tracking-wider text-[#94949e]">
                          登録人材総数
                        </div>
                        <div className="mt-1 font-mono text-2xl font-bold text-[#edeeef] tabular-nums">
                          {allPeopleCombined.length} 名
                        </div>
                        <div className="mt-1 text-[11px] text-[#94949e] tabular-nums">
                          一般相談 {candidates.length}名 · 留学生 {schoolStudents.length}名
                        </div>
                      </div>
                      <div className="rounded-lg border border-[#27272a] bg-[#141417] p-4">
                        <div className="font-mono text-[10px] uppercase tracking-wider text-[#94949e]">
                          企業提案済み・選考中
                        </div>
                        <div className="mt-1 font-mono text-2xl font-bold text-[#22c55e] tabular-nums">
                          {allProposedPeople.length} 名
                        </div>
                        <div className="mt-1 text-[11px] text-[#94949e] tabular-nums">
                          面談実績あり {mendanPeopleSorted.length}名
                        </div>
                      </div>
                      <div className="rounded-lg border border-[#27272a] bg-[#141417] p-4">
                        <div className="font-mono text-[10px] uppercase tracking-wider text-[#94949e]">
                          保有求人枠（募集中）
                        </div>
                        <div className="mt-1 font-mono text-2xl font-bold text-[#f59e0b] tabular-nums">
                          {kyujinStats.openSlots} 枠
                        </div>
                        <div className="mt-1 text-[11px] text-[#94949e] tabular-nums">
                          全 {kyujinStats.totalCompanies} 社 · 賞与あり {kyujinStats.bonusSlots}枠
                        </div>
                      </div>
                      <div className="rounded-lg border border-[#27272a] bg-[#141417] p-4">
                        <div className="font-mono text-[10px] uppercase tracking-wider text-[#94949e]">
                          提携PT機関数
                        </div>
                        <div className="mt-1 font-mono text-2xl font-bold text-[#edeeef] tabular-nums">
                          {ptPartners.length} 社
                        </div>
                        <div className="mt-1 text-[11px] text-[#94949e]">継続連携・案件共有中</div>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
                      <div className="rounded-lg border border-[#27272a] bg-[#141417] p-4">
                        <div className="mb-3 flex items-center justify-between">
                          <h3 className="text-sm font-bold text-[#edeeef]">
                            チャネル別・国籍別 登録人材構成（計 {allPeopleCombined.length} 名）
                          </h3>
                          <span className="font-mono text-[10px] uppercase tracking-wider text-[#22c55e]">
                            Summary
                          </span>
                        </div>

                        <div className="mb-3 flex h-3 w-full overflow-hidden rounded-full bg-[#0c0c0e]">
                          {channelBreakdown.map((ch) =>
                            ch.count > 0 ? (
                              <div
                                key={ch.label}
                                style={{ width: `${Math.max(4, ch.pct)}%` }}
                                className={`${ch.barClass} transition-all`}
                                title={`${ch.label}: ${ch.count}名 (${ch.pct}%)`}
                              />
                            ) : null
                          )}
                        </div>

                        <div className="space-y-2.5">
                          {channelBreakdown.map((ch) => (
                            <div key={ch.label} className="space-y-1">
                              <div className="flex items-center justify-between text-xs">
                                <span className="font-semibold text-[#edeeef]">{ch.label}</span>
                                <span className="font-mono font-bold text-[#edeeef] tabular-nums">
                                  {ch.count} 名 <span className="text-[#94949e]">({ch.pct}%)</span>
                                </span>
                              </div>
                              <div className="h-2 w-full overflow-hidden rounded-full bg-[#0c0c0e]">
                                <div
                                  style={{ width: `${ch.pct}%` }}
                                  className={`h-full rounded-full ${ch.barClass}`}
                                />
                              </div>
                            </div>
                          ))}
                        </div>

                        <div className="mt-4 border-t border-[#27272a] pt-3">
                          <div className="mb-2 font-mono text-[10px] uppercase tracking-wider text-[#94949e]">
                            国籍別内訳
                          </div>
                          <div className="flex flex-wrap gap-2 text-xs">
                            {nationalityChartRows.map((item) => (
                              <div
                                key={item.nat}
                                className="flex items-center gap-1.5 rounded border border-[#27272a] bg-[#0c0c0e] px-2.5 py-1"
                              >
                                <span className="font-semibold text-[#edeeef]">{item.nat}</span>
                                <span className="font-mono font-bold text-[#22c55e] tabular-nums">
                                  {item.count}名 ({item.pct}%)
                                </span>
                              </div>
                            ))}
                          </div>
                        </div>
                      </div>

                      <div className="rounded-lg border border-[#27272a] bg-[#141417] p-4">
                        <div className="mb-3 flex items-center justify-between">
                          <h3 className="text-sm font-bold text-[#edeeef]">
                            進捗ステータス別 選考パイプライン
                          </h3>
                          <span className="font-mono text-[10px] text-[#94949e]">人数・構成比</span>
                        </div>
                        <div className="space-y-2">
                          {statusChartRows.map((row) => (
                            <div key={row.status} className="space-y-1">
                              <div className="flex items-center justify-between text-xs">
                                <span className="font-semibold text-[#edeeef]">{row.status}</span>
                                <span className="font-mono font-bold text-[#22c55e] tabular-nums">
                                  {row.count} 名 ({row.pct}%)
                                </span>
                              </div>
                              <div className="h-2 w-full overflow-hidden rounded-full bg-[#0c0c0e]">
                                <div
                                  style={{ width: `${Math.max(3, row.pct)}%` }}
                                  className={`h-full rounded-full ${
                                    row.status.includes('内定') || row.status.includes('入社')
                                      ? 'bg-[#22c55e]'
                                      : row.status.includes('面接') || row.status.includes('提案')
                                      ? 'bg-sky-400'
                                      : row.status.includes('面談')
                                      ? 'bg-teal-400'
                                      : row.status.includes('保留') || row.status.includes('辞退')
                                      ? 'bg-[#ef4444]'
                                      : 'bg-[#94949e]'
                                  }`}
                                />
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
                      <div className="space-y-4 rounded-lg border border-[#27272a] bg-[#141417] p-4">
                        <div>
                          <h3 className="mb-2.5 text-sm font-bold text-[#edeeef]">
                            希望分野別 人材分布
                          </h3>
                          <div className="space-y-2">
                            {industryChartRows.map((row) => (
                              <div key={row.industry} className="space-y-1">
                                <div className="flex items-center justify-between text-xs">
                                  <span className="font-semibold text-[#edeeef]">{row.industry}</span>
                                  <span className="font-mono font-bold text-[#22c55e] tabular-nums">
                                    {row.count} 名 ({row.pct}%)
                                  </span>
                                </div>
                                <div className="h-2 w-full overflow-hidden rounded-full bg-[#0c0c0e]">
                                  <div
                                    style={{ width: `${Math.max(3, row.pct)}%` }}
                                    className="h-full rounded-full bg-[#22c55e]"
                                  />
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>

                        <div className="border-t border-[#27272a] pt-3">
                          <h3 className="mb-2.5 text-sm font-bold text-[#edeeef]">
                            日本語レベル（JLPT）別 人材分布
                          </h3>
                          <div className="grid grid-cols-2 gap-2 sm:grid-cols-4">
                            {jlptChartRows.map((row) => (
                              <div
                                key={row.level}
                                className="rounded border border-[#27272a] bg-[#0c0c0e] p-2.5 text-center"
                              >
                                <div className="font-mono text-xs font-bold text-[#22c55e]">
                                  {row.level}
                                </div>
                                <div className="mt-0.5 font-mono text-lg font-bold text-[#edeeef] tabular-nums">
                                  {row.count}名
                                </div>
                                <div className="font-mono text-[10px] font-semibold text-[#94949e] tabular-nums">
                                  {row.pct}%
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>
                      </div>

                      <div className="space-y-4 rounded-lg border border-[#27272a] bg-[#141417] p-4">
                        <div className="flex items-center justify-between">
                          <h3 className="text-sm font-bold text-[#edeeef]">
                            求人票ステータス・賞与・住宅サポート分析
                          </h3>
                          <span className="font-mono text-[11px] font-semibold text-[#94949e] tabular-nums">
                            全 {kyujinStats.totalCompanies}社 / {kyujinStats.totalSlots}枠
                          </span>
                        </div>

                        <div className="grid grid-cols-3 gap-2 text-center">
                          <div className="rounded border border-[#22c55e]/30 bg-[#22c55e]/10 p-2.5">
                            <div className="text-[11px] font-semibold text-[#22c55e]">募集中</div>
                            <div className="mt-0.5 font-mono text-lg font-bold text-[#edeeef] tabular-nums">
                              {kyujinStats.openSlots} 枠
                            </div>
                            <div className="font-mono text-[10px] text-[#22c55e] tabular-nums">
                              {Math.round((kyujinStats.openSlots / totalSlotsSafe) * 100)}%
                            </div>
                          </div>
                          <div className="rounded border border-[#f59e0b]/30 bg-[#f59e0b]/10 p-2.5">
                            <div className="text-[11px] font-semibold text-[#f59e0b]">賞与あり</div>
                            <div className="mt-0.5 font-mono text-lg font-bold text-[#edeeef] tabular-nums">
                              {kyujinStats.bonusSlots} 枠
                            </div>
                            <div className="font-mono text-[10px] text-[#f59e0b] tabular-nums">
                              {Math.round((kyujinStats.bonusSlots / totalSlotsSafe) * 100)}%
                            </div>
                          </div>
                          <div className="rounded border border-sky-400/30 bg-sky-400/10 p-2.5">
                            <div className="text-[11px] font-semibold text-sky-400">寮・社宅あり</div>
                            <div className="mt-0.5 font-mono text-lg font-bold text-[#edeeef] tabular-nums">
                              {housingAvailableSlots} 枠
                            </div>
                            <div className="font-mono text-[10px] text-sky-400 tabular-nums">
                              {Math.round((housingAvailableSlots / totalSlotsSafe) * 100)}%
                            </div>
                          </div>
                        </div>

                        <div>
                          <div className="mb-2 text-xs font-bold text-[#94949e]">
                            分野別 求人枠数（募集中 / 賞与あり内訳）
                          </div>
                          <div className="space-y-2.5">
                            {kyujinIndustryChartRows.map((row) => {
                              const pct = Math.round((row.total / totalSlotsSafe) * 100);
                              return (
                                <div key={row.industry} className="space-y-1">
                                  <div className="flex flex-wrap items-center justify-between gap-1 text-xs">
                                    <span className="font-semibold text-[#edeeef]">{row.industry}</span>
                                    <span className="font-mono font-bold text-[#edeeef] tabular-nums">
                                      計 {row.total}枠（募集 {row.open} / 賞与 {row.bonus}）
                                    </span>
                                  </div>
                                  <div className="h-2 w-full overflow-hidden rounded-full bg-[#0c0c0e]">
                                    <div
                                      style={{ width: `${Math.max(4, pct)}%` }}
                                      className="h-full rounded-full bg-[#f59e0b]"
                                    />
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      </div>
                    </div>

                    {/* Facebook Messenger Webhook Panel */}
                    <div className="rounded-lg border border-[#27272a] bg-[#141417] p-4">
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <div className="flex items-center gap-2">
                          <ShieldCheck className="h-4 w-4 text-[#22c55e]" aria-hidden="true" />
                          <h3 className="text-sm font-bold text-[#edeeef]">
                            Facebook Messenger Webhook (Meta Developers)
                          </h3>
                        </div>
                        <span className="rounded border border-[#22c55e]/40 bg-[#22c55e]/20 px-2 py-0.5 font-mono text-[10px] font-bold text-[#22c55e]">
                          Sẵn sàng
                        </span>
                      </div>

                      <p className="mt-1.5 text-xs text-[#94949e]">
                        Tin nhắn mới từ ứng viên qua Fanpage sẽ được bóc tách tự động bằng Gemini AI và lưu vào mục <strong className="text-[#edeeef]">FB会社公式</strong>.
                      </p>

                      <div className="mt-3 flex flex-wrap items-center gap-2">
                        <input
                          type="text"
                          readOnly
                          value={liveMessengerWebhookUrl}
                          aria-label="Webhook URL"
                          className="min-w-[260px] flex-1 rounded-md border border-[#27272a] bg-[#0c0c0e] px-3 py-2 font-mono text-xs text-[#edeeef]"
                        />
                        <button
                          type="button"
                          onClick={handleCopyWebhookUrl}
                          className="inline-flex items-center gap-1.5 rounded-md bg-[#22c55e] px-4 py-2 text-xs font-semibold text-[#0c0c0e] hover:opacity-90 whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          <Copy className="h-3.5 w-3.5" aria-hidden="true" />
                          Sao chép Webhook URL
                        </button>
                        <button
                          type="button"
                          onClick={openMessengerLinkModal}
                          className="inline-flex items-center gap-1.5 rounded-md border border-[#38bdf8]/50 bg-[#38bdf8]/10 px-3 py-2 text-xs font-semibold text-[#38bdf8] hover:bg-[#38bdf8]/20 whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          <Link2 className="h-3.5 w-3.5" aria-hidden="true" />
                          Liên kết Messenger FB
                        </button>
                        <button
                          type="button"
                          disabled={isRefreshingSql}
                          onClick={() => void refreshCrmFromSql(false)}
                          className="inline-flex items-center gap-1.5 rounded-md border border-[#27272a] bg-[#0c0c0e] px-3 py-2 text-xs font-semibold text-[#edeeef] hover:border-[#94949e] disabled:opacity-50 whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          <RefreshCw
                            className={`h-3.5 w-3.5 ${isRefreshingSql ? 'animate-spin' : ''}`}
                            aria-hidden="true"
                          />
                          Đồng bộ Messenger
                        </button>
                      </div>

                      <div className="mt-3 flex flex-wrap gap-3 font-mono text-[10px] text-[#94949e]">
                        <span>
                          Verify Token:{' '}
                          <strong
                            className={
                              webhookStatus?.hasVerifyToken ? 'text-[#22c55e]' : 'text-[#38bdf8]'
                            }
                          >
                            {webhookStatus?.hasVerifyToken ? 'Đã đặt' : 'Tùy chọn (Mở)'}
                          </strong>
                        </span>
                        <span>·</span>
                        <span>
                          App Secret (HMAC):{' '}
                          <strong
                            className={
                              webhookStatus?.hasAppSecret ? 'text-[#22c55e]' : 'text-[#38bdf8]'
                            }
                          >
                            {webhookStatus?.hasAppSecret ? 'Đã bảo vệ' : 'Tùy chọn (Mở)'}
                          </strong>
                        </span>
                        <span>·</span>
                        <span>
                          Tài khoản:{' '}
                          <strong className="text-[#22c55e]">
                            {authUser?.email || 'Đã đăng nhập'}
                          </strong>
                        </span>
                      </div>
                    </div>

                    <div className="rounded-lg border border-[#27272a] bg-[#141417] p-4">
                      <h3 className="text-sm font-bold text-[#edeeef]">
                        Google Apps Script (GAS) 連携設定
                      </h3>
                      <p className="mt-1 text-xs text-[#94949e]">
                        誤上書き防止のため初期状態は「GAS上書き保護中」になっています。同期が必要な場合のみヘッダーの保護ボタンを解除して送信してください。
                      </p>
                      <div className="mt-3 flex flex-wrap items-center gap-2">
                        <input
                          type="text"
                          value={gasWebAppUrl}
                          onChange={(e) => setGasWebAppUrl(e.target.value)}
                          placeholder="https://script.google.com/macros/s/..."
                          className="min-w-[280px] flex-1 rounded-md border border-[#27272a] bg-[#0c0c0e] px-3 py-2 text-xs text-[#edeeef] placeholder:text-[#94949e] focus:border-[#22c55e] focus:outline-none"
                        />
                        <button
                          type="button"
                          disabled={gasWriteLocked || isSyncingGas}
                          onClick={handleSyncToGas}
                          className="rounded-md bg-[#22c55e] px-4 py-2 text-xs font-semibold text-[#0c0c0e] disabled:opacity-40 whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          {isSyncingGas ? '送信中...' : 'GASへ同期送信'}
                        </button>
                      </div>
                    </div>
                  </div>
                );
              })()}
          </main>

          {/* RIGHT SYSTEM ALERTS PANE */}
          <aside className="border-t border-[#27272a] bg-[#141417] p-5 lg:overflow-y-auto lg:border-t-0">
            <div className="mb-4 font-mono text-[10px] uppercase tracking-widest text-[#94949e]">
              System Alerts
            </div>

            {/* Messenger Webhook Quick Status Card */}
            <div className="mb-4 rounded-lg border border-[#27272a] bg-[#0c0c0e] p-3.5">
              <div className="flex items-center justify-between gap-2">
                <span className="font-mono text-[10px] uppercase tracking-wider text-[#94949e]">
                  Messenger Webhook
                </span>
                <span className="rounded bg-[#22c55e]/20 px-1.5 py-0.5 font-mono text-[10px] font-bold text-[#22c55e]">
                  Sẵn sàng
                </span>
              </div>
              <div className="mt-2 truncate font-mono text-[10px] text-[#94949e]" title={liveMessengerWebhookUrl}>
                {liveMessengerWebhookUrl}
              </div>
              <div className="mt-2.5 flex flex-col gap-1.5">
                <button
                  type="button"
                  onClick={openMessengerLinkModal}
                  className="inline-flex w-full items-center justify-center gap-1.5 rounded bg-[#38bdf8] px-2.5 py-1.5 text-[11px] font-semibold text-[#0c0c0e] hover:opacity-90 whitespace-nowrap cursor-pointer"
                >
                  <Link2 className="h-3 w-3" aria-hidden="true" />
                  Liên kết Messenger FB
                </button>
                <div className="flex gap-1.5">
                  <button
                    type="button"
                    onClick={handleCopyWebhookUrl}
                    className="inline-flex flex-1 items-center justify-center gap-1 rounded border border-[#27272a] bg-[#141417] px-2.5 py-1 text-[10px] font-semibold text-[#edeeef] hover:border-[#22c55e] whitespace-nowrap cursor-pointer"
                  >
                    <Copy className="h-3 w-3" aria-hidden="true" />
                    Sao chép URL
                  </button>
                  {fbPageLink && (
                    <a
                      href={fbPageLink}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="inline-flex items-center gap-1 rounded border border-[#27272a] bg-[#141417] px-2.5 py-1 text-[10px] font-semibold text-[#38bdf8] hover:border-[#38bdf8] whitespace-nowrap"
                    >
                      <ExternalLink className="h-3 w-3" aria-hidden="true" />
                      Mở Page
                    </a>
                  )}
                </div>
              </div>
            </div>

            {/* Alert Card 1: Upcoming PV / Mendan + Google Calendar Sync */}
            <div className="mb-4 space-y-3">
              <div className="flex items-center justify-between gap-2">
                <span className="font-mono text-[10px] uppercase tracking-wider text-[#94949e]">
                  Next PV / Mendan
                </span>
                <button
                  type="button"
                  disabled={isBulkSyncingCalendar}
                  onClick={requestBulkSyncUpcomingCalendar}
                  className="inline-flex items-center gap-1 rounded border border-[#22c55e]/50 bg-[#22c55e]/10 px-2 py-1 text-[10px] font-semibold text-[#22c55e] hover:bg-[#22c55e]/20 disabled:opacity-50 whitespace-nowrap shrink-0 cursor-pointer"
                  title={
                    hasCalendarToken
                      ? 'Đẩy toàn bộ lịch Mendan & Phỏng vấn lên Google Calendar'
                      : 'Kết nối & đồng bộ toàn bộ lịch sắp tới lên Google Calendar'
                  }
                >
                  <Calendar className="h-3 w-3" aria-hidden="true" />
                  {isBulkSyncingCalendar
                    ? 'Đang đồng bộ...'
                    : 'Đồng bộ toàn bộ lịch sắp tới'}
                </button>
              </div>

              {upcomingMendanAlerts.length > 0 ? (
                upcomingMendanAlerts.slice(0, 3).map(({ person, mendan }) => {
                  const isSchool = Boolean(person._isSchool || person.channel === 'school');
                  const hasSyncedCal = Boolean(person.googleCalendarEventId);
                  return (
                    <div
                      key={`mendan_${person.id}`}
                      className="rounded-lg border border-[#27272a] border-l-[3px] border-l-[#22c55e] bg-[#0c0c0e] p-3.5"
                    >
                      <div className="mb-1.5 flex items-center justify-between gap-2">
                        <span className="font-mono text-[10px] uppercase tracking-wider text-[#94949e]">
                          Next PV / Mendan
                        </span>
                        {hasSyncedCal && (
                          <span className="rounded bg-[#22c55e]/15 px-1.5 py-0.5 font-mono text-[9px] font-bold text-[#22c55e]">
                            Calendar OK
                          </span>
                        )}
                      </div>
                      <div className="flex items-center gap-2 text-xs font-semibold text-[#edeeef]">
                        <span className="inline-block h-2 w-2 rounded-full bg-[#22c55e]" />
                        <span className="truncate">{person.name}</span>
                      </div>
                      <div className="mt-1 font-mono text-[11px] text-[#94949e] tabular-nums">
                        {person.interviewDate || mendan.dateLabel || person.progressStatus}
                      </div>
                      <div className="mt-3 flex flex-wrap gap-1.5">
                        <button
                          type="button"
                          onClick={() =>
                            setReplyModal({ person, scenario: 'remind_interview' })
                          }
                          className="flex-1 rounded bg-[#22c55e] px-2.5 py-1.5 text-[11px] font-semibold text-[#0c0c0e] hover:opacity-90 whitespace-nowrap cursor-pointer"
                        >
                          Nhắn lịch
                        </button>
                        <button
                          type="button"
                          onClick={() => handleMarkContactedToday(person, isSchool)}
                          className="flex-1 rounded border border-[#27272a] bg-transparent px-2.5 py-1.5 text-[11px] font-semibold text-[#94949e] hover:border-[#94949e] hover:text-[#edeeef] whitespace-nowrap cursor-pointer"
                        >
                          連絡済
                        </button>
                      </div>

                      <div className="mt-2 flex items-center gap-1.5 border-t border-dashed border-[#27272a] pt-2">
                        <button
                          type="button"
                          disabled={syncingCalendarId === person.id}
                          onClick={() => requestSyncCandidateCalendar(person)}
                          className={`inline-flex flex-1 items-center justify-center gap-1 rounded border px-2 py-1 text-[10px] font-semibold transition-colors whitespace-nowrap cursor-pointer disabled:opacity-50 ${
                            hasSyncedCal
                              ? 'border-[#22c55e]/40 bg-[#22c55e]/10 text-[#22c55e]'
                              : 'border-[#27272a] bg-[#141417] text-[#edeeef] hover:border-[#22c55e]'
                          }`}
                        >
                          <Calendar className="h-3 w-3" aria-hidden="true" />
                          {syncingCalendarId === person.id
                            ? 'Đang đồng bộ...'
                            : hasSyncedCal
                            ? 'Cập nhật Google Calendar'
                            : 'Đồng bộ Google Calendar'}
                        </button>
                        {hasSyncedCal &&
                          isValidGoogleCalendarUrl(person.googleCalendarEventLink) && (
                            <a
                              href={person.googleCalendarEventLink}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="inline-flex items-center gap-1 rounded border border-[#27272a] bg-[#141417] px-2 py-1 text-[10px] font-semibold text-[#94949e] hover:border-[#22c55e] hover:text-[#22c55e] whitespace-nowrap"
                              title="Mở sự kiện trên Google Calendar"
                            >
                              <ExternalLink className="h-3 w-3" aria-hidden="true" />
                              Mở
                            </a>
                          )}
                      </div>
                    </div>
                  );
                })
              ) : (
                <div className="rounded-lg border border-[#27272a] bg-[#0c0c0e] p-3.5">
                  <div className="font-mono text-[10px] uppercase tracking-wider text-[#94949e]">
                    Next PV / Mendan
                  </div>
                  <p className="mt-1.5 text-xs text-[#94949e]">
                    直近の面談・面接予定はありません。
                  </p>
                </div>
              )}
            </div>

            {/* Alert Card 2: Critical Follow-Up */}
            {urgentFollowUpList.length > 0 && (
              <div className="mb-4 rounded-lg border border-[#27272a] border-l-[3px] border-l-[#ef4444] bg-[rgba(239,68,68,0.06)] p-3.5">
                <div className="mb-2 font-mono text-[10px] uppercase tracking-wider text-[#ef4444]">
                  Critical Follow-up ({urgentFollowUpList.length})
                </div>
                <div className="space-y-2 text-xs">
                  {urgentFollowUpList.slice(0, 4).map(({ person, reason }) => (
                    <div
                      key={`urg_${person.id}`}
                      className="flex items-center justify-between gap-2"
                    >
                      <button
                        type="button"
                        onClick={() =>
                          setReplyModal({ person, scenario: 'followup_silent' })
                        }
                        className="truncate text-left font-medium text-[#edeeef] hover:text-[#ef4444] cursor-pointer"
                      >
                        <strong>{person.name}</strong>{' '}
                        <span className="text-[#94949e]">({reason})</span>
                      </button>
                    </div>
                  ))}
                </div>
                <button
                  type="button"
                  onClick={() => {
                    setActiveTab('all');
                    setCareFilter((prev) => (prev === 'urgent' ? 'all' : 'urgent'));
                  }}
                  className="mt-3 text-[11px] font-semibold text-[#ef4444] underline cursor-pointer"
                >
                  {careFilter === 'urgent'
                    ? '絞り込みを解除 →'
                    : `全${urgentFollowUpList.length}名を表示 →`}
                </button>
              </div>
            )}

            {/* Daily Focus Checklist */}
            <div className="mb-4 rounded-lg border border-[#27272a] bg-[#0c0c0e] p-3.5">
              <div className="mb-2 flex items-center justify-between gap-2">
                <span className="font-mono text-[10px] uppercase tracking-wider text-[#22c55e]">
                  Việc cần làm hôm nay (Daily Focus)
                </span>
                <span className="font-mono text-[10px] text-[#94949e] tabular-nums">
                  {dailyFocusItems.filter((i) => dailyFocusCompleted[i.id]).length}/
                  {dailyFocusItems.length}
                </span>
              </div>
              {dailyFocusItems.length === 0 ? (
                <p className="text-xs text-[#94949e]">
                  Tuyệt vời! Hiện không có việc quá hạn hay gấp trong hôm nay.
                </p>
              ) : (
                <div className="max-h-60 space-y-1.5 overflow-y-auto pr-1">
                  {dailyFocusItems.map((item) => {
                    const checked = Boolean(dailyFocusCompleted[item.id]);
                    return (
                      <div
                        key={item.id}
                        className={`flex items-start gap-2 rounded border px-2.5 py-2 text-xs transition-colors ${
                          checked
                            ? 'border-[#27272a] bg-[#141417]/50 opacity-55'
                            : item.urgency === 'critical'
                            ? 'border-[#ef4444]/40 bg-[#ef4444]/5'
                            : 'border-[#27272a] bg-[#141417]'
                        }`}
                      >
                        <button
                          type="button"
                          onClick={() => toggleDailyFocusItem(item.id)}
                          aria-label={`Đánh dấu hoàn thành ${item.title}`}
                          className="mt-0.5 shrink-0 text-[#22c55e] cursor-pointer"
                        >
                          {checked ? (
                            <CheckSquare className="h-4 w-4" aria-hidden="true" />
                          ) : (
                            <Square className="h-4 w-4 text-[#94949e]" aria-hidden="true" />
                          )}
                        </button>
                        <button
                          type="button"
                          onClick={() => setActiveTab(item.targetTab)}
                          className="min-w-0 flex-1 text-left cursor-pointer"
                        >
                          <div
                            className={`font-semibold ${
                              checked
                                ? 'line-through text-[#94949e]'
                                : 'text-[#edeeef] hover:text-[#22c55e]'
                            }`}
                          >
                            {item.title}
                          </div>
                          <div className="truncate text-[10px] text-[#94949e]">
                            {item.subtitle}
                          </div>
                        </button>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>

            {/* Resignation Paperwork Deadline Alerts (<= 7 days / <= 3 days) */}
            {resignationAlerts.length > 0 && (
              <div className="mb-4 space-y-2.5">
                <div className="flex items-center justify-between gap-2">
                  <span className="font-mono text-[10px] uppercase tracking-wider text-[#f59e0b]">
                    Hạn giấy tờ nghỉ việc ({resignationAlerts.length})
                  </span>
                  <button
                    type="button"
                    onClick={() => setActiveTab('resignation')}
                    className="text-[10px] font-semibold text-[#38bdf8] hover:underline cursor-pointer"
                  >
                    Xem tất cả &rarr;
                  </button>
                </div>
                {resignationAlerts.slice(0, 3).map(({ record, assessment }) => {
                  const isCrit =
                    assessment.urgency === 'overdue' || assessment.urgency === 'urgent3';
                  return (
                    <div
                      key={`resig_alert_${record.id}`}
                      className={`rounded-lg border border-[#27272a] border-l-[3px] bg-[#0c0c0e] p-3 text-xs ${
                        isCrit ? 'border-l-[#ef4444]' : 'border-l-[#f59e0b]'
                      }`}
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span className="font-semibold text-[#edeeef] truncate">
                          {record.fullName}
                        </span>
                        <span
                          className={`rounded px-1.5 py-0.5 font-mono text-[9px] font-bold ${
                            isCrit
                              ? 'bg-[#ef4444]/20 text-[#ef4444]'
                              : 'bg-[#f59e0b]/20 text-[#f59e0b]'
                          }`}
                        >
                          {assessment.label}
                        </span>
                      </div>
                      <div className="mt-1 text-[11px] text-[#94949e] truncate">
                        Cty cũ: {record.previousCompany}
                      </div>
                      <div className="mt-2.5 flex flex-wrap gap-1.5">
                        <button
                          type="button"
                          onClick={() => setResignationReminderModal(record)}
                          className="flex-1 rounded border border-[#f59e0b]/50 bg-[#f59e0b]/10 px-2 py-1 text-[10px] font-semibold text-[#f59e0b] hover:bg-[#f59e0b]/20 whitespace-nowrap cursor-pointer"
                        >
                          Tạo tin nhắn nhắc
                        </button>
                        <button
                          type="button"
                          onClick={() => void handleSyncResignationCalendarSingle(record)}
                          className="rounded border border-[#22c55e]/50 bg-[#22c55e]/10 px-2 py-1 text-[10px] font-semibold text-[#22c55e] hover:bg-[#22c55e]/20 whitespace-nowrap cursor-pointer"
                          title="Đồng bộ lên Google Calendar"
                        >
                          <Calendar className="h-3 w-3" aria-hidden="true" />
                        </button>
                        <button
                          type="button"
                          onClick={() => handleConvertResignationToCandidate(record)}
                          className="rounded border border-[#38bdf8]/50 bg-[#38bdf8]/10 px-2 py-1 text-[10px] font-semibold text-[#38bdf8] hover:bg-[#38bdf8]/20 whitespace-nowrap cursor-pointer"
                          title="Chuyển sang Ứng viên tìm việc mới"
                        >
                          <UserPlus className="h-3 w-3" aria-hidden="true" />
                        </button>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}

            {/* Alert Card 3: Incomplete Profiles */}
            {incompleteProfileAlerts.length > 0 && (
              <div className="space-y-3">
                {incompleteProfileAlerts.slice(0, 2).map(({ person, missing }) => (
                  <div
                    key={`inc_${person.id}`}
                    className="rounded-lg border border-dashed border-[#f59e0b]/70 bg-[#0c0c0e] p-3.5"
                  >
                    <div className="mb-1.5 font-mono text-[10px] uppercase tracking-wider text-[#f59e0b]">
                      Incomplete Profile
                    </div>
                    <div className="text-xs font-semibold text-[#edeeef]">
                      {person.name}
                    </div>
                    <div className="mt-1 font-mono text-[10px] text-[#94949e]">
                      Need: {missing.map((m) => m.vi).join(', ')}
                    </div>
                    <button
                      type="button"
                      onClick={() => setReplyModal({ person, scenario: 'missing_info_request' })}
                      className="mt-3 w-full rounded border border-[#f59e0b] px-2.5 py-1.5 text-[11px] font-semibold text-[#f59e0b] hover:bg-[#f59e0b]/15 cursor-pointer"
                    >
                      Xin thông tin thiếu
                    </button>
                  </div>
                ))}
              </div>
            )}
          </aside>
        </div>
      )}

      {/* MODAL 1: EDIT / CREATE CANDIDATE */}
      {editingPerson && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-black/60 p-4">
          <div
            className="max-h-[90vh] w-full max-w-3xl overflow-y-auto rounded-xl bg-white p-5 text-xs text-slate-900 shadow-2xl"
            onPaste={(e) => {
              const img = extractFirstImageFromClipboard(e.clipboardData);
              if (img) {
                e.preventDefault();
                void handleExtractCandidateScreenshot(img, 'edit');
              }
            }}
          >
            <div className="mb-3 flex items-center justify-between border-b border-slate-200 pb-2.5">
              <h3 className="text-base font-bold text-slate-900">
                {editingPerson.isNew
                  ? '＋ 新規相談者登録'
                  : `プロフィール編集：${editingPerson.data.name || editingPerson.data.fbName || ''}`}
              </h3>
              <button
                type="button"
                onClick={() => setEditingPerson(null)}
                className="rounded-lg bg-slate-100 px-3 py-1 font-semibold text-slate-700"
              >
                閉じる
              </button>
            </div>

            <div className="mb-4 rounded-lg border border-sky-200 bg-sky-50/70 p-3">
              <div className="mb-1.5 flex flex-wrap items-center justify-between gap-2">
                <span className="font-bold text-sky-950">
                  Dán đoạn chat / ghi chú dài hoặc Ctrl+V ảnh chụp Messenger / Thẻ ngoại kiều (Bảo toàn 100% nội dung)
                </span>
                <div className="flex flex-wrap items-center gap-1.5">
                  <label className="inline-flex cursor-pointer items-center gap-1 rounded-md border border-sky-300 bg-white px-2.5 py-1 text-[11px] font-semibold text-sky-800 hover:bg-sky-50">
                    <ImageIcon className="h-3.5 w-3.5" aria-hidden="true" />
                    Tải ảnh chụp
                    <input
                      type="file"
                      accept="image/*"
                      className="hidden"
                      onChange={(e) => {
                        const f = e.target.files?.[0];
                        e.target.value = '';
                        if (f) void handleExtractCandidateScreenshot(f, 'edit');
                      }}
                    />
                  </label>
                  <button
                    type="button"
                    disabled={isExtractingCandidateAi}
                    onClick={() => void handleApplyChatToEditingPerson()}
                    className="inline-flex items-center gap-1 rounded-md bg-sky-600 px-3 py-1 text-[11px] font-semibold text-white hover:bg-sky-700 disabled:opacity-60"
                  >
                    <Sparkles className="h-3 w-3" aria-hidden="true" />
                    {isExtractingCandidateAi ? 'Đang bóc tách...' : 'AI bóc tách nguyên văn'}
                  </button>
                </div>
              </div>
              <textarea
                rows={3}
                value={editChatSnippet}
                onChange={(e) => setEditChatSnippet(e.target.value)}
                placeholder="Dán đoạn chat Messenger, ghi chú đàm phán lương/điều kiện hoặc nhấn Ctrl+V ảnh chụp màn hình vào đây..."
                className="w-full rounded-lg border border-sky-200 bg-white p-2 text-xs"
              />
              {!editingPerson.data.name && editingPerson.data.fbName && (
                <div className="mt-1.5 rounded border border-amber-200 bg-amber-50 px-2.5 py-1 text-[11px] font-medium text-amber-900">
                  ⚠️ Đoạn chat mới ghi nhận Tên Facebook ({editingPerson.data.fbName}), chưa thấy họ tên khai sinh. Có thể bấm Lưu để dùng tạm tên FB hoặc nhập họ tên thật bên dưới.
                </div>
              )}
            </div>

            <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
              <div>
                <label className="mb-1 block font-semibold">氏名 / Họ tên khai sinh</label>
                <input
                  type="text"
                  placeholder="Để trống nếu mới có tên FB"
                  value={editingPerson.data.name || ''}
                  onChange={(e) =>
                    setEditingPerson({
                      ...editingPerson,
                      data: { ...editingPerson.data, name: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="mb-1 block font-semibold">FBアカウント名 / Tên nick FB</label>
                <input
                  type="text"
                  value={editingPerson.data.fbName || ''}
                  onChange={(e) =>
                    setEditingPerson({
                      ...editingPerson,
                      data: { ...editingPerson.data, fbName: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="mb-1 block font-semibold">Link Facebook / Messenger</label>
                <input
                  type="url"
                  placeholder="https://m.me/... hoặc https://facebook.com/..."
                  value={editingPerson.data.fbProfileUrl || ''}
                  onChange={(e) =>
                    setEditingPerson({
                      ...editingPerson,
                      data: { ...editingPerson.data, fbProfileUrl: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                />
              </div>
              <div className="sm:col-span-3">
                <label className="mb-1 block font-semibold">流入チャネル / Nguồn kênh CRM (Chuyển đổi mượt)</label>
                <div className="flex flex-wrap gap-2">
                  {(
                    [
                      { id: 'personal', label: 'FB個人 (Facebook cá nhân)' },
                      { id: 'company', label: 'FB会社公式 (Fanpage công ty)' },
                      { id: 'school', label: '日本語学校・留学生 (Trường Nhật ngữ)' },
                    ] as Array<{ id: ChannelType; label: string }>
                  ).map((ch) => (
                    <button
                      key={ch.id}
                      type="button"
                      onClick={() =>
                        setEditingPerson({
                          ...editingPerson,
                          isSchool: ch.id === 'school',
                          data: { ...editingPerson.data, channel: ch.id },
                        })
                      }
                      className={`rounded-lg border px-3 py-1.5 text-[11px] font-semibold transition-colors ${
                        (editingPerson.data.channel || 'personal') === ch.id
                          ? 'border-indigo-500 bg-indigo-50 text-indigo-900'
                          : 'border-slate-200 bg-white text-slate-700 hover:bg-slate-50'
                      }`}
                    >
                      {ch.label}
                    </button>
                  ))}
                </div>
              </div>

              <div>
                <label className="mb-1 block font-semibold">国籍</label>
                <input
                  type="text"
                  value={editingPerson.data.nationality || ''}
                  onChange={(e) =>
                    setEditingPerson({
                      ...editingPerson,
                      data: { ...editingPerson.data, nationality: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="mb-1 block font-semibold">年齢 / 生年月日</label>
                <div className="flex gap-1.5">
                  <input
                    type="text"
                    placeholder="年齢"
                    value={editingPerson.data.age || ''}
                    onChange={(e) =>
                      setEditingPerson({
                        ...editingPerson,
                        data: { ...editingPerson.data, age: e.target.value },
                      })
                    }
                    className="w-16 rounded-lg border border-slate-300 px-2 py-1.5"
                  />
                  <input
                    type="text"
                    placeholder="YYYY-MM-DD"
                    value={editingPerson.data.birthDate || ''}
                    onChange={(e) =>
                      setEditingPerson({
                        ...editingPerson,
                        data: { ...editingPerson.data, birthDate: e.target.value },
                      })
                    }
                    className="flex-1 rounded-lg border border-slate-300 px-2 py-1.5"
                  />
                </div>
              </div>
              <div>
                <label className="mb-1 block font-semibold">日本語レベル</label>
                <select
                  value={editingPerson.data.jlpt || '要確認'}
                  onChange={(e) =>
                    setEditingPerson({
                      ...editingPerson,
                      data: { ...editingPerson.data, jlpt: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5"
                >
                  {JLPT_LEVELS.map((lvl) => (
                    <option key={lvl} value={lvl}>
                      {lvl}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="mb-1 block font-semibold">希望分野</label>
                <input
                  type="text"
                  value={editingPerson.data.desiredIndustry || ''}
                  onChange={(e) =>
                    setEditingPerson({
                      ...editingPerson,
                      data: { ...editingPerson.data, desiredIndustry: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="mb-1 block font-semibold">希望勤務エリア</label>
                <input
                  type="text"
                  value={editingPerson.data.desiredArea || ''}
                  onChange={(e) =>
                    setEditingPerson({
                      ...editingPerson,
                      data: { ...editingPerson.data, desiredArea: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="mb-1 block font-semibold">進捗ステータス</label>
                <select
                  value={editingPerson.data.progressStatus || '新規相談'}
                  onChange={(e) =>
                    setEditingPerson({
                      ...editingPerson,
                      data: { ...editingPerson.data, progressStatus: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 font-semibold"
                >
                  {PROGRESS_STATUSES.map((st) => (
                    <option key={st} value={st}>
                      {st}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="mb-1 block font-semibold">在留資格 / 在留期限</label>
                <div className="flex gap-1.5">
                  <input
                    type="text"
                    placeholder="在留資格"
                    value={editingPerson.data.visaStatus || ''}
                    onChange={(e) =>
                      setEditingPerson({
                        ...editingPerson,
                        data: { ...editingPerson.data, visaStatus: e.target.value },
                      })
                    }
                    className="w-1/2 rounded-lg border border-slate-300 px-2 py-1.5"
                  />
                  <input
                    type="text"
                    placeholder="期限 YYYY-MM-DD"
                    value={editingPerson.data.visaExpiry || ''}
                    onChange={(e) =>
                      setEditingPerson({
                        ...editingPerson,
                        data: { ...editingPerson.data, visaExpiry: e.target.value },
                      })
                    }
                    className="w-1/2 rounded-lg border border-slate-300 px-2 py-1.5"
                  />
                </div>
              </div>
              <div>
                <label className="mb-1 block font-semibold">合格済み試験・資格</label>
                <input
                  type="text"
                  value={editingPerson.data.passedExam || ''}
                  onChange={(e) =>
                    setEditingPerson({
                      ...editingPerson,
                      data: { ...editingPerson.data, passedExam: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="mb-1 block font-semibold">提案先企業</label>
                <input
                  type="text"
                  value={editingPerson.data.proposedCompany || ''}
                  onChange={(e) =>
                    setEditingPerson({
                      ...editingPerson,
                      data: { ...editingPerson.data, proposedCompany: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                />
              </div>

              <div>
                <label className="mb-1 block font-semibold">
                  面談日 / Mendan (YYYY-MM-DD HH:mm)
                </label>
                <input
                  type="text"
                  placeholder="2026-11-25 14:00"
                  value={editingPerson.data.mendanDate || ''}
                  onChange={(e) =>
                    setEditingPerson({
                      ...editingPerson,
                      data: { ...editingPerson.data, mendanDate: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="mb-1 block font-semibold">
                  企業面接日時 / Phỏng vấn cty (YYYY-MM-DD HH:mm)
                </label>
                <input
                  type="text"
                  placeholder="2026-11-28 15:00"
                  value={editingPerson.data.interviewDate || ''}
                  onChange={(e) =>
                    setEditingPerson({
                      ...editingPerson,
                      data: { ...editingPerson.data, interviewDate: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="mb-1 block font-semibold">電話番号 / LINE</label>
                <input
                  type="text"
                  value={editingPerson.data.phone || ''}
                  onChange={(e) =>
                    setEditingPerson({
                      ...editingPerson,
                      data: { ...editingPerson.data, phone: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="mb-1 block font-semibold">入社可能時期</label>
                <input
                  type="text"
                  value={editingPerson.data.startDate || ''}
                  onChange={(e) =>
                    setEditingPerson({
                      ...editingPerson,
                      data: { ...editingPerson.data, startDate: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                />
              </div>
            </div>

            <div className="mt-3 grid grid-cols-1 gap-3 sm:grid-cols-2">
              <div>
                <label className="mb-1 block font-semibold">
                  Kinh nghiệm &amp; Bằng cấp chi tiết (日本での職歴・経験)
                </label>
                <input
                  type="text"
                  placeholder="VD: Đang là Kaigo năm cuối, 3 năm viện có phí..."
                  value={editingPerson.data.japanExperience || ''}
                  onChange={(e) =>
                    setEditingPerson({
                      ...editingPerson,
                      data: { ...editingPerson.data, japanExperience: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="mb-1 block font-semibold">
                  Lương kỳ vọng / Mốc đàm phán (給与希望)
                </label>
                <input
                  type="text"
                  placeholder="VD: Kỳ vọng thực nhận ~16 man sau trừ nhà 3 man"
                  value={editingPerson.data.salaryExpectation || ''}
                  onChange={(e) =>
                    setEditingPerson({
                      ...editingPerson,
                      data: { ...editingPerson.data, salaryExpectation: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                />
              </div>
            </div>

            <div className="mt-3">
              <label className="mb-1 block font-semibold">
                Nhu cầu &amp; Mục tiêu chuyển đổi (転職・在留資格変更の目的)
              </label>
              <input
                type="text"
                placeholder="VD: Tìm công ty mới phỏng vấn chuyển đổi visa Tokutei ngành Kaigo"
                value={editingPerson.data.conversionGoal || ''}
                onChange={(e) =>
                  setEditingPerson({
                    ...editingPerson,
                    data: { ...editingPerson.data, conversionGoal: e.target.value },
                  })
                }
                className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
              />
            </div>

            <div className="mt-3">
              <label className="mb-1 block font-semibold">
                Nhật ký &amp; Tình trạng trao đổi nguyên văn (Không cắt bớt)
              </label>
              <textarea
                rows={4}
                value={editingPerson.data.negotiationLog || ''}
                onChange={(e) =>
                  setEditingPerson({
                    ...editingPerson,
                    data: { ...editingPerson.data, negotiationLog: e.target.value },
                  })
                }
                placeholder="Lưu nguyên văn các mốc đàm phán lương/điều kiện, đơn đang hỏi, phản hồi của ứng viên..."
                className="w-full rounded-lg border border-slate-300 p-2.5"
              />
            </div>

            <div className="mt-3">
              <label className="mb-1 block font-semibold">Ghi chú bổ sung (面談記録・対応メモ)</label>
              <textarea
                rows={3}
                value={editingPerson.data.memo || ''}
                onChange={(e) =>
                  setEditingPerson({
                    ...editingPerson,
                    data: { ...editingPerson.data, memo: e.target.value },
                  })
                }
                className="w-full rounded-lg border border-slate-300 p-2.5"
              />
            </div>

            <div className="mt-4 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setEditingPerson(null)}
                className="rounded-lg border border-slate-300 px-4 py-2 font-semibold text-slate-700"
              >
                キャンセル
              </button>
              <button
                type="button"
                onClick={handleSaveEditingPerson}
                className="rounded-lg bg-indigo-600 px-5 py-2 font-semibold text-white hover:bg-indigo-700"
              >
                保存する
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 2: FB CHAT AUTO-REGISTER / UPDATE */}
      {fbChatModalOpen && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-black/60 p-4">
          <div
            className="max-h-[90vh] w-full max-w-2xl overflow-y-auto rounded-xl bg-white p-5 text-xs text-slate-900 shadow-2xl"
            onPaste={(e) => {
              const img = extractFirstImageFromClipboard(e.clipboardData);
              if (img) {
                e.preventDefault();
                void handleExtractCandidateScreenshot(img, 'fbChat');
              }
            }}
          >
            <div className="mb-3 flex items-center justify-between border-b border-slate-200 pb-2.5">
              <h3 className="text-base font-bold text-slate-900">
                FBチャットから自動登録・既存相談者の状況更新
              </h3>
              <button
                type="button"
                onClick={() => setFbChatModalOpen(false)}
                className="rounded-lg bg-slate-100 px-3 py-1 font-semibold text-slate-700"
              >
                閉じる
              </button>
            </div>

            <div className="mb-3 flex flex-wrap gap-2">
              <button
                type="button"
                onClick={() => setFbChatMode('new')}
                className={`rounded-lg px-3 py-2 font-semibold ${
                  fbChatMode === 'new'
                    ? 'bg-indigo-600 text-white'
                    : 'bg-slate-100 text-slate-700'
                }`}
              >
                ＋ 新規相談者として登録
              </button>
              <button
                type="button"
                onClick={() => setFbChatMode('update')}
                className={`rounded-lg px-3 py-2 font-semibold ${
                  fbChatMode === 'update'
                    ? 'bg-emerald-600 text-white'
                    : 'bg-slate-100 text-slate-700'
                }`}
              >
                既存相談者の状況を更新
              </button>
            </div>

            {fbChatMode === 'update' && (
              <div className="mb-3">
                <label className="mb-1 block font-semibold">更新対象の相談者を選択</label>
                <select
                  value={fbChatTargetId}
                  onChange={(e) => {
                    const nextId = e.target.value;
                    setFbChatTargetId(nextId);
                    const found = allPeopleCombined.find((p) => p.id === nextId);
                    if (found) {
                      setFbChatChannel(
                        found._isSchool || found.channel === 'school'
                          ? 'school'
                          : found.channel === 'company'
                          ? 'company'
                          : 'personal'
                      );
                    }
                  }}
                  className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 font-semibold"
                >
                  <option value="">-- 相談者を選択してください --</option>
                  {allPeopleCombined.map((p) => (
                    <option key={p.id} value={p.id}>
                      {p.name} {p.fbName && p.fbName !== p.name ? `(${p.fbName})` : ''} - {p.progressStatus}
                    </option>
                  ))}
                </select>
              </div>
            )}

            <div className="mb-3">
              <label className="mb-1 block font-semibold">
                {fbChatMode === 'new'
                  ? '登録先チャネル / Chọn kênh nguồn'
                  : 'チャネル設定 / Kênh lưu trữ (có thể chuyển đổi kênh ngay)'}
              </label>
              <div className="flex flex-wrap gap-2">
                {(
                  [
                    { id: 'personal', label: 'FB個人' },
                    { id: 'company', label: 'FB会社公式' },
                    { id: 'school', label: '日本語学校・留学生' },
                  ] as Array<{ id: ChannelType; label: string }>
                ).map((ch) => (
                  <button
                    key={ch.id}
                    type="button"
                    onClick={() => setFbChatChannel(ch.id)}
                    className={`rounded-lg border px-3 py-1.5 font-semibold ${
                      fbChatChannel === ch.id
                        ? 'border-indigo-500 bg-indigo-50 text-indigo-900'
                        : 'border-slate-200 bg-white text-slate-700'
                    }`}
                  >
                    {ch.label}
                  </button>
                ))}
              </div>
            </div>

            <div>
              <div className="mb-1.5 flex flex-wrap items-center justify-between gap-2">
                <label className="font-semibold">
                  Dán đoạn chat Messenger, ghi chú hoặc Ctrl+V ảnh chụp màn hình (Chat / Thẻ ngoại kiều)
                </label>
                <label className="inline-flex cursor-pointer items-center gap-1 rounded-md border border-sky-300 bg-sky-50 px-2.5 py-1 text-[11px] font-semibold text-sky-800 hover:bg-sky-100">
                  <ImageIcon className="h-3.5 w-3.5" aria-hidden="true" />
                  Chọn ảnh chụp màn hình
                  <input
                    type="file"
                    accept="image/*"
                    className="hidden"
                    onChange={(e) => {
                      const f = e.target.files?.[0];
                      e.target.value = '';
                      if (f) void handleExtractCandidateScreenshot(f, 'fbChat');
                    }}
                  />
                </label>
              </div>
              <textarea
                rows={6}
                value={fbChatRawText}
                onChange={(e) => setFbChatRawText(e.target.value)}
                placeholder={
                  'VD:\nLê Thị Bích Phương (FB: Thị Bích Phương Lê)\nĐang là Kaigo năm cuối ở Tochigi, muốn chuyển sang Saitama hoặc Gunma. Hạn thẻ ngoại kiều 09/04/2027, JLPT N3...'
                }
                className="w-full rounded-lg border border-slate-300 p-3 text-xs"
              />
            </div>

            <div className="mt-4 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setFbChatModalOpen(false)}
                className="rounded-lg border border-slate-300 px-4 py-2 font-semibold text-slate-700"
              >
                Đóng
              </button>
              <button
                type="button"
                disabled={isExtractingFbChatAi}
                onClick={() => void handleSubmitFbChatModal()}
                className="rounded-lg bg-sky-600 px-5 py-2 font-semibold text-white hover:bg-sky-700 disabled:opacity-60"
              >
                {isExtractingFbChatAi
                  ? 'Đang phân tích & bảo toàn dữ liệu...'
                  : 'AI bóc tách & Lưu nguyên văn'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 3: 2-COLUMN COMPANY STANDARD RIREKISHO */}
      {rirekishoModal && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-black/60 p-4">
          <div className="max-h-[92vh] w-full max-w-6xl overflow-y-auto rounded-xl bg-white p-5 text-xs text-slate-900 shadow-2xl">
            <div className="mb-3 flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 pb-3">
              <div>
                <h3 className="text-base font-bold text-slate-900">
                  履歴書作成（企業提出用2カラム標準様式）：{rirekishoModal.person?.name}
                </h3>
                <p className="text-[11px] text-slate-500">
                  既存の履歴書ファイル読込・チャット自動作成・直接編集・HTML/PDF出力に対応
                </p>
              </div>
              <div className="flex flex-wrap items-center gap-2">
                <button
                  type="button"
                  disabled={isParsingRirekisho}
                  onClick={() => rirekishoFileInputRef.current?.click()}
                  className="rounded-lg bg-amber-600 px-3.5 py-1.5 text-xs font-semibold text-white hover:bg-amber-700 disabled:opacity-60"
                >
                  {isParsingRirekisho ? '読込・解析中...' : '既存履歴書ファイル読込'}
                </button>
                <button
                  type="button"
                  onClick={handleDownloadRirekishoFile}
                  className="rounded-lg bg-emerald-600 px-3.5 py-1.5 text-xs font-semibold text-white hover:bg-emerald-700"
                >
                  履歴書をダウンロード
                </button>
                <button
                  type="button"
                  onClick={handlePrintRirekishoSafely}
                  className="rounded-lg bg-indigo-600 px-3.5 py-1.5 text-xs font-semibold text-white hover:bg-indigo-700"
                >
                  印刷 / PDF保存
                </button>
                <button
                  type="button"
                  onClick={() => setRirekishoModal(null)}
                  className="rounded-lg bg-slate-100 px-3 py-1.5 font-semibold text-slate-700"
                >
                  閉じる
                </button>
              </div>
            </div>

            <div className="mb-3 flex flex-wrap items-center justify-between gap-3 rounded-lg border border-amber-300 bg-amber-50/80 p-3">
              <div>
                <div className="font-bold text-amber-950">
                  応募者が作成した既存の履歴書ファイルから自動入力
                </div>
                <div className="mt-0.5 text-[11px] text-amber-900">
                  対応形式：PDF・画像（JPG/PNG）・Excel（.xlsx / .xls）・Word（.docx）・CSV・Text
                </div>
              </div>
              <button
                type="button"
                disabled={isParsingRirekisho}
                onClick={() => rirekishoFileInputRef.current?.click()}
                className="rounded-lg bg-amber-600 px-4 py-2 text-xs font-semibold text-white hover:bg-amber-700 disabled:opacity-60"
              >
                {isParsingRirekisho
                  ? 'ファイル解析中...'
                  : 'ファイルを選択して自動反映（PDF / Excel / Word / 画像）'}
              </button>
            </div>

            {rirekishoStatusBanner && (
              <div
                className={`mb-3 rounded-lg border px-3.5 py-2.5 text-xs font-semibold ${
                  rirekishoStatusBanner.type === 'loading'
                    ? 'border-indigo-200 bg-indigo-50 text-indigo-900'
                    : rirekishoStatusBanner.type === 'success'
                    ? 'border-emerald-200 bg-emerald-50 text-emerald-900'
                    : 'border-rose-200 bg-rose-50 text-rose-900'
                }`}
              >
                {rirekishoStatusBanner.text}
              </div>
            )}

            <div className="mb-4 rounded-lg border border-indigo-200 bg-indigo-50/60 p-3">
              <div className="mb-1.5 flex flex-wrap items-center justify-between gap-2">
                <span className="font-bold text-indigo-950">
                  応募者のチャット・ヒアリングメモから自動入力
                </span>
                <button
                  type="button"
                  disabled={isParsingRirekisho}
                  onClick={handleGenerateRirekishoFromChat}
                  className="rounded-md bg-indigo-600 px-3 py-1.5 text-[11px] font-semibold text-white hover:bg-indigo-700 disabled:opacity-60"
                >
                  {isParsingRirekisho ? 'AI解析中...' : 'チャットから履歴書を自動作成'}
                </button>
              </div>
              <textarea
                rows={2}
                value={rirekishoChatInput}
                onChange={(e) => setRirekishoChatInput(e.target.value)}
                placeholder="応募者とのFBメッセンジャーやり取りや面談メモを貼り付けてボタンを押すと、生年月日・国籍・在留期限・学歴職歴・資格・志望動機などを2カラム様式へ自動入力します..."
                className="w-full rounded-lg border border-indigo-200 bg-white p-2 text-xs"
              />
            </div>

            <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
              <div className="space-y-3 rounded-lg border-2 border-slate-800 p-3.5">
                <div className="flex items-center justify-between border-b border-slate-400 pb-1.5">
                  <span className="text-lg font-bold tracking-widest">履 歴 書</span>
                  <input
                    type="text"
                    value={rirekishoModal.draft.dateStr}
                    onChange={(e) =>
                      setRirekishoModal({
                        ...rirekishoModal,
                        draft: { ...rirekishoModal.draft, dateStr: e.target.value },
                      })
                    }
                    className="w-44 border-b border-slate-400 text-right text-xs font-semibold"
                  />
                </div>

                <div className="grid grid-cols-12 gap-2 border border-slate-400 p-2">
                  <div className="col-span-3 font-semibold text-slate-600">フリガナ</div>
                  <div className="col-span-9">
                    <input
                      type="text"
                      value={rirekishoModal.draft.furigana}
                      onChange={(e) =>
                        setRirekishoModal({
                          ...rirekishoModal,
                          draft: { ...rirekishoModal.draft, furigana: e.target.value },
                        })
                      }
                      className="w-full border-b border-slate-300 px-1"
                    />
                  </div>
                  <div className="col-span-3 font-semibold text-slate-600">氏名</div>
                  <div className="col-span-9">
                    <input
                      type="text"
                      value={rirekishoModal.draft.fullName}
                      onChange={(e) =>
                        setRirekishoModal({
                          ...rirekishoModal,
                          draft: { ...rirekishoModal.draft, fullName: e.target.value },
                        })
                      }
                      className="w-full border-b border-slate-300 px-1 text-sm font-bold"
                    />
                  </div>
                  <div className="col-span-3 font-semibold text-slate-600">生年月日</div>
                  <div className="col-span-9 flex items-center gap-2">
                    <input
                      type="text"
                      value={rirekishoModal.draft.birthDate}
                      onChange={(e) => {
                        const bd = e.target.value;
                        const autoAge = calcJapaneseAgeFromBirthDate(bd);
                        setRirekishoModal({
                          ...rirekishoModal,
                          draft: {
                            ...rirekishoModal.draft,
                            birthDate: bd,
                            age: autoAge || rirekishoModal.draft.age,
                          },
                        });
                      }}
                      placeholder="2000-01-01"
                      className="w-32 border-b border-slate-300 px-1"
                    />
                    <span>（満</span>
                    <input
                      type="text"
                      value={rirekishoModal.draft.age}
                      onChange={(e) =>
                        setRirekishoModal({
                          ...rirekishoModal,
                          draft: { ...rirekishoModal.draft, age: e.target.value },
                        })
                      }
                      className="w-10 border-b border-slate-300 text-center font-bold"
                    />
                    <span>歳）性別：</span>
                    <input
                      type="text"
                      value={rirekishoModal.draft.gender}
                      onChange={(e) =>
                        setRirekishoModal({
                          ...rirekishoModal,
                          draft: { ...rirekishoModal.draft, gender: e.target.value },
                        })
                      }
                      className="w-14 border-b border-slate-300 text-center"
                    />
                  </div>
                  <div className="col-span-3 font-semibold text-slate-600">現住所 〒</div>
                  <div className="col-span-9 space-y-1">
                    <input
                      type="text"
                      placeholder="000-0000"
                      value={rirekishoModal.draft.postalCode}
                      onChange={(e) =>
                        setRirekishoModal({
                          ...rirekishoModal,
                          draft: { ...rirekishoModal.draft, postalCode: e.target.value },
                        })
                      }
                      className="w-28 border-b border-slate-300 px-1"
                    />
                    <input
                      type="text"
                      value={rirekishoModal.draft.address}
                      onChange={(e) =>
                        setRirekishoModal({
                          ...rirekishoModal,
                          draft: { ...rirekishoModal.draft, address: e.target.value },
                        })
                      }
                      className="w-full border-b border-slate-300 px-1"
                    />
                  </div>
                  <div className="col-span-3 font-semibold text-slate-600">電話 / メール</div>
                  <div className="col-span-9 grid grid-cols-2 gap-2">
                    <input
                      type="text"
                      placeholder="電話番号"
                      value={rirekishoModal.draft.phone}
                      onChange={(e) =>
                        setRirekishoModal({
                          ...rirekishoModal,
                          draft: { ...rirekishoModal.draft, phone: e.target.value },
                        })
                      }
                      className="border-b border-slate-300 px-1"
                    />
                    <input
                      type="text"
                      placeholder="Email"
                      value={rirekishoModal.draft.email}
                      onChange={(e) =>
                        setRirekishoModal({
                          ...rirekishoModal,
                          draft: { ...rirekishoModal.draft, email: e.target.value },
                        })
                      }
                      className="border-b border-slate-300 px-1"
                    />
                  </div>
                </div>

                <table className="w-full border-collapse border border-slate-500 text-center">
                  <thead>
                    <tr className="bg-slate-100 font-semibold">
                      <th className="border border-slate-400 p-1">国籍</th>
                      <th className="border border-slate-400 p-1">在留期限</th>
                      <th className="border border-slate-400 p-1">在留資格</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td className="border border-slate-400 p-1">
                        <input
                          type="text"
                          value={rirekishoModal.draft.nationality}
                          onChange={(e) =>
                            setRirekishoModal({
                              ...rirekishoModal,
                              draft: { ...rirekishoModal.draft, nationality: e.target.value },
                            })
                          }
                          className="w-full text-center"
                        />
                      </td>
                      <td className="border border-slate-400 p-1">
                        <input
                          type="text"
                          value={rirekishoModal.draft.visaExpiry}
                          onChange={(e) =>
                            setRirekishoModal({
                              ...rirekishoModal,
                              draft: { ...rirekishoModal.draft, visaExpiry: e.target.value },
                            })
                          }
                          className="w-full text-center"
                        />
                      </td>
                      <td className="border border-slate-400 p-1">
                        <input
                          type="text"
                          value={rirekishoModal.draft.visaStatus}
                          onChange={(e) =>
                            setRirekishoModal({
                              ...rirekishoModal,
                              draft: { ...rirekishoModal.draft, visaStatus: e.target.value },
                            })
                          }
                          className="w-full text-center"
                        />
                      </td>
                    </tr>
                  </tbody>
                </table>

                <table className="w-full border-collapse border border-slate-500">
                  <thead>
                    <tr className="bg-slate-100 font-semibold">
                      <th className="w-14 border border-slate-400 p-1">年</th>
                      <th className="w-10 border border-slate-400 p-1">月</th>
                      <th className="border border-slate-400 p-1">学歴・職歴</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td
                        colSpan={3}
                        className="border border-slate-400 bg-slate-50 text-center font-semibold"
                      >
                        学 歴
                      </td>
                    </tr>
                    {rirekishoModal.draft.educationRows.map((row, idx) => (
                      <tr key={`edu_${idx}`}>
                        <td className="border border-slate-400 p-1">
                          <input
                            type="text"
                            value={row.year}
                            onChange={(e) => {
                              const next = [...rirekishoModal.draft.educationRows];
                              next[idx] = { ...next[idx], year: e.target.value };
                              setRirekishoModal({
                                ...rirekishoModal,
                                draft: { ...rirekishoModal.draft, educationRows: next },
                              });
                            }}
                            className="w-full text-center"
                          />
                        </td>
                        <td className="border border-slate-400 p-1">
                          <input
                            type="text"
                            value={row.month}
                            onChange={(e) => {
                              const next = [...rirekishoModal.draft.educationRows];
                              next[idx] = { ...next[idx], month: e.target.value };
                              setRirekishoModal({
                                ...rirekishoModal,
                                draft: { ...rirekishoModal.draft, educationRows: next },
                              });
                            }}
                            className="w-full text-center"
                          />
                        </td>
                        <td className="border border-slate-400 p-1">
                          <input
                            type="text"
                            value={row.content}
                            onChange={(e) => {
                              const next = [...rirekishoModal.draft.educationRows];
                              next[idx] = { ...next[idx], content: e.target.value };
                              setRirekishoModal({
                                ...rirekishoModal,
                                draft: { ...rirekishoModal.draft, educationRows: next },
                              });
                            }}
                            className="w-full"
                          />
                        </td>
                      </tr>
                    ))}
                    <tr>
                      <td
                        colSpan={3}
                        className="border border-slate-400 bg-slate-50 text-center font-semibold"
                      >
                        職 歴
                      </td>
                    </tr>
                    {rirekishoModal.draft.workRows.map((row, idx) => (
                      <tr key={`wrk_${idx}`}>
                        <td className="border border-slate-400 p-1">
                          <input
                            type="text"
                            value={row.year}
                            onChange={(e) => {
                              const next = [...rirekishoModal.draft.workRows];
                              next[idx] = { ...next[idx], year: e.target.value };
                              setRirekishoModal({
                                ...rirekishoModal,
                                draft: { ...rirekishoModal.draft, workRows: next },
                              });
                            }}
                            className="w-full text-center"
                          />
                        </td>
                        <td className="border border-slate-400 p-1">
                          <input
                            type="text"
                            value={row.month}
                            onChange={(e) => {
                              const next = [...rirekishoModal.draft.workRows];
                              next[idx] = { ...next[idx], month: e.target.value };
                              setRirekishoModal({
                                ...rirekishoModal,
                                draft: { ...rirekishoModal.draft, workRows: next },
                              });
                            }}
                            className="w-full text-center"
                          />
                        </td>
                        <td className="border border-slate-400 p-1">
                          <input
                            type="text"
                            value={row.content}
                            onChange={(e) => {
                              const next = [...rirekishoModal.draft.workRows];
                              next[idx] = { ...next[idx], content: e.target.value };
                              setRirekishoModal({
                                ...rirekishoModal,
                                draft: { ...rirekishoModal.draft, workRows: next },
                              });
                            }}
                            className="w-full"
                          />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>

                <div className="border border-slate-500 p-2">
                  <div className="mb-1 font-semibold text-slate-700">
                    本人希望記入欄（特に給料・職種・勤務時間・勤務地・その他についての希望などがあれば記入）
                  </div>
                  <textarea
                    rows={3}
                    value={rirekishoModal.draft.personalRequests}
                    onChange={(e) =>
                      setRirekishoModal({
                        ...rirekishoModal,
                        draft: { ...rirekishoModal.draft, personalRequests: e.target.value },
                      })
                    }
                    className="w-full resize-none border border-slate-200 p-1.5"
                  />
                </div>
              </div>

              <div className="space-y-3 rounded-lg border-2 border-slate-800 p-3.5">
                <table className="w-full border-collapse border border-slate-500">
                  <thead>
                    <tr className="bg-slate-100 font-semibold">
                      <th className="w-14 border border-slate-400 p-1">年</th>
                      <th className="w-10 border border-slate-400 p-1">月</th>
                      <th className="border border-slate-400 p-1">免許・資格</th>
                    </tr>
                  </thead>
                  <tbody>
                    {rirekishoModal.draft.licenseRows.map((row, idx) => (
                      <tr key={`lic_${idx}`}>
                        <td className="border border-slate-400 p-1">
                          <input
                            type="text"
                            value={row.year}
                            onChange={(e) => {
                              const next = [...rirekishoModal.draft.licenseRows];
                              next[idx] = { ...next[idx], year: e.target.value };
                              setRirekishoModal({
                                ...rirekishoModal,
                                draft: { ...rirekishoModal.draft, licenseRows: next },
                              });
                            }}
                            className="w-full text-center"
                          />
                        </td>
                        <td className="border border-slate-400 p-1">
                          <input
                            type="text"
                            value={row.month}
                            onChange={(e) => {
                              const next = [...rirekishoModal.draft.licenseRows];
                              next[idx] = { ...next[idx], month: e.target.value };
                              setRirekishoModal({
                                ...rirekishoModal,
                                draft: { ...rirekishoModal.draft, licenseRows: next },
                              });
                            }}
                            className="w-full text-center"
                          />
                        </td>
                        <td className="border border-slate-400 p-1">
                          <input
                            type="text"
                            value={row.content}
                            onChange={(e) => {
                              const next = [...rirekishoModal.draft.licenseRows];
                              next[idx] = { ...next[idx], content: e.target.value };
                              setRirekishoModal({
                                ...rirekishoModal,
                                draft: { ...rirekishoModal.draft, licenseRows: next },
                              });
                            }}
                            className="w-full"
                          />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>

                <table className="w-full border-collapse border border-slate-500">
                  <thead>
                    <tr className="bg-slate-100 font-semibold">
                      <th colSpan={2} className="border border-slate-400 p-1 text-left">
                        話せる言語
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td className="w-28 border border-slate-400 bg-slate-50 p-1 font-semibold">
                        母国語
                      </td>
                      <td className="border border-slate-400 p-1">
                        <input
                          type="text"
                          value={rirekishoModal.draft.nativeLanguage}
                          onChange={(e) =>
                            setRirekishoModal({
                              ...rirekishoModal,
                              draft: { ...rirekishoModal.draft, nativeLanguage: e.target.value },
                            })
                          }
                          className="w-full"
                        />
                      </td>
                    </tr>
                    <tr>
                      <td className="border border-slate-400 bg-slate-50 p-1 font-semibold">
                        使用可能言語
                      </td>
                      <td className="border border-slate-400 p-1">
                        <input
                          type="text"
                          value={rirekishoModal.draft.availableLanguages}
                          onChange={(e) =>
                            setRirekishoModal({
                              ...rirekishoModal,
                              draft: {
                                ...rirekishoModal.draft,
                                availableLanguages: e.target.value,
                              },
                            })
                          }
                          className="w-full"
                        />
                      </td>
                    </tr>
                  </tbody>
                </table>

                <table className="w-full border-collapse border border-slate-500">
                  <thead>
                    <tr className="bg-slate-100 font-semibold">
                      <th colSpan={2} className="border border-slate-400 p-1 text-left">
                        日本語能力CA評価
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td className="w-28 border border-slate-400 bg-slate-50 p-1 font-semibold">
                        話す
                      </td>
                      <td className="border border-slate-400 p-1">
                        <input
                          type="text"
                          value={rirekishoModal.draft.caSpeaking}
                          onChange={(e) =>
                            setRirekishoModal({
                              ...rirekishoModal,
                              draft: { ...rirekishoModal.draft, caSpeaking: e.target.value },
                            })
                          }
                          className="w-full"
                        />
                      </td>
                    </tr>
                    <tr>
                      <td className="border border-slate-400 bg-slate-50 p-1 font-semibold">
                        聞く
                      </td>
                      <td className="border border-slate-400 p-1">
                        <input
                          type="text"
                          value={rirekishoModal.draft.caListening}
                          onChange={(e) =>
                            setRirekishoModal({
                              ...rirekishoModal,
                              draft: { ...rirekishoModal.draft, caListening: e.target.value },
                            })
                          }
                          className="w-full"
                        />
                      </td>
                    </tr>
                  </tbody>
                </table>

                <div className="space-y-2 border border-slate-500 p-2">
                  <div className="font-semibold text-slate-700">志望の動機、自己PRなど</div>
                  <div>
                    <div className="text-[11px] font-bold text-indigo-800">＜志望動機＞</div>
                    <textarea
                      rows={3}
                      value={rirekishoModal.draft.motivation}
                      onChange={(e) =>
                        setRirekishoModal({
                          ...rirekishoModal,
                          draft: { ...rirekishoModal.draft, motivation: e.target.value },
                        })
                      }
                      className="w-full border border-slate-200 p-1.5"
                    />
                  </div>
                  <div>
                    <div className="text-[11px] font-bold text-indigo-800">＜自己PR＞</div>
                    <textarea
                      rows={3}
                      value={rirekishoModal.draft.selfPr}
                      onChange={(e) =>
                        setRirekishoModal({
                          ...rirekishoModal,
                          draft: { ...rirekishoModal.draft, selfPr: e.target.value },
                        })
                      }
                      className="w-full border border-slate-200 p-1.5"
                    />
                  </div>
                </div>

                <table className="w-full border-collapse border border-slate-500 text-center">
                  <thead>
                    <tr className="bg-slate-100 font-semibold">
                      <th className="border border-slate-400 p-1">アレルギー</th>
                      <th className="border border-slate-400 p-1">扶養家族数(配偶者を除く)</th>
                      <th className="border border-slate-400 p-1">配偶者 / 配偶者の扶養義務</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td className="border border-slate-400 p-1">
                        <input
                          type="text"
                          value={rirekishoModal.draft.allergy}
                          onChange={(e) =>
                            setRirekishoModal({
                              ...rirekishoModal,
                              draft: { ...rirekishoModal.draft, allergy: e.target.value },
                            })
                          }
                          className="w-full text-center"
                        />
                      </td>
                      <td className="border border-slate-400 p-1">
                        <input
                          type="text"
                          value={rirekishoModal.draft.dependentsCount}
                          onChange={(e) =>
                            setRirekishoModal({
                              ...rirekishoModal,
                              draft: {
                                ...rirekishoModal.draft,
                                dependentsCount: e.target.value,
                              },
                            })
                          }
                          className="w-full text-center"
                        />
                      </td>
                      <td className="border border-slate-400 p-1">
                        <div className="flex items-center justify-center gap-1">
                          <input
                            type="text"
                            value={rirekishoModal.draft.spouse}
                            onChange={(e) =>
                              setRirekishoModal({
                                ...rirekishoModal,
                                draft: { ...rirekishoModal.draft, spouse: e.target.value },
                              })
                            }
                            className="w-12 border-b border-slate-300 text-center"
                          />
                          <span>/</span>
                          <input
                            type="text"
                            value={rirekishoModal.draft.spouseSupport}
                            onChange={(e) =>
                              setRirekishoModal({
                                ...rirekishoModal,
                                draft: {
                                  ...rirekishoModal.draft,
                                  spouseSupport: e.target.value,
                                },
                              })
                            }
                            className="w-12 border-b border-slate-300 text-center"
                          />
                        </div>
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 4: SMART REPLY MESSAGE GENERATOR */}
      {replyModal && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-black/60 p-4">
          <div className="max-h-[90vh] w-full max-w-3xl overflow-y-auto rounded-xl bg-white p-5 text-xs text-slate-900 shadow-2xl">
            <div className="mb-3 flex items-center justify-between border-b border-slate-200 pb-2.5">
              <h3 className="text-base font-bold text-slate-900">
                返信文提案（VN / JP / EN）：{replyModal.person.name}
              </h3>
              <button
                type="button"
                onClick={() => setReplyModal(null)}
                className="rounded-lg bg-slate-100 px-3 py-1 font-semibold text-slate-700"
              >
                閉じる
              </button>
            </div>

            <div className="mb-3 flex flex-wrap gap-1.5">
              {(
                [
                  { id: 'first_mendan', label: '1. Chào & Hẹn Mendan đầu' },
                  { id: 'missing_info_request', label: '2. Xin thông tin thiếu' },
                  { id: 'recommend_job', label: '3. Gửi 3 Job phù hợp nhất' },
                  { id: 'remind_interview', label: '4. Nhắc lịch Mendan / PV' },
                  { id: 'post_interview_followup', label: '5. Hỏi thăm sau phỏng vấn' },
                  { id: 'followup_silent', label: '6. Chăm sóc lại / Chưa phản hồi' },
                ] as Array<{ id: ReplyScenario; label: string }>
              ).map((sc) => (
                <button
                  key={sc.id}
                  type="button"
                  onClick={() => setReplyModal({ ...replyModal, scenario: sc.id })}
                  className={`rounded-lg px-3 py-1.5 font-semibold transition-colors ${
                    replyModal.scenario === sc.id
                      ? 'bg-indigo-600 text-white'
                      : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                  }`}
                >
                  {sc.label}
                </button>
              ))}
            </div>

            {(() => {
              const msgs = buildSmartReplyMessages(
                replyModal.person,
                replyModal.scenario,
                kyujins
              );
              return (
                <div className="space-y-3">
                  {[
                    { lang: 'ベトナム語 (Tiếng Việt)', text: msgs.vi },
                    { lang: '日本語 (Japanese)', text: msgs.jp },
                    { lang: '英語 (English)', text: msgs.en },
                  ].map((item) => (
                    <div
                      key={item.lang}
                      className="rounded-lg border border-slate-200 bg-slate-50/70 p-3"
                    >
                      <div className="mb-1.5 flex items-center justify-between">
                        <span className="font-bold text-slate-900">{item.lang}</span>
                        <button
                          type="button"
                          onClick={async () => {
                            const ok = await copyTextSafely(item.text);
                            showToast(ok ? `📋 ${item.lang} をコピーしました` : '⚠️ コピー失敗');
                          }}
                          className="rounded-md bg-slate-900 px-3 py-1 text-[11px] font-semibold text-white hover:bg-slate-700"
                        >
                          コピー
                        </button>
                      </div>
                      <pre className="whitespace-pre-wrap font-sans leading-relaxed text-slate-800">
                        {item.text}
                      </pre>
                    </div>
                  ))}
                </div>
              );
            })()}
          </div>
        </div>
      )}

      {/* MODAL 5: CANDIDATE-FACING KYUJIN PREVIEW */}
      {previewModal && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-black/60 p-4">
          <div className="max-h-[90vh] w-full max-w-2xl overflow-y-auto rounded-xl bg-white p-5 text-slate-900 shadow-2xl">
            <div className="mb-3 flex items-center justify-between border-b border-slate-200 pb-2.5">
              <h3 className="text-base font-bold text-slate-900">
                求職者用 紹介文プレビュー（企業名・詳細住所非公開）
              </h3>
              <button
                type="button"
                onClick={() => setPreviewModal(null)}
                className="rounded-lg bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-700"
              >
                閉じる
              </button>
            </div>

            {(() => {
              const { company, slot } = previewModal;
              const jpIntroText = buildAnonymousJobPostText(company, slot, 'jp');
              const viIntroText = buildAnonymousJobPostText(company, slot, 'vi');

              return (
                <div className="space-y-3 text-xs">
                  <div className="rounded-lg border border-emerald-200 bg-emerald-50/50 p-3">
                    <div className="mb-1.5 flex items-center justify-between">
                      <span className="font-bold text-emerald-950">Bản tóm tắt Tiếng Việt (Ẩn danh)</span>
                      <button
                        type="button"
                        onClick={async () => {
                          const ok = await copyTextSafely(viIntroText);
                          showToast(ok ? '📋 Đã sao chép bản tóm tắt Tiếng Việt' : '⚠️ Sao chép thất bại');
                        }}
                        className="rounded-md bg-emerald-600 px-3 py-1 text-[11px] font-semibold text-white hover:bg-emerald-700"
                      >
                        Sao chép Tiếng Việt
                      </button>
                    </div>
                    <pre className="whitespace-pre-wrap font-sans leading-relaxed text-slate-800">
                      {viIntroText}
                    </pre>
                  </div>

                  <div className="rounded-lg border border-indigo-200 bg-indigo-50/50 p-3">
                    <div className="mb-1.5 flex items-center justify-between">
                      <span className="font-bold text-indigo-950">Bản tóm tắt Tiếng Nhật (Ẩn danh)</span>
                      <button
                        type="button"
                        onClick={async () => {
                          const ok = await copyTextSafely(jpIntroText);
                          showToast(ok ? '📋 Đã sao chép bản tóm tắt Tiếng Nhật' : '⚠️ Sao chép thất bại');
                        }}
                        className="rounded-md bg-indigo-600 px-3 py-1 text-[11px] font-semibold text-white hover:bg-indigo-700"
                      >
                        Sao chép Tiếng Nhật
                      </button>
                    </div>
                    <pre className="whitespace-pre-wrap font-sans leading-relaxed text-slate-800">
                      {jpIntroText}
                    </pre>
                  </div>
                </div>
              );
            })()}
          </div>
        </div>
      )}

      {/* MODAL 6: PT PARTNER EDIT / CREATE */}
      {ptModal && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-black/60 p-4">
          <div className="max-h-[90vh] w-full max-w-lg overflow-y-auto rounded-xl bg-white p-5 text-xs text-slate-900 shadow-2xl">
            <div className="mb-3 flex items-center justify-between border-b border-slate-200 pb-2">
              <h3 className="text-base font-bold text-slate-900">
                {ptModal.isNew ? '＋ 新規PT提携先登録' : 'PT提携先編集'}
              </h3>
              <button
                type="button"
                onClick={() => setPtModal(null)}
                className="rounded-lg bg-slate-100 px-3 py-1 font-semibold"
              >
                閉じる
              </button>
            </div>
            <div className="space-y-2.5">
              <div>
                <label className="mb-1 block font-semibold">PT団体・企業名 *</label>
                <input
                  type="text"
                  value={ptModal.data.partnerName}
                  onChange={(e) =>
                    setPtModal({
                      ...ptModal,
                      data: { ...ptModal.data, partnerName: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                />
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="mb-1 block font-semibold">担当者名</label>
                  <input
                    type="text"
                    value={ptModal.data.contactPerson}
                    onChange={(e) =>
                      setPtModal({
                        ...ptModal,
                        data: { ...ptModal.data, contactPerson: e.target.value },
                      })
                    }
                    className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                  />
                </div>
                <div>
                  <label className="mb-1 block font-semibold">提携ステータス</label>
                  <select
                    value={ptModal.data.status}
                    onChange={(e) =>
                      setPtModal({
                        ...ptModal,
                        data: { ...ptModal.data, status: e.target.value },
                      })
                    }
                    className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5"
                  >
                    <option value="提携中（稼働）">提携中（稼働）</option>
                    <option value="条件調整中">条件調整中</option>
                    <option value="休止・保留">休止・保留</option>
                  </select>
                </div>
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="mb-1 block font-semibold">電話番号</label>
                  <input
                    type="text"
                    value={ptModal.data.phone}
                    onChange={(e) =>
                      setPtModal({
                        ...ptModal,
                        data: { ...ptModal.data, phone: e.target.value },
                      })
                    }
                    className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                  />
                </div>
                <div>
                  <label className="mb-1 block font-semibold">メール</label>
                  <input
                    type="text"
                    value={ptModal.data.email}
                    onChange={(e) =>
                      setPtModal({
                        ...ptModal,
                        data: { ...ptModal.data, email: e.target.value },
                      })
                    }
                    className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                  />
                </div>
              </div>
              <div>
                <label className="mb-1 block font-semibold">得意分野</label>
                <input
                  type="text"
                  value={ptModal.data.specialty}
                  onChange={(e) =>
                    setPtModal({
                      ...ptModal,
                      data: { ...ptModal.data, specialty: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="mb-1 block font-semibold">契約・紹介シェア条件</label>
                <input
                  type="text"
                  value={ptModal.data.commissionTerms}
                  onChange={(e) =>
                    setPtModal({
                      ...ptModal,
                      data: { ...ptModal.data, commissionTerms: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="mb-1 block font-semibold">備考メモ</label>
                <textarea
                  rows={2}
                  value={ptModal.data.memo}
                  onChange={(e) =>
                    setPtModal({
                      ...ptModal,
                      data: { ...ptModal.data, memo: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 p-2"
                />
              </div>
            </div>
            <div className="mt-4 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setPtModal(null)}
                className="rounded-lg border border-slate-300 px-4 py-2 font-semibold"
              >
                キャンセル
              </button>
              <button
                type="button"
                onClick={() => {
                  if (!ptModal.data.partnerName.trim()) {
                    showToast('⚠️ PT団体名を入力してください');
                    return;
                  }
                  if (ptModal.isNew) {
                    setPtPartners((prev) => [ptModal.data, ...prev]);
                  } else {
                    setPtPartners((prev) =>
                      prev.map((x) => (x.id === ptModal.data.id ? ptModal.data : x))
                    );
                  }
                  setPtModal(null);
                  showToast('✅ PT提携先を保存しました');
                }}
                className="rounded-lg bg-indigo-600 px-4 py-2 font-semibold text-white"
              >
                保存
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 7: CUSTOM CONFIRMATION DIALOG */}
      {confirmDialog && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-black/70 p-4">
          <div className="w-full max-w-md rounded-xl border border-[#27272a] bg-[#141417] p-5 text-[#edeeef] shadow-2xl">
            <h4 className="text-base font-bold text-[#edeeef]">{confirmDialog.title}</h4>
            <p className="mt-2 whitespace-pre-line text-xs leading-relaxed text-[#94949e]">
              {confirmDialog.message}
            </p>
            <div className="mt-4 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => {
                  confirmDialog.onCancel?.();
                  setConfirmDialog(null);
                }}
                className="rounded-md border border-[#27272a] px-4 py-2 text-xs font-semibold text-[#94949e] hover:text-[#edeeef] cursor-pointer"
              >
                {confirmDialog.cancelLabel || 'キャンセル'}
              </button>
              <button
                type="button"
                onClick={confirmDialog.onConfirm}
                className={`rounded-md px-4 py-2 text-xs font-semibold cursor-pointer hover:opacity-90 ${
                  confirmDialog.tone === 'accent'
                    ? 'bg-[#22c55e] text-[#0c0c0e]'
                    : 'bg-[#ef4444] text-white'
                }`}
              >
                {confirmDialog.confirmLabel || '削除を実行する'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 8: FACEBOOK MESSENGER CONNECTION MODAL */}
      {messengerLinkModalOpen && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-black/75 p-4">
          <div className="max-h-[90vh] w-full max-w-xl overflow-y-auto rounded-xl border border-[#27272a] bg-[#141417] p-6 text-xs text-[#edeeef] shadow-2xl">
            <div className="mb-4 flex items-center justify-between border-b border-[#27272a] pb-3">
              <div className="flex items-center gap-2.5">
                <div className="flex h-8 w-8 items-center justify-center rounded-lg border border-[#38bdf8]/40 bg-[#38bdf8]/15 text-[#38bdf8]">
                  <Link2 className="h-4 w-4" aria-hidden="true" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-[#edeeef]">
                    Liên kết Facebook Messenger
                  </h3>
                  <p className="text-[11px] text-[#94949e]">
                    Kết nối Fanpage &amp; Meta Messenger Webhook tự động nhận hồ sơ ứng viên
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setMessengerLinkModalOpen(false)}
                className="rounded-md border border-[#27272a] bg-[#0c0c0e] px-3 py-1.5 text-xs font-semibold text-[#94949e] hover:text-[#edeeef] cursor-pointer"
              >
                Đóng
              </button>
            </div>

            {/* Connection Status Grid */}
            <div className="mb-4 grid grid-cols-3 gap-2">
              <div className="rounded-lg border border-[#27272a] bg-[#0c0c0e] p-3">
                <div className="font-mono text-[10px] uppercase text-[#94949e]">
                  Webhook
                </div>
                <div className="mt-1 font-mono text-xs font-bold text-[#22c55e]">
                  Sẵn sàng
                </div>
              </div>
              <div className="rounded-lg border border-[#27272a] bg-[#0c0c0e] p-3">
                <div className="font-mono text-[10px] uppercase text-[#94949e]">
                  Verify Token
                </div>
                <div
                  className={`mt-1 font-mono text-xs font-bold ${
                    webhookStatus?.hasVerifyToken ? 'text-[#22c55e]' : 'text-[#38bdf8]'
                  }`}
                >
                  {webhookStatus?.hasVerifyToken ? 'Đã kiểm tra' : 'Tùy chọn (Mở)'}
                </div>
              </div>
              <div className="rounded-lg border border-[#27272a] bg-[#0c0c0e] p-3">
                <div className="font-mono text-[10px] uppercase text-[#94949e]">
                  App Secret
                </div>
                <div
                  className={`mt-1 font-mono text-xs font-bold ${
                    webhookStatus?.hasAppSecret ? 'text-[#22c55e]' : 'text-[#38bdf8]'
                  }`}
                >
                  {webhookStatus?.hasAppSecret ? 'Đã bảo vệ HMAC' : 'Tùy chọn (Mở)'}
                </div>
              </div>
            </div>

            {/* Webhook Callback URL */}
            <div className="mb-4 rounded-lg border border-[#27272a] bg-[#0c0c0e] p-4">
              <label className="mb-1.5 block font-mono text-[10px] uppercase tracking-wider text-[#94949e]">
                Callback URL (Dán vào Meta Developers &rarr; Messenger &rarr; Webhooks)
              </label>
              <div className="flex flex-wrap gap-2">
                <input
                  type="text"
                  readOnly
                  value={sharedMessengerWebhookUrl || liveMessengerWebhookUrl}
                  aria-label="Callback URL Messenger Webhook"
                  className="min-w-[240px] flex-1 rounded-md border border-[#27272a] bg-[#141417] px-3 py-2 font-mono text-[11px] text-[#edeeef]"
                />
                <button
                  type="button"
                  onClick={() => {
                    const target = sharedMessengerWebhookUrl || liveMessengerWebhookUrl;
                    copyTextSafely(target);
                    showToast(`Đã sao chép Callback URL: ${target}`);
                  }}
                  className="inline-flex items-center gap-1.5 rounded-md bg-[#22c55e] px-3.5 py-2 text-xs font-semibold text-[#0c0c0e] hover:opacity-90 cursor-pointer"
                >
                  <Copy className="h-3.5 w-3.5" aria-hidden="true" />
                  Sao chép URL
                </button>
              </div>
              {sharedMessengerWebhookUrl && (
                <div className="mt-2 rounded border border-[#f59e0b]/40 bg-[#f59e0b]/10 px-3 py-2 text-[11px] leading-relaxed text-[#edeeef]/90">
                  <strong className="text-[#f59e0b]">Lưu ý quan trọng:</strong> Link <span className="font-mono text-[#edeeef]">ais-dev-...</span> chỉ dùng để bạn xem trước trong trình duyệt. Để Meta Developers gọi Webhook từ bên ngoài, hãy bấm <strong>Share / Publish</strong> ứng dụng và dán URL có đuôi <span className="font-mono text-[#22c55e]">/api/webhooks/messenger</span> ở trên.
                </div>
              )}
              <p className="mt-2 text-[11px] leading-relaxed text-[#94949e]">
                Bật 2 trường đăng ký webhook: <span className="font-mono text-[#edeeef]">messages</span> và <span className="font-mono text-[#edeeef]">messaging_postbacks</span>. Nếu muốn tăng cường bảo mật, bạn có thể đặt thêm <span className="font-mono text-[#edeeef]">FB_VERIFY_TOKEN</span> và <span className="font-mono text-[#edeeef]">FB_APP_SECRET</span> trong Settings &rarr; Secrets (không bắt buộc để nhận tin).
              </p>
            </div>

            {/* Fanpage Messenger Quick Link */}
            <div className="mb-4 rounded-lg border border-[#27272a] bg-[#0c0c0e] p-4">
              <label className="mb-1.5 block font-mono text-[10px] uppercase tracking-wider text-[#94949e]">
                Liên kết Fanpage Messenger (m.me / Page ID / Facebook URL)
              </label>
              <div className="flex flex-wrap gap-2">
                <input
                  type="text"
                  value={fbPageLinkDraft}
                  onChange={(e) => setFbPageLinkDraft(e.target.value)}
                  placeholder="VD: tokuteicareer hoặc https://m.me/tokuteicareer"
                  className="min-w-[220px] flex-1 rounded-md border border-[#27272a] bg-[#141417] px-3 py-2 text-xs text-[#edeeef] placeholder:text-[#94949e] focus:border-[#38bdf8] focus:outline-none"
                />
                <button
                  type="button"
                  onClick={handleSaveFbPageLink}
                  className="rounded-md bg-[#38bdf8] px-3.5 py-2 text-xs font-semibold text-[#0c0c0e] hover:opacity-90 cursor-pointer"
                >
                  Lưu liên kết
                </button>
              </div>

              <div className="mt-3 flex flex-wrap gap-2">
                {fbPageLink && (
                  <a
                    href={fbPageLink}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-1.5 rounded-md border border-[#38bdf8]/50 bg-[#38bdf8]/10 px-3 py-1.5 text-[11px] font-semibold text-[#38bdf8] hover:bg-[#38bdf8]/20"
                  >
                    <MessageCircle className="h-3.5 w-3.5" aria-hidden="true" />
                    Mở Messenger Fanpage
                  </a>
                )}
                <a
                  href="https://business.facebook.com/latest/inbox/all"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="inline-flex items-center gap-1.5 rounded-md border border-[#27272a] bg-[#141417] px-3 py-1.5 text-[11px] font-semibold text-[#edeeef] hover:border-[#38bdf8]"
                >
                  <ExternalLink className="h-3.5 w-3.5" aria-hidden="true" />
                  Mở Meta Business Inbox
                </a>
                <a
                  href="https://developers.facebook.com/apps/"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="inline-flex items-center gap-1.5 rounded-md border border-[#27272a] bg-[#141417] px-3 py-1.5 text-[11px] font-semibold text-[#edeeef] hover:border-[#22c55e]"
                >
                  <ExternalLink className="h-3.5 w-3.5" aria-hidden="true" />
                  Mở Meta Developers
                </a>
                <button
                  type="button"
                  disabled={isTestingWebhook}
                  onClick={() => void handleTestMessengerWebhook()}
                  className="inline-flex items-center gap-1.5 rounded-md border border-[#38bdf8]/50 bg-[#38bdf8]/10 px-3 py-1.5 text-[11px] font-semibold text-[#38bdf8] hover:bg-[#38bdf8]/20 disabled:opacity-50 cursor-pointer"
                >
                  <MessageCircle className="h-3.5 w-3.5" aria-hidden="true" />
                  {isTestingWebhook
                    ? 'Đang thử nghiệm...'
                    : 'Thử nhận tin nhắn Messenger'}
                </button>
                <button
                  type="button"
                  disabled={isRefreshingSql}
                  onClick={() => void refreshCrmFromSql(false)}
                  className="inline-flex items-center gap-1.5 rounded-md border border-[#27272a] bg-[#141417] px-3 py-1.5 text-[11px] font-semibold text-[#22c55e] hover:border-[#22c55e] disabled:opacity-50 cursor-pointer"
                >
                  <RefreshCw
                    className={`h-3.5 w-3.5 ${isRefreshingSql ? 'animate-spin' : ''}`}
                    aria-hidden="true"
                  />
                  Đồng bộ tin nhắn mới
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 9: KYUJIN RAW TEXT PASTE MODAL */}
      {kyujinRawModalOpen && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-black/75 p-4">
          <div className="max-h-[90vh] w-full max-w-2xl overflow-y-auto rounded-xl border border-[#27272a] bg-[#141417] p-5 text-xs text-[#edeeef] shadow-2xl">
            <div className="mb-3 flex items-center justify-between border-b border-[#27272a] pb-2.5">
              <div className="flex items-center gap-2">
                <FileText className="h-4 w-4 text-[#22c55e]" aria-hidden="true" />
                <h3 className="text-base font-bold text-[#edeeef]">
                  Dán văn bản Phiếu tuyển dụng (Kyujin / Baito / Nhiều đơn cùng lúc)
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setKyujinRawModalOpen(false)}
                className="rounded-md border border-[#27272a] bg-[#0c0c0e] px-3 py-1 text-xs font-semibold text-[#94949e] hover:text-[#edeeef] cursor-pointer"
              >
                Đóng
              </button>
            </div>
            <p className="mb-2 text-[11px] text-[#94949e]">
              Dán nội dung tuyển dụng (tiếng Nhật hoặc tiếng Việt, hỗ trợ nhiều đơn cách nhau bằng dòng trống hoặc ---). Hệ thống tự chia Lĩnh vực, Ngành nghề, Hình thức (Tokutei / Chính thức / Baito), lương giờ hoặc thực nhận và yêu cầu kinh nghiệm.
            </p>
            <textarea
              rows={9}
              value={kyujinRawText}
              onChange={(e) => setKyujinRawText(e.target.value)}
              placeholder={
                'VD:\nCông ty ABC - Tuyển Baito Konbini tại Tokyo (Shinjuku)\n時給: 1,250円 / 交通費支給 / 週3〜5日OK / N3 / 未経験OK...'
              }
              className="w-full rounded-lg border border-[#27272a] bg-[#0c0c0e] p-3 text-xs text-[#edeeef] placeholder:text-[#94949e] focus:border-[#22c55e] focus:outline-none"
            />
            <div className="mt-4 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setKyujinRawModalOpen(false)}
                className="rounded-md border border-[#27272a] px-4 py-2 font-semibold text-[#94949e] hover:text-[#edeeef] cursor-pointer"
              >
                Hủy
              </button>
              <button
                type="button"
                disabled={isImportingKyujin}
                onClick={() => void handleImportKyujinRawText()}
                className="rounded-md bg-[#22c55e] px-4 py-2 font-semibold text-[#0c0c0e] hover:opacity-90 disabled:opacity-60 cursor-pointer"
              >
                {isImportingKyujin ? 'Đang bóc tách...' : 'Bóc tách & Nhập Kyujin'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 10: RESIGNATION PAPERWORK REMINDER MODAL (FROM RIGHT SIDEBAR) */}
      {resignationReminderModal && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-black/75 p-4">
          <div className="max-h-[90vh] w-full max-w-2xl overflow-y-auto rounded-xl border border-[#27272a] bg-[#141417] p-5 text-xs text-[#edeeef] shadow-2xl">
            <div className="mb-3 flex items-center justify-between border-b border-[#27272a] pb-2.5">
              <h3 className="text-base font-bold text-[#edeeef]">
                Nhắn nhắc nộp giấy tờ nghỉ việc: {resignationReminderModal.fullName}
              </h3>
              <button
                type="button"
                onClick={() => setResignationReminderModal(null)}
                className="rounded-md border border-[#27272a] bg-[#0c0c0e] px-3 py-1 text-xs font-semibold text-[#94949e] hover:text-[#edeeef] cursor-pointer"
              >
                Đóng
              </button>
            </div>
            {(() => {
              const msgs = buildResignationReminderMessages(resignationReminderModal);
              return (
                <div className="space-y-3">
                  {[
                    { label: 'Tiếng Việt', text: msgs.vi },
                    { label: 'Tiếng Nhật (日本語)', text: msgs.jp },
                  ].map((item) => (
                    <div
                      key={item.label}
                      className="rounded-lg border border-[#27272a] bg-[#0c0c0e] p-3"
                    >
                      <div className="mb-1.5 flex items-center justify-between">
                        <span className="font-semibold text-[#22c55e]">{item.label}</span>
                        <button
                          type="button"
                          onClick={async () => {
                            const ok = await copyTextSafely(item.text);
                            showToast(
                              ok ? `📋 Đã sao chép tin nhắn (${item.label})` : '⚠️ Sao chép thất bại'
                            );
                          }}
                          className="rounded bg-[#22c55e] px-3 py-1 text-[11px] font-semibold text-[#0c0c0e] hover:opacity-90 cursor-pointer"
                        >
                          Sao chép
                        </button>
                      </div>
                      <pre className="whitespace-pre-wrap font-sans leading-relaxed text-[#edeeef]/90">
                        {item.text}
                      </pre>
                    </div>
                  ))}
                </div>
              );
            })()}
          </div>
        </div>
      )}

      {/* Floating Chatbot Launcher Button */}
      {!chatbotOpen && (
        <button
          type="button"
          onClick={() => setChatbotOpen(true)}
          aria-label="Mở Trợ lý AI Gemini"
          className="fixed bottom-5 right-5 z-40 inline-flex items-center gap-2 rounded-full bg-[#22c55e] px-4 py-3 text-xs font-semibold text-[#0c0c0e] shadow-lg transition-transform hover:scale-[1.02] hover:opacity-95 cursor-pointer"
        >
          <MessageSquareText className="h-4 w-4" aria-hidden="true" />
          Trợ lý AI
        </button>
      )}

      {/* Multi-turn Gemini Chatbot Drawer */}
      <GeminiChatbotDrawer
        isOpen={chatbotOpen}
        onClose={() => setChatbotOpen(false)}
        getIdToken={getFirebaseIdToken}
        userEmail={authUser?.email || null}
        isSignedIn={authStatus === 'signed_in'}
        crmSummary={crmSummaryForAi}
        onCopySuccess={(msg) => showToast(msg)}
      />
    </div>
  );
}
