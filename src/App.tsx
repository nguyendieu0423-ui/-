import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import { onAuthStateChanged, signOut, type User } from 'firebase/auth';
import {
  AlertTriangle,
  Bell,
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
  X,
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
  CRM_I18N,
  getCandidateDisplayName,
  getResignationDisplayName,
  readStoredUiLang,
  saveStoredUiLang,
  type UiLang,
} from './i18n/crmTranslations';
import {
  STORAGE_KEYS,
  SECTORS,
  EMPLOYMENT_TYPES,
  INDUSTRIES,
  KAIGO_FACILITY_TYPES,
  PROGRESS_STATUSES,
  PROPOSED_NAITEI_STATUSES,
  PROPOSED_JITAI_STATUSES,
  JLPT_LEVELS,
  CANDIDATE_NATIONALITIES,
  INITIAL_CANDIDATES,
  INITIAL_SCHOOL_STUDENTS,
  INITIAL_KYUJINS,
  INITIAL_PT_PARTNERS,
  INITIAL_RESIGNATIONS,
} from './data/seedData';
import type {
  Candidate,
  CandidateViewMode,
  CareFilterType,
  ChannelType,
  DailyFocusCompletionMap,
  DailyFocusItem,
  JitaiSubFilter,
  KyujinCompany,
  KyujinSlot,
  MainTabId,
  ProposedGroupFilter,
  PtPartner,
  ReplyScenario,
  ResignationRecord,
  RirekishoDraft,
  SavedRirekishoRecord,
} from './types/crm';
import {
  autoFixCandidatesNationalities,
  buildAnonymousJobPostText,
  buildRirekishoFallbackFromText,
  buildRirekishoStandaloneHtml,
  buildSmartReplyMessages,
  calcJapaneseAgeFromBirthDate,
  classifyProposedSubTab,
  cleanupAndReclassifyCandidates,
  copyTextSafely,
  deduplicateAndNormalizeKyujins,
  deduplicateCandidatesUnion,
  detectEmploymentTypeFromText,
  detectIndustryFromText,
  detectKyujinShiftType,
  detectSectorFromIndustryOrText,
  doesAreaMatchFilter,
  evaluateCandidateKyujinMatches,
  exportRirekishoToExcelFile,
  extractCandidateViaServer,
  extractDeclineReason,
  extractKyujinRawTextViaServer,
  extractKyujinViaServer,
  extractNormalizedKyujinConditions,
  extractNormalizedMendanInfo,
  extractRirekishoTextFromLocalFile,
  extractRirekishoViaServer,
  findDuplicateCandidate,
  formatPublicArea,
  getEffectiveProposedCompany,
  getInitialRirekishoDraftForCandidate,
  getMissingInfoForCandidate,
  getUrgentFollowUpReason,
  isActualProposedRecord,
  isStale7DayProposal,
  migrateCandidateRecord,
  migrateLegacyData,
  normalizeBonusInfo,
  normalizeCandidateNationality,
  normalizeHousingDisplay,
  normalizeKyujinCompanyName,
  normalizeKyujinItem,
  parseBackupCsvToCandidates,
  parseChatIntoUpdates,
  parseExcelFileToKyujinEntries,
  parseFlexibleDateToTimestamp,
  parseRawKyujinTextFallback,
  recognizeCandidateImageWithTesseract,
  sanitizeJobDescriptionForCandidate,
  stripKyujinPdfNoise,
} from './utils/crmHelpers';
import {
  compressSavedRirekishoInCandidates,
  loadCandidateCvsFromIdb,
  persistCandidatesWithIdbFallback,
  persistJsonWithQuotaCheck,
} from './utils/cvStorage';
import {
  buildCandidateCalendarEventDraft,
  buildGoogleCalendarTemplateUrl,
  buildResignationCalendarEventDraft,
  GAS_CALENDAR_APP_SNIPPET,
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
  getResignationCompanyLabel,
  hasUnconfirmedResignationDate,
  mergeResignationWithGasProtection,
  normalizeAndAutoComputeResignationDates,
  normalizeResignationNationality,
} from './utils/resignationHelpers';
import {
  extractKyujinSalaryInfo,
  getKyujinSlotSalaryDisplayMeta,
  isKyujinSalaryMissing,
  normalizeManualKyujinSalaryInput,
} from './utils/salaryParser';

const MESSENGER_WEBHOOK_URL = 'https://fb-crm.ai.studio/api/webhooks/messenger';
const FB_PAGE_LINK_STORAGE_KEY = 'tokutei_fb_page_link_v1';
const CHANNEL_CLEANUP_VERSION_KEY = 'tokutei_crm_channel_cleanup_v1';

function readFallbackCandidates(): Candidate[] {
  try {
    const saved = localStorage.getItem(STORAGE_KEYS.CANDIDATES);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed.map(migrateCandidateRecord);
      }
    }
  } catch {
    // ignore storage parse errors
  }
  return INITIAL_CANDIDATES.map(migrateCandidateRecord);
}

function readFallbackSchoolStudents(): Candidate[] {
  try {
    const saved = localStorage.getItem(STORAGE_KEYS.SCHOOL_STUDENTS);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed)) {
        return parsed.map(migrateCandidateRecord);
      }
    }
  } catch {
    // ignore storage parse errors
  }
  return INITIAL_SCHOOL_STUDENTS.map(migrateCandidateRecord);
}

function readFallbackKyujins(): KyujinCompany[] {
  try {
    const saved = localStorage.getItem(STORAGE_KEYS.KYUJINS);
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return deduplicateAndNormalizeKyujins(parsed);
      }
    }
  } catch {
    // ignore storage parse errors
  }
  return deduplicateAndNormalizeKyujins(INITIAL_KYUJINS);
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
        return parsed.map((item) =>
          normalizeAndAutoComputeResignationDates({
            ...item,
            nationality: normalizeResignationNationality(item?.nationality),
          })
        );
      }
    }
  } catch {
    // ignore storage parse errors
  }
  return INITIAL_RESIGNATIONS.map((item) => normalizeAndAutoComputeResignationDates(item));
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

  // UI Language & Navigation & Filters
  const [uiLang, setUiLang] = useState<UiLang>(() => readStoredUiLang());
  useEffect(() => {
    saveStoredUiLang(uiLang);
  }, [uiLang]);
  const t = CRM_I18N[uiLang];

  const [activeTab, setActiveTab] = useState<MainTabId>('all');
  const [candidateViewMode, setCandidateViewMode] = useState<CandidateViewMode>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEYS.VIEW_MODE);
      if (saved === 'card' || saved === 'list') return saved;
    } catch {
      // ignore storage errors
    }
    return 'list';
  });
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEYS.VIEW_MODE, candidateViewMode);
    } catch {
      // ignore storage errors
    }
  }, [candidateViewMode]);

  const [drawerCandidateId, setDrawerCandidateId] = useState<string | null>(null);
  const [naiteiPromptModal, setNaiteiPromptModal] = useState<{
    personId: string;
    isSchool: boolean;
    candidateName: string;
    targetStatus: string;
    naiteiCompany: string;
    expectedJoinDate: string;
    visaSubmissionDeadline: string;
  } | null>(null);
  const [jitaiPromptModal, setJitaiPromptModal] = useState<{
    personId: string;
    isSchool: boolean;
    candidateName: string;
    targetStatus: string;
    reason: string;
  } | null>(null);
  const [kyujinConditionModal, setKyujinConditionModal] = useState<{
    companyId: string;
    slotId: string;
    companyName: string;
    excludedNationalitiesText: string;
    excludedGenders: string[];
    requiresInJapan: boolean;
    requiresTokutei: boolean;
    extraRequiredText: string;
    extraNgText: string;
  } | null>(null);

  const [searchQuery, setSearchQuery] = useState('');
  const [careFilter, setCareFilter] = useState<CareFilterType>('all');
  const [statusFilter, setStatusFilter] = useState('all');
  const [industryFilter, setIndustryFilter] = useState('all');
  const [nationalityFilter, setNationalityFilter] = useState('all');
  const [proposedCompanyFilter, setProposedCompanyFilter] = useState('all');
  const [expandedJobMatchIds, setExpandedJobMatchIds] = useState<Record<string, boolean>>({});
  const [proposedMonthFilter, setProposedMonthFilter] = useState('all');
  const [proposedStageFilter, setProposedStageFilter] = useState<ProposedGroupFilter>('in_progress');
  const [jitaiSubFilter, setJitaiSubFilter] = useState<JitaiSubFilter>('all');
  const [declineReasonQuery, setDeclineReasonQuery] = useState('');
  const [kyujinSearch, setKyujinSearch] = useState('');
  const [kyujinStatusFilter, setKyujinStatusFilter] = useState<'all' | 'open' | 'closed'>('all');
  const [kyujinBonusOnly, setKyujinBonusOnly] = useState(false);
  const [kyujinMissingSalaryOnly, setKyujinMissingSalaryOnly] = useState(false);
  const [kyujinSectorFilter, setKyujinSectorFilter] = useState('all');
  const [kyujinIndustryFilter, setKyujinIndustryFilter] = useState('all');
  const [kyujinEmploymentTypeFilter, setKyujinEmploymentTypeFilter] = useState('all');
  const [kyujinFacilityFilter, setKyujinFacilityFilter] = useState('all');
  const [kyujinShiftTypeFilter, setKyujinShiftTypeFilter] = useState<
    'all' | '日勤のみ' | 'シフト制' | '夜勤あり'
  >('all');
  const [selectedSlotIds, setSelectedSlotIds] = useState<string[]>([]);
  const [editingKyujinSalary, setEditingKyujinSalary] = useState<{
    companyId: string;
    slotId: string;
    value: string;
  } | null>(null);
  const [editingKyujinField, setEditingKyujinField] = useState<{
    companyId: string;
    slotId: string;
    field: 'shiftTime' | 'jobCode' | 'area';
    value: string;
  } | null>(null);
  const salaryEditCommitLockRef = useRef<string | null>(null);
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
  const [currentImageBase64, setCurrentImageBase64] = useState<{
    mimeType: string;
    base64Data: string;
    dataUrl: string;
    fileName: string;
  } | null>(null);
  const [candidateExtractBanner, setCandidateExtractBanner] = useState<{
    type: 'loading' | 'success' | 'warning';
    text: string;
  } | null>(null);

  const editingPersonSessionKey = editingPerson
    ? `${editingPerson.isNew ? 'new' : 'edit'}:${editingPerson.data.id}`
    : '';

  useEffect(() => {
    setCurrentImageBase64(null);
    setCandidateExtractBanner(null);
  }, [editingPersonSessionKey]);
  const [fbChatModalOpen, setFbChatModalOpen] = useState(false);
  const [fbChatMode, setFbChatMode] = useState<'new' | 'update'>('new');
  const [fbChatChannel, setFbChatChannel] = useState<ChannelType>('personal');
  const [fbChatTargetId, setFbChatTargetId] = useState('');
  const [fbChatRawText, setFbChatRawText] = useState('');
  const [isExtractingFbChatAi, setIsExtractingFbChatAi] = useState(false);
  const [resignationReminderModal, setResignationReminderModal] =
    useState<ResignationRecord | null>(null);
  const [resignationCalendarModal, setResignationCalendarModal] = useState<{
    mode: 'single' | 'bulk';
    gasAttempted: boolean;
    items: Array<{
      record: ResignationRecord;
      draft: CalendarEventDraft;
      templateUrl: string;
    }>;
  } | null>(null);
  const [notificationPermission, setNotificationPermission] = useState<
    NotificationPermission | 'unsupported'
  >(() => {
    if (typeof window === 'undefined' || !('Notification' in window)) return 'unsupported';
    return window.Notification.permission;
  });
  const [dismissedUrgentBannerSig, setDismissedUrgentBannerSig] = useState('');
  const [periodicScanTick, setPeriodicScanTick] = useState(0);
  const notifiedResignationSigRef = useRef('');
  const [syncingResignationCalIds, setSyncingResignationCalIds] = useState<
    Record<string, boolean>
  >({});
  const [isBulkSyncingResignationsCal, setIsBulkSyncingResignationsCal] = useState(false);

  const [rirekishoModal, setRirekishoModal] = useState<{
    person: Candidate;
    draft: RirekishoDraft;
    sourceFileName?: string;
    exportedFileName?: string;
    exportedAt?: string;
    hasSavedVersion?: boolean;
    lastSavedAt?: string;
    showUploadTools?: boolean;
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

  // Toast Notification & Backup Restore Diagnostic Alert
  const [toastMsg, setToastMsg] = useState('');
  const [backupRestoreError, setBackupRestoreError] = useState<string | null>(null);
  const toastTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const backupInputRef = useRef<HTMLInputElement | null>(null);
  const kyujinFileInputRef = useRef<HTMLInputElement | null>(null);
  const rirekishoFileInputRef = useRef<HTMLInputElement | null>(null);
  const rirekishoPhotoInputRef = useRef<HTMLInputElement | null>(null);
  const rirekishoSavedSnapshotRef = useRef<string>('');
  const rirekishoAutoSaveTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

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

  const isValidSafeGasUrl = (rawUrl: string): boolean => {
    try {
      const parsed = new URL(rawUrl.trim());
      return parsed.protocol === 'https:';
    } catch {
      return false;
    }
  };

  const triggerSafeGoogleCalendarAnchor = (url: string) => {
    if (!isValidGoogleCalendarUrl(url) || typeof document === 'undefined') return;
    try {
      const a = document.createElement('a');
      a.href = url;
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
      a.style.display = 'none';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
    } catch {
      // ignore sandbox click restrictions; modal 1-click link remains visible
    }
  };

  const handleSyncResignationCalendarSingle = useCallback(
    async (record: ResignationRecord) => {
      if (record.paperworkStatus === 'Hoàn tất') {
        showToast('ℹ️ Hồ sơ này đã hoàn tất giấy tờ nên không cần đồng bộ lịch nhắc.');
        return;
      }

      const normalized = normalizeAndAutoComputeResignationDates(record);
      const draft = buildResignationCalendarEventDraft(normalized);
      if (!draft) {
        showToast(
          '⚠️ Vui lòng cập nhật Ngày nghỉ việc hoặc Hạn giấy tờ trước khi đồng bộ lịch.'
        );
        return;
      }

      const templateUrl = buildGoogleCalendarTemplateUrl(draft);
      const trimmedGasUrl = gasWebAppUrl.trim();

      setSyncingResignationCalIds((prev) => ({ ...prev, [normalized.id]: true }));
      try {
        // Save auto-computed deadline + direct 1-click Calendar link onto the record
        setResignations((prev) =>
          prev.map((r) =>
            r.id === normalized.id
              ? {
                  ...normalized,
                  googleCalendarEventLink: templateUrl,
                  updatedAt: new Date().toISOString(),
                }
              : r
          )
        );

        if (trimmedGasUrl && isValidSafeGasUrl(trimmedGasUrl)) {
          await fetch(trimmedGasUrl, {
            method: 'POST',
            mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              action: 'createCalendarEvent',
              useCalendarApp: true,
              entityType: 'resignation',
              recordId: normalized.id,
              title: draft.eventTitle,
              eventTitle: draft.eventTitle,
              date: draft.dateIso,
              startIso: draft.startIso,
              endIso: draft.endIso,
              timeZone: 'Asia/Tokyo',
              description: draft.description,
              reminderMinutes: [1440, 120],
            }),
          });
          showToast(
            `✅ Đã đẩy lịch "${draft.eventTitle}" qua GAS CalendarApp (ngày ${draft.dateIso}, nhắc 1 ngày & 2 giờ trước)`
          );
          return;
        }

        triggerSafeGoogleCalendarAnchor(templateUrl);
        setResignationCalendarModal({
          mode: 'single',
          gasAttempted: false,
          items: [{ record: normalized, draft, templateUrl }],
        });
        showToast(
          `📅 Đã tạo sẵn lịch Google Calendar (${draft.dateIso}) — bấm "Lưu lên Google Calendar" để hoàn tất trong 1 click!`
        );
      } catch {
        setResignationCalendarModal({
          mode: 'single',
          gasAttempted: Boolean(trimmedGasUrl),
          items: [{ record: normalized, draft, templateUrl }],
        });
        showToast(
          '📅 Đã chuyển sang link tạo lịch Google Calendar 1-click (không bị chặn OAuth).'
        );
      } finally {
        setSyncingResignationCalIds((prev) => ({ ...prev, [normalized.id]: false }));
      }
    },
    [gasWebAppUrl, showToast]
  );

  const handleBulkSyncResignationsCalendar = useCallback(async () => {
    const activeRecords = resignations.filter((r) => r.paperworkStatus !== 'Hoàn tất');
    if (activeRecords.length === 0) {
      showToast('ℹ️ Không có hồ sơ nghỉ việc nào đang chờ hạn nộp giấy tờ.');
      return;
    }

    const prepared = activeRecords
      .map((rec) => {
        const normalized = normalizeAndAutoComputeResignationDates(rec);
        const draft = buildResignationCalendarEventDraft(normalized);
        if (!draft) return null;
        return {
          record: normalized,
          draft,
          templateUrl: buildGoogleCalendarTemplateUrl(draft),
        };
      })
      .filter(
        (
          item
        ): item is {
          record: ResignationRecord;
          draft: CalendarEventDraft;
          templateUrl: string;
        } => Boolean(item)
      );

    if (prepared.length === 0) {
      showToast(
        '⚠️ Vui lòng cập nhật Ngày nghỉ việc hoặc Hạn giấy tờ trước khi đồng bộ lịch.'
      );
      return;
    }

    const trimmedGasUrl = gasWebAppUrl.trim();
    setIsBulkSyncingResignationsCal(true);
    try {
      const preparedById = new Map(prepared.map((p) => [p.record.id, p]));
      setResignations((prev) =>
        prev.map((r) => {
          const hit = preparedById.get(r.id);
          return hit
            ? {
                ...hit.record,
                googleCalendarEventLink: hit.templateUrl,
                updatedAt: new Date().toISOString(),
              }
            : r;
        })
      );

      if (trimmedGasUrl && isValidSafeGasUrl(trimmedGasUrl)) {
        await fetch(trimmedGasUrl, {
          method: 'POST',
          mode: 'no-cors',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            action: 'createBulkCalendarEvents',
            useCalendarApp: true,
            timeZone: 'Asia/Tokyo',
            reminderMinutes: [1440, 120],
            items: prepared.map(({ record, draft }) => ({
              entityType: 'resignation',
              recordId: record.id,
              title: draft.eventTitle,
              eventTitle: draft.eventTitle,
              date: draft.dateIso,
              startIso: draft.startIso,
              endIso: draft.endIso,
              description: draft.description,
              reminderMinutes: [1440, 120],
            })),
          }),
        });
        showToast(
          `✅ Đã đẩy toàn bộ ${prepared.length} lịch hạn giấy tờ qua GAS CalendarApp (nhắc trước 1 ngày & 2 giờ)!`
        );
        return;
      }

      setResignationCalendarModal({
        mode: 'bulk',
        gasAttempted: false,
        items: prepared,
      });
      showToast(
        `📅 Đã chuẩn bị ${prepared.length} link tạo lịch Google Calendar 1-click (nhắc trước 1 ngày & 2 giờ)!`
      );
    } catch {
      setResignationCalendarModal({
        mode: 'bulk',
        gasAttempted: Boolean(trimmedGasUrl),
        items: prepared,
      });
      showToast(
        `📅 Đã mở danh sách ${prepared.length} link Google Calendar 1-click.`
      );
    } finally {
      setIsBulkSyncingResignationsCal(false);
    }
  }, [resignations, gasWebAppUrl, showToast]);

  const handleSaveResignation = useCallback(
    async (record: ResignationRecord, isNew: boolean) => {
      let protectedRec: ResignationRecord = normalizeAndAutoComputeResignationDates(record);
      setResignations((prev) => {
        const existing = prev.find((r) => r.id === record.id) || null;
        protectedRec = applyManualProtectionOnResignationSave(
          existing,
          normalizeAndAutoComputeResignationDates({
            ...record,
            updatedAt: new Date().toISOString(),
          })
        );
        return isNew
          ? [protectedRec, ...prev]
          : prev.map((r) => (r.id === protectedRec.id ? protectedRec : r));
      });
      showToast(
        isNew
          ? uiLang === 'jp'
            ? '✅ 新規退職者データを登録しました'
            : '✅ Đã thêm hồ sơ nghỉ việc mới'
          : uiLang === 'jp'
          ? `✅ 退職者データを更新しました（GAS上書き保護ON）: ${getResignationDisplayName(protectedRec, uiLang)}`
          : `✅ Đã cập nhật hồ sơ (đã bật bảo vệ chống ghi đè GAS): ${getResignationDisplayName(protectedRec, uiLang)}`
      );
    },
    [executeCalendarDelete, executeCalendarSyncForDraft, showToast, uiLang]
  );

  const handleResignationStatusChange = useCallback(
    async (record: ResignationRecord, nextStatus: string) => {
      const updated: ResignationRecord = {
        ...record,
        paperworkStatus: nextStatus as ResignationRecord['paperworkStatus'],
        updatedAt: new Date().toISOString(),
      };
      setResignations((prev) => prev.map((r) => (r.id === record.id ? updated : r)));
      showToast(
        uiLang === 'jp'
          ? `✅ 書類ステータスを「${nextStatus}」に更新しました`
          : `✅ Đã chuyển trạng thái giấy tờ sang "${nextStatus}"`
      );

      if (nextStatus === 'Hoàn tất' && record.googleCalendarEventId) {
        void executeCalendarDelete(record.id, record.googleCalendarEventId, true, 'resignation');
      }
    },
    [executeCalendarDelete, showToast, uiLang]
  );

  const handleDeleteResignation = useCallback(
    async (record: ResignationRecord) => {
      const displayName = getResignationDisplayName(record, uiLang);
      setResignations((prev) => prev.filter((r) => r.id !== record.id));
      showToast(
        uiLang === 'jp'
          ? `🗑️ 退職者データを削除しました: ${displayName}`
          : `🗑️ Đã xóa hồ sơ nghỉ việc: ${displayName}`
      );
      if (record.googleCalendarEventId && auth.currentUser) {
        void executeCalendarDelete(record.id, record.googleCalendarEventId, false, 'resignation');
      }
    },
    [executeCalendarDelete, showToast, uiLang]
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
      showToast(
        uiLang === 'jp'
          ? `✅ Excel/CSVから退職者データ ${rows.length}件を取り込みました`
          : `✅ Đã nhập ${rows.length} hồ sơ nghỉ việc từ Excel/CSV`
      );
    },
    [gasWriteLocked, showToast, uiLang]
  );

  const handleBulkUpdateResignations = useCallback((updatedRecords: ResignationRecord[]) => {
    setResignations(
      updatedRecords.map((item) =>
        normalizeAndAutoComputeResignationDates({
          ...item,
          nationality: normalizeResignationNationality(item.nationality),
        })
      )
    );
  }, []);

  const handleConvertResignationToCandidate = useCallback(
    (record: ResignationRecord) => {
      const displayName = getResignationDisplayName(record, uiLang);
      const existing = candidates.find(
        (c) =>
          (record.convertedCandidateId && c.id === record.convertedCandidateId) ||
          (c.name.toLowerCase() === record.fullName.toLowerCase() &&
            c.memo?.includes(record.previousCompany))
      );
      if (existing) {
        setActiveTab('all');
        setSearchQuery(getCandidateDisplayName(existing, uiLang));
        showToast(
          uiLang === 'jp'
            ? `ℹ️ ${displayName} さんは既に求職者一覧に登録されています`
            : `ℹ️ ${displayName} đã có trong danh sách Ứng viên tìm việc mới`
        );
        return;
      }
      const newCand = convertResignationToCandidate(record);
      setCandidates((prev) => [newCand, ...prev]);
      setResignations((prev) =>
        prev.map((r) => (r.id === record.id ? { ...r, convertedCandidateId: newCand.id } : r))
      );
      setActiveTab('all');
      setEditingPerson({ data: newCand, isSchool: false, isNew: false });
      showToast(
        uiLang === 'jp'
          ? `✅ 「${displayName}」さんを新規求職者へ移行しました`
          : `✅ Đã chuyển "${displayName}" sang Ứng viên tìm việc mới`
      );
    },
    [candidates, showToast, uiLang]
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
            const migrated = data.candidates!.map(migrateCandidateRecord);
            const merged = mergePendingWebhookCandidates(
              migrated,
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
            const migrated = data.schoolStudents!.map(migrateCandidateRecord);
            const nextJson = JSON.stringify(migrated);
            if (JSON.stringify(prev) === nextJson) return prev;
            skipNextSqlSaveRef.current = true;
            return migrated;
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
                  return normalizeAndAutoComputeResignationDates(
                    mergeResignationWithGasProtection(
                      existing,
                      incoming as ResignationRecord & Record<string, unknown>,
                      true
                    )
                  );
                }
              }
              return normalizeAndAutoComputeResignationDates({
                ...incoming,
                nationality: normalizeResignationNationality(incoming.nationality),
              });
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
            ? data.candidates.map(migrateCandidateRecord)
            : [];
          const mergedWebhooks = mergePendingWebhookCandidates(
            loadedCandidates,
            data.pendingWebhookCandidates
          );
          const loadedStudents = Array.isArray(data.schoolStudents)
            ? data.schoolStudents.map(migrateCandidateRecord)
            : [];
          let needsCleanup = false;
          try {
            needsCleanup = localStorage.getItem(CHANNEL_CLEANUP_VERSION_KEY) !== 'done';
          } catch {
            needsCleanup = false;
          }
          if (needsCleanup) {
            const cleaned = cleanupAndReclassifyCandidates(mergedWebhooks, loadedStudents);
            setCandidates(cleaned.candidates);
            setSchoolStudents(cleaned.schoolStudents);
            try {
              localStorage.setItem(CHANNEL_CLEANUP_VERSION_KEY, 'done');
            } catch {
              // ignore
            }
          } else {
            setCandidates(mergedWebhooks);
            setSchoolStudents(loadedStudents);
          }
          setKyujins(
            Array.isArray(data.kyujins)
              ? deduplicateAndNormalizeKyujins(data.kyujins)
              : []
          );
          setPtPartners(Array.isArray(data.ptPartners) ? data.ptPartners : []);
          if (Array.isArray(data.resignations) && data.resignations.length > 0) {
            setResignations(
              data.resignations.map((r) =>
                normalizeAndAutoComputeResignationDates({
                  ...r,
                  nationality: normalizeResignationNationality(r.nationality),
                })
              )
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

  // Hydrate any offloaded CV / portrait photo payloads stored in IndexedDB on startup
  useEffect(() => {
    let cancelled = false;
    void loadCandidateCvsFromIdb().then((cvMap) => {
      if (cancelled || !cvMap || Object.keys(cvMap).length === 0) return;
      const mergeCv = (list: Candidate[]): Candidate[] => {
        let changed = false;
        const next = list.map((cand) => {
          const storedCv = cvMap[cand.id];
          if (!storedCv) return cand;
          const missingEntireCv = !cand.savedRirekisho;
          const missingPhoto =
            Boolean(storedCv.draft?.photoDataUrl) && !cand.savedRirekisho?.draft?.photoDataUrl;
          if (missingEntireCv) {
            changed = true;
            return { ...cand, savedRirekisho: storedCv };
          }
          if (missingPhoto && cand.savedRirekisho) {
            changed = true;
            return {
              ...cand,
              savedRirekisho: {
                ...cand.savedRirekisho,
                draft: {
                  ...cand.savedRirekisho.draft,
                  photoDataUrl: storedCv.draft.photoDataUrl,
                },
              },
            };
          }
          return cand;
        });
        return changed ? next : list;
      };
      setCandidates((prev) => mergeCv(prev));
      setSchoolStudents((prev) => mergeCv(prev));
    });
    return () => {
      cancelled = true;
    };
  }, []);

  const oneTimeChannelCleanupRanRef = useRef(false);
  useEffect(() => {
    if (oneTimeChannelCleanupRanRef.current) return;
    oneTimeChannelCleanupRanRef.current = true;
    try {
      if (localStorage.getItem(CHANNEL_CLEANUP_VERSION_KEY) === 'done') return;
    } catch {
      return;
    }
    const cleaned = cleanupAndReclassifyCandidates(candidates, schoolStudents);
    if (cleaned.changedCount > 0) {
      setCandidates(cleaned.candidates);
      setSchoolStudents(cleaned.schoolStudents);
    }
    try {
      localStorage.setItem(CHANNEL_CLEANUP_VERSION_KEY, 'done');
    } catch {
      // ignore storage errors
    }
  }, [candidates, schoolStudents]);

  useEffect(() => {
    void persistCandidatesWithIdbFallback(STORAGE_KEYS.CANDIDATES, candidates);
  }, [candidates]);

  useEffect(() => {
    void persistCandidatesWithIdbFallback(STORAGE_KEYS.SCHOOL_STUDENTS, schoolStudents);
  }, [schoolStudents]);

  useEffect(() => {
    persistJsonWithQuotaCheck(STORAGE_KEYS.KYUJINS, kyujins);
  }, [kyujins]);

  useEffect(() => {
    persistJsonWithQuotaCheck(STORAGE_KEYS.PT_PARTNERS, ptPartners);
  }, [ptPartners]);

  useEffect(() => {
    persistJsonWithQuotaCheck(STORAGE_KEYS.RESIGNATIONS, resignations);
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
  // RIREKISHO AUTO-SAVE, PHOTO, EXPORT & FILE/CHAT AUTO-FILL PARSER
  // ==========================================================================
  const persistCandidateRirekisho = useCallback(
    (
      person: Candidate,
      record: SavedRirekishoRecord,
      options?: { syncBasicProfile?: boolean }
    ) => {
      const isSchool = Boolean(person._isSchool || person.channel === 'school');
      const d = record.draft;
      const updater = (list: Candidate[]) =>
        list.map((c) => {
          if (c.id !== person.id) return c;
          const next: Candidate = {
            ...c,
            savedRirekisho: record,
          };
          if (options?.syncBasicProfile) {
            if (!next.furigana && d.furigana) next.furigana = d.furigana;
            if (!next.birthDate && d.birthDate) next.birthDate = d.birthDate;
            if (!next.age && d.age) next.age = d.age;
            if (!next.gender && d.gender) next.gender = d.gender;
            if (!next.postalCode && d.postalCode) next.postalCode = d.postalCode;
            if (!next.address && d.address) next.address = d.address;
            if (!next.phone && d.phone) next.phone = d.phone;
            if (!next.email && d.email) next.email = d.email;
            if ((!next.nationality || next.nationality === '国籍不明') && d.nationality) {
              next.nationality = d.nationality;
            }
            if (!next.visaExpiry && d.visaExpiry) next.visaExpiry = d.visaExpiry;
            if (!next.visaStatus && d.visaStatus) next.visaStatus = d.visaStatus;
          }
          return next;
        });

      if (isSchool) {
        setSchoolStudents(updater);
      } else {
        setCandidates(updater);
      }
      rirekishoSavedSnapshotRef.current = JSON.stringify(record.draft);
    },
    []
  );

  const openRirekishoModalForCandidate = useCallback((person: Candidate) => {
    const existing = person.savedRirekisho;
    const initialDraft = getInitialRirekishoDraftForCandidate(person);
    rirekishoSavedSnapshotRef.current = JSON.stringify(initialDraft);
    setRirekishoStatusBanner(null);
    setRirekishoChatInput(person.memo || '');
    setRirekishoModal({
      person,
      draft: initialDraft,
      sourceFileName: existing?.sourceFileName,
      exportedFileName: existing?.exportedFileName,
      exportedAt: existing?.exportedAt,
      hasSavedVersion: Boolean(existing?.draft),
      lastSavedAt: existing?.updatedAt,
      showUploadTools: false,
    });
  }, []);

  useEffect(() => {
    if (!rirekishoModal) {
      if (rirekishoAutoSaveTimerRef.current) {
        clearTimeout(rirekishoAutoSaveTimerRef.current);
        rirekishoAutoSaveTimerRef.current = null;
      }
      return;
    }

    const serialized = JSON.stringify(rirekishoModal.draft);
    if (serialized === rirekishoSavedSnapshotRef.current) return;

    if (rirekishoAutoSaveTimerRef.current) {
      clearTimeout(rirekishoAutoSaveTimerRef.current);
    }

    const activeSnapshot = rirekishoModal;
    rirekishoAutoSaveTimerRef.current = setTimeout(() => {
      rirekishoAutoSaveTimerRef.current = null;
      const nowIso = new Date().toISOString();
      const record: SavedRirekishoRecord = {
        draft: activeSnapshot.draft,
        sourceFileName: activeSnapshot.sourceFileName,
        exportedFileName: activeSnapshot.exportedFileName,
        exportedAt: activeSnapshot.exportedAt,
        updatedAt: nowIso,
      };
      persistCandidateRirekisho(activeSnapshot.person, record);
      setRirekishoModal((prev) =>
        prev && prev.person.id === activeSnapshot.person.id
          ? {
              ...prev,
              hasSavedVersion: true,
              lastSavedAt: nowIso,
              person: { ...prev.person, savedRirekisho: record },
            }
          : prev
      );
    }, 550);

    return () => {
      if (rirekishoAutoSaveTimerRef.current) {
        clearTimeout(rirekishoAutoSaveTimerRef.current);
        rirekishoAutoSaveTimerRef.current = null;
      }
    };
  }, [rirekishoModal, persistCandidateRirekisho]);

  const handleSaveRirekishoManual = () => {
    if (!rirekishoModal) return;
    if (rirekishoAutoSaveTimerRef.current) {
      clearTimeout(rirekishoAutoSaveTimerRef.current);
      rirekishoAutoSaveTimerRef.current = null;
    }
    const nowIso = new Date().toISOString();
    const record: SavedRirekishoRecord = {
      draft: rirekishoModal.draft,
      sourceFileName: rirekishoModal.sourceFileName,
      exportedFileName: rirekishoModal.exportedFileName,
      exportedAt: rirekishoModal.exportedAt,
      updatedAt: nowIso,
    };
    persistCandidateRirekisho(rirekishoModal.person, record, { syncBasicProfile: true });
    setRirekishoModal((prev) =>
      prev
        ? {
            ...prev,
            hasSavedVersion: true,
            lastSavedAt: nowIso,
            person: { ...prev.person, savedRirekisho: record },
          }
        : prev
    );
    setRirekishoStatusBanner({ type: 'success', text: t.rirekisho.savedToast });
    showToast(t.rirekisho.savedToast);
  };

  const handleCloseRirekishoModal = () => {
    if (rirekishoAutoSaveTimerRef.current) {
      clearTimeout(rirekishoAutoSaveTimerRef.current);
      rirekishoAutoSaveTimerRef.current = null;
    }
    if (rirekishoModal) {
      const serialized = JSON.stringify(rirekishoModal.draft);
      if (serialized !== rirekishoSavedSnapshotRef.current) {
        const nowIso = new Date().toISOString();
        persistCandidateRirekisho(rirekishoModal.person, {
          draft: rirekishoModal.draft,
          sourceFileName: rirekishoModal.sourceFileName,
          exportedFileName: rirekishoModal.exportedFileName,
          exportedAt: rirekishoModal.exportedAt,
          updatedAt: nowIso,
        });
      }
    }
    setRirekishoModal(null);
  };

  const handleUploadRirekishoPhoto = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file || !rirekishoModal) return;
    try {
      const compressed = await compressImageFileToBase64(file, 480, 0.8);
      const photoDataUrl = `data:${compressed.mimeType};base64,${compressed.base64Data}`;
      const nextDraft: RirekishoDraft = {
        ...rirekishoModal.draft,
        photoDataUrl,
      };
      const nowIso = new Date().toISOString();
      const record: SavedRirekishoRecord = {
        draft: nextDraft,
        sourceFileName: rirekishoModal.sourceFileName,
        exportedFileName: rirekishoModal.exportedFileName,
        exportedAt: rirekishoModal.exportedAt,
        updatedAt: nowIso,
      };
      persistCandidateRirekisho(rirekishoModal.person, record);
      setRirekishoModal((prev) =>
        prev
          ? {
              ...prev,
              draft: nextDraft,
              hasSavedVersion: true,
              lastSavedAt: nowIso,
              person: { ...prev.person, savedRirekisho: record },
            }
          : prev
      );
      showToast(t.rirekisho.photoUploadedToast);
    } catch {
      showToast(
        uiLang === 'jp'
          ? '⚠️ 証明写真の読み込みに失敗しました'
          : '⚠️ Không thể đọc file ảnh chân dung'
      );
    }
  };

  const applyExtractedRirekishoToDraft = (
    extracted: Record<string, unknown> = {},
    sourceFileName?: string
  ) => {
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

      return {
        ...prev,
        draft: d,
        sourceFileName: sourceFileName || prev.sourceFileName,
      };
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
        const filledCount = applyExtractedRirekishoToDraft(extracted, file.name);
        if (filledCount > 0) {
          const msg = `✅ Đã bóc tách ${filledCount} mục từ "${file.name}" và tự động lưu vào Sơ yếu lý lịch 2 cột!`;
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
            localParsed as Record<string, unknown>,
            file.name
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
          localParsed as Record<string, unknown>,
          file.name
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

  const recordRirekishoExport = (exportedFileName: string) => {
    if (!rirekishoModal) return;
    const nowIso = new Date().toISOString();
    const record: SavedRirekishoRecord = {
      draft: rirekishoModal.draft,
      sourceFileName: rirekishoModal.sourceFileName,
      exportedFileName,
      exportedAt: nowIso,
      updatedAt: nowIso,
    };
    persistCandidateRirekisho(rirekishoModal.person, record);
    setRirekishoModal((prev) =>
      prev
        ? {
            ...prev,
            hasSavedVersion: true,
            exportedFileName,
            exportedAt: nowIso,
            lastSavedAt: nowIso,
            person: { ...prev.person, savedRirekisho: record },
          }
        : prev
    );
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
      const fileName = `履歴書_${safeName}_${dateTag}.html`;
      const blob = new Blob(['\uFEFF' + htmlContent], {
        type: 'text/html;charset=utf-8',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      recordRirekishoExport(fileName);
      showToast(`📥 「${fileName}」をダウンロード＆保存しました`);
    } catch {
      showToast('⚠️ 履歴書ファイルのダウンロードに失敗しました');
    }
  };

  const handleExportRirekishoExcel = () => {
    if (!rirekishoModal?.draft) return;
    try {
      const fileName = exportRirekishoToExcelFile(rirekishoModal.draft);
      recordRirekishoExport(fileName);
      showToast(t.rirekisho.excelExportedToast(fileName));
    } catch {
      showToast(
        uiLang === 'jp'
          ? '⚠️ Excel履歴書の出力に失敗しました'
          : '⚠️ Lỗi xuất file Excel Rirekisho'
      );
    }
  };

  const handlePrintRirekishoSafely = () => {
    if (!rirekishoModal?.draft) return;
    const safeName = String(rirekishoModal.draft.fullName || '応募者')
      .replace(/[\\/:*?"<>|\s]+/g, '_')
      .trim();
    const dateTag = new Date().toISOString().slice(0, 10);
    recordRirekishoExport(`履歴書_${safeName}_${dateTag}.pdf`);
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
    setBackupRestoreError(null);
    try {
      const rawText = await file.text();
      // Strip UTF-8 BOM and NUL characters commonly left by older exports
      const cleanText = rawText.replace(/^\uFEFF/, '').replace(/\u0000/g, '').trim();
      if (!cleanText) {
        throw new Error(
          uiLang === 'jp'
            ? '選択されたファイルは空です（データが含まれていません）。'
            : 'File backup được chọn đang rỗng (không có nội dung).'
        );
      }

      const isCsvFile =
        file.name.toLowerCase().endsWith('.csv') ||
        (!cleanText.startsWith('{') && !cleanText.startsWith('['));

      if (isCsvFile) {
        const migratedRows = parseBackupCsvToCandidates(cleanText);
        if (migratedRows.length === 0) {
          throw new Error(
            uiLang === 'jp'
              ? 'CSVファイルに有効な相談者データが見つかりません（ヘッダー行またはデータ行が空です）。'
              : 'Không tìm thấy dữ liệu ứng viên hợp lệ trong file CSV (thiếu dòng tiêu đề hoặc dòng dữ liệu).'
          );
        }
        const csvStudents: Candidate[] = [];
        const csvCandidates: Candidate[] = [];
        migratedRows.forEach((item) => {
          if (item.channel === 'school' || item.schoolName) {
            csvStudents.push({ ...item, channel: 'school', _isSchool: true });
          } else {
            csvCandidates.push({ ...item, _isSchool: false });
          }
        });
        const csvCleaned = cleanupAndReclassifyCandidates(csvCandidates, csvStudents);
        if (csvCleaned.candidates.length > 0) {
          setCandidates(csvCleaned.candidates);
          await persistCandidatesWithIdbFallback(STORAGE_KEYS.CANDIDATES, csvCleaned.candidates);
        }
        if (csvCleaned.schoolStudents.length > 0) {
          setSchoolStudents(csvCleaned.schoolStudents);
          await persistCandidatesWithIdbFallback(STORAGE_KEYS.SCHOOL_STUDENTS, csvCleaned.schoolStudents);
        }
        setGasWriteLocked(true);
        showToast(
          uiLang === 'jp'
            ? `📂 CSVデータ復元完了：相談者 ${csvCleaned.candidates.length}名 ／ 留学生 ${csvCleaned.schoolStudents.length}名（新構造へ自動変換済）`
            : `📂 Đã khôi phục & chuẩn hóa CSV: ${csvCleaned.candidates.length} ứng viên / ${csvCleaned.schoolStudents.length} DHS`
        );
        return;
      }

      let parsed: unknown;
      try {
        parsed = JSON.parse(cleanText);
      } catch (jsonErr) {
        const detail = jsonErr instanceof Error ? jsonErr.message : String(jsonErr);
        throw new Error(
          uiLang === 'jp'
            ? `JSON構文エラーのためファイルを読み込めませんでした（${detail}）。`
            : `Lỗi cú pháp JSON không thể đọc file (${detail}).`
        );
      }

      const migrated = migrateLegacyData(parsed);
      const cleanedBackup =
        migrated.hasCandidatesSection || migrated.hasStudentsSection
          ? cleanupAndReclassifyCandidates(migrated.candidates, migrated.schoolStudents)
          : null;

      const compressedCandidates =
        migrated.hasCandidatesSection && cleanedBackup
          ? await compressSavedRirekishoInCandidates(cleanedBackup.candidates)
          : [];
      const compressedStudents =
        (migrated.hasStudentsSection || (cleanedBackup && cleanedBackup.schoolStudents.length > 0)) &&
        cleanedBackup
          ? await compressSavedRirekishoInCandidates(cleanedBackup.schoolStudents)
          : [];

      let usedIndexedDbOffload = false;
      const storageErrors: string[] = [];

      if (migrated.hasCandidatesSection) {
        setCandidates(compressedCandidates);
        const persistRes = await persistCandidatesWithIdbFallback(
          STORAGE_KEYS.CANDIDATES,
          compressedCandidates
        );
        if (persistRes.mode === 'indexeddb_offload') usedIndexedDbOffload = true;
        if (persistRes.mode === 'error' && persistRes.errorMessage) {
          storageErrors.push(persistRes.errorMessage);
        }
      }

      if (migrated.hasStudentsSection || (cleanedBackup && cleanedBackup.schoolStudents.length > 0)) {
        setSchoolStudents(compressedStudents);
        const persistStuRes = await persistCandidatesWithIdbFallback(
          STORAGE_KEYS.SCHOOL_STUDENTS,
          compressedStudents
        );
        if (persistStuRes.mode === 'indexeddb_offload') usedIndexedDbOffload = true;
        if (persistStuRes.mode === 'error' && persistStuRes.errorMessage) {
          storageErrors.push(persistStuRes.errorMessage);
        }
      }

      if (migrated.hasKyujinsSection) {
        setKyujins(migrated.kyujins);
        const kjPersist = persistJsonWithQuotaCheck(STORAGE_KEYS.KYUJINS, migrated.kyujins);
        if (!kjPersist.ok && kjPersist.errorMessage) {
          storageErrors.push(`Kyujins: ${kjPersist.errorMessage}`);
        }
      }

      if (migrated.hasPtSection) {
        setPtPartners(migrated.ptPartners);
        persistJsonWithQuotaCheck(STORAGE_KEYS.PT_PARTNERS, migrated.ptPartners);
      }

      if (migrated.hasResignationsSection) {
        setResignations((prev) => {
          const prevById = new Map(prev.map((item) => [item.id, item]));
          return migrated.resignations.map((rawRow) =>
            normalizeAndAutoComputeResignationDates(
              mergeResignationWithGasProtection(
                rawRow.id ? prevById.get(String(rawRow.id)) : undefined,
                {
                  ...rawRow,
                  nationality: normalizeResignationNationality(
                    typeof rawRow.nationality === 'string' ? rawRow.nationality : ''
                  ),
                },
                gasWriteLocked
              )
            )
          );
        });
      }

      if (storageErrors.length > 0) {
        throw new Error(
          uiLang === 'jp'
            ? `ブラウザの保存容量上限（localStorage）を超えました：${storageErrors.join(' / ')}`
            : `Vượt quá giới hạn lưu trữ trình duyệt (localStorage): ${storageErrors.join(' / ')}`
        );
      }

      setGasWriteLocked(true);
      const candCount = migrated.hasCandidatesSection
        ? compressedCandidates.length
        : candidates.length;
      const stuCount = migrated.hasStudentsSection
        ? compressedStudents.length
        : schoolStudents.length;
      const kyujinCount = migrated.hasKyujinsSection
        ? migrated.kyujins.length
        : kyujins.length;
      const idbNote = usedIndexedDbOffload
        ? uiLang === 'jp'
          ? ' ・大容量CV/写真はIndexedDBへ退避済'
          : ' (Ảnh/CV dung lượng lớn đã tự động lưu sang IndexedDB)'
        : '';

      showToast(
        uiLang === 'jp'
          ? `📂 データ復元完了（自動移行済）：相談者 ${candCount}名 ／ 留学生 ${stuCount}名 ／ 求人 ${kyujinCount}社${idbNote}`
          : `📂 Khôi phục & chuẩn hóa xong: ${candCount} ứng viên / ${stuCount} DHS / ${kyujinCount} công ty${idbNote}`
      );
    } catch (err) {
      const errorDetail =
        err instanceof Error ? err.message : String(err || 'Lỗi không xác định');
      setBackupRestoreError(errorDetail);
      showToast(
        uiLang === 'jp'
          ? `⚠️ データ復元エラー：${errorDetail}`
          : `⚠️ Lỗi khôi phục backup: ${errorDetail}`
      );
    } finally {
      e.target.value = '';
    }
  };

  // ==========================================================================
  // DERIVED COLLECTIONS
  // ==========================================================================
  const allPeopleCombined = useMemo<Candidate[]>(
    () => deduplicateCandidatesUnion(candidates, schoolStudents),
    [candidates, schoolStudents]
  );

  const schoolCandidates = useMemo(
    () => allPeopleCombined.filter((p) => Boolean(p._isSchool || p.channel === 'school')),
    [allPeopleCombined]
  );

  const nonSchoolCandidates = useMemo(
    () => allPeopleCombined.filter((p) => !p._isSchool && p.channel !== 'school'),
    [allPeopleCombined]
  );

  const personalCandidates = useMemo(
    () =>
      nonSchoolCandidates.filter(
        (c) =>
          c.channel !== 'company' &&
          c.channel !== 'line' &&
          c.channel !== 'other'
      ),
    [nonSchoolCandidates]
  );

  const companyCandidates = useMemo(
    () => nonSchoolCandidates.filter((c) => c.channel === 'company'),
    [nonSchoolCandidates]
  );

  const lineCandidates = useMemo(
    () => nonSchoolCandidates.filter((c) => c.channel === 'line'),
    [nonSchoolCandidates]
  );

  const otherCandidates = useMemo(
    () => nonSchoolCandidates.filter((c) => c.channel === 'other'),
    [nonSchoolCandidates]
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
      in_progress: 0,
      naitei: 0,
      jitai: 0,
      jitaiBreakdown: {
        内定辞退: 0,
        選考中辞退: 0,
        不合格: 0,
        '辞退・終了': 0,
      },
    };
    allProposedPeople.forEach((p) => {
      const grp = classifyProposedSubTab(p);
      counts[grp] += 1;
      if (grp === 'jitai') {
        const st = String(p.progressStatus || '').trim();
        if (st === '内定辞退') counts.jitaiBreakdown['内定辞退'] += 1;
        else if (st === '選考中辞退') counts.jitaiBreakdown['選考中辞退'] += 1;
        else if (st === '不合格') counts.jitaiBreakdown['不合格'] += 1;
        else counts.jitaiBreakdown['辞退・終了'] += 1;
      }
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
        if ([...PROPOSED_NAITEI_STATUSES, ...PROPOSED_JITAI_STATUSES].includes(st)) return false;
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
      .map((person) => ({ person, reason: getUrgentFollowUpReason(person, uiLang) }))
      .filter((item) => Boolean(item.reason))
      .slice(0, 10);
  }, [allPeopleCombined, uiLang]);

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
    void periodicScanTick;
    return resignations
      .map((rawRecord) => {
        const record = normalizeAndAutoComputeResignationDates(rawRecord);
        return {
          record,
          assessment: assessResignationDeadline(
            record.paperworkDeadline,
            record.paperworkStatus,
            uiLang
          ),
        };
      })
      .filter(
        ({ assessment }) =>
          assessment.urgency === 'overdue' ||
          assessment.urgency === 'urgent3' ||
          assessment.urgency === 'warn7'
      )
      .sort((a, b) => (a.assessment.daysLeft ?? 999) - (b.assessment.daysLeft ?? 999));
  }, [resignations, periodicScanTick, uiLang]);

  const urgentResignationAlerts = useMemo(
    () =>
      resignationAlerts.filter(
        ({ assessment }) =>
          assessment.urgency === 'overdue' || assessment.urgency === 'urgent3'
      ),
    [resignationAlerts]
  );

  const urgentResignationSignature = useMemo(
    () =>
      urgentResignationAlerts
        .map(
          ({ record, assessment }) =>
            `${record.id}:${record.paperworkDeadline}:${assessment.daysLeft}`
        )
        .join('|'),
    [urgentResignationAlerts]
  );

  // Periodic scan every 5 minutes + automatic browser notification for deadlines <= 3 days
  useEffect(() => {
    const timer = setInterval(() => {
      setPeriodicScanTick((t) => t + 1);
    }, 5 * 60 * 1000);
    return () => clearInterval(timer);
  }, []);

  useEffect(() => {
    if (!urgentResignationSignature || urgentResignationAlerts.length === 0) return;
    if (
      typeof window !== 'undefined' &&
      'Notification' in window &&
      window.Notification.permission === 'granted' &&
      notifiedResignationSigRef.current !== urgentResignationSignature
    ) {
      notifiedResignationSigRef.current = urgentResignationSignature;
      const lines = urgentResignationAlerts.slice(0, 3).map(({ record, assessment }) => {
        const displayName = getResignationDisplayName(record, uiLang);
        if (assessment.daysLeft !== null && assessment.daysLeft < 0) {
          return uiLang === 'jp'
            ? `• ${displayName}: ${Math.abs(assessment.daysLeft)}日超過`
            : `• ${displayName}: quá hạn ${Math.abs(assessment.daysLeft)} ngày`;
        }
        if (assessment.daysLeft === 0) {
          return uiLang === 'jp'
            ? `• ${displayName}: 提出期限 本日`
            : `• ${displayName}: hạn nộp giấy tờ HÔM NAY`;
        }
        return uiLang === 'jp'
          ? `• ${displayName}: 残り${assessment.daysLeft}日 (${record.paperworkDeadline})`
          : `• ${displayName}: còn ${assessment.daysLeft} ngày (${record.paperworkDeadline})`;
      });
      try {
        new window.Notification(
          t.cornerResigTitle(urgentResignationAlerts.length),
          { body: lines.join('\n') }
        );
      } catch {
        // ignore unsupported Notification constructor contexts
      }
    }
  }, [urgentResignationAlerts, urgentResignationSignature, uiLang, t]);

  const handleRequestBrowserNotifications = useCallback(async () => {
    if (typeof window === 'undefined' || !('Notification' in window)) {
      setNotificationPermission('unsupported');
      showToast(
        uiLang === 'jp'
          ? 'ℹ️ このブラウザはWeb通知に対応していません。'
          : 'ℹ️ Trình duyệt hiện tại không hỗ trợ Web Notification.'
      );
      return;
    }
    try {
      const perm = await window.Notification.requestPermission();
      setNotificationPermission(perm);
      if (perm === 'granted') {
        showToast(
          uiLang === 'jp'
            ? '✅ 退職書類リマインドのブラウザ通知を有効にしました！'
            : '✅ Đã bật thông báo trình duyệt cho nhắc lịch giấy tờ nghỉ việc!'
        );
      } else {
        showToast(
          uiLang === 'jp'
            ? 'ℹ️ ブラウザ通知は未許可です（画面右下のバナーは引き続き表示されます）。'
            : 'ℹ️ Bạn chưa cấp quyền thông báo trình duyệt (vẫn có cảnh báo góc màn hình).'
        );
      }
    } catch {
      showToast(
        uiLang === 'jp'
          ? '⚠️ この環境では通知権限をリクエストできません。'
          : '⚠️ Không thể yêu cầu quyền thông báo trên trình duyệt này.'
      );
    }
  }, [showToast, uiLang]);

  const dailyFocusItems = useMemo<DailyFocusItem[]>(() => {
    const items: DailyFocusItem[] = [];
    const todayYmd = getJstTodayYmd();
    const [ty, tm, td] = todayYmd.split('-').map(Number);
    const todayUtc = Date.UTC(ty, tm - 1, td);
    const dfLabels = t.dailyFocusLabels;

    // 0. URGENT (<= 3 days or overdue) Resignation paperwork deadlines pinned at VERY TOP
    urgentResignationAlerts.forEach(({ record, assessment }) => {
      const displayName = getResignationDisplayName(record, uiLang);
      const compLabel = getResignationCompanyLabel(record.previousCompany, uiLang);
      const urgentTitle =
        assessment.daysLeft !== null && assessment.daysLeft < 0
          ? dfLabels.overdue(Math.abs(assessment.daysLeft), displayName)
          : assessment.daysLeft === 0
          ? dfLabels.dueToday(displayName)
          : dfLabels.dueInDays(assessment.daysLeft ?? 0, displayName);
      items.push({
        id: `focus_resig_${record.id}`,
        category: 'resignation',
        title: urgentTitle,
        subtitle: dfLabels.resigSub(
          record.paperworkDeadline || '-',
          record.resignationDate || '-',
          compLabel
        ),
        urgency: 'critical',
        targetTab: 'resignation',
        entityId: record.id,
      });
    });

    // 0b. Unconfirmed resignation dates (退職日未確定)
    resignations
      .map((r) => normalizeAndAutoComputeResignationDates(r))
      .filter(hasUnconfirmedResignationDate)
      .slice(0, 5)
      .forEach((record) => {
        const displayName = getResignationDisplayName(record, uiLang);
        const compLabel = getResignationCompanyLabel(record.previousCompany, uiLang);
        items.push({
          id: `focus_resig_unconf_${record.id}`,
          category: 'resignation_unconfirmed',
          title:
            uiLang === 'jp'
              ? `⚠️ 退職日未確定・要確認: ${displayName}`
              : `⚠️ Chưa rõ ngày nghỉ việc: ${displayName}`,
          subtitle:
            uiLang === 'jp'
              ? `${compLabel} · 本人へ最終出社日を確認してください`
              : `${compLabel} · Cần nhắn xác nhận ngày nghỉ chính thức`,
          urgency: 'warning',
          targetTab: 'resignation',
          entityId: record.id,
        });
      });

    // 0c. Stale proposed candidates (> 7 days in 企業提案中 / 面接待ち)
    allPeopleCombined
      .filter((p) => isStale7DayProposal(p))
      .slice(0, 5)
      .forEach((p) => {
        const displayName = getCandidateDisplayName(p, uiLang);
        const comp = getEffectiveProposedCompany(p) || '-';
        items.push({
          id: `focus_stale_prop_${p.id}`,
          category: 'stale_proposed',
          title:
            uiLang === 'jp'
              ? `⏰ 推薦後7日経過・要確認: ${displayName}`
              : `⏰ Quá 7 ngày chưa có KQ: ${displayName}`,
          subtitle: `${comp} · ${p.proposalDate || p.lastContactDate || '-'}`,
          urgency: 'critical',
          targetTab: 'proposed',
          entityId: p.id,
        });
      });

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
          const displayName = getCandidateDisplayName(p, uiLang);
          items.push({
            id: `focus_mendan_${p.id}_${ymd}`,
            category: 'mendan',
            title: `${diffDays === 0 ? dfLabels.today : dfLabels.tomorrow}: ${displayName}`,
            subtitle:
              p.interviewDate || mInfo.dateLabel || p.progressStatus || dfLabels.mendanFallback,
            urgency: diffDays === 0 ? 'critical' : 'warning',
            targetTab: p._isSchool ? 'school' : 'mendan',
            entityId: p.id,
          });
        }
      }
    });

    // 2. Remaining Resignation paperwork deadlines (4..7 days)
    resignationAlerts
      .filter(({ assessment }) => assessment.urgency === 'warn7')
      .forEach(({ record, assessment }) => {
        const displayName = getResignationDisplayName(record, uiLang);
        const compLabel = getResignationCompanyLabel(record.previousCompany, uiLang);
        items.push({
          id: `focus_resig_${record.id}`,
          category: 'resignation',
          title: dfLabels.resigWarn(displayName),
          subtitle: `${assessment.label} · ${compLabel}`,
          urgency: 'warning',
          targetTab: 'resignation',
          entityId: record.id,
        });
      });

    // 3. Uncontacted > 3 days
    urgentFollowUpList.slice(0, 5).forEach(({ person, reason }) => {
      const displayName = getCandidateDisplayName(person, uiLang);
      items.push({
        id: `focus_uncontacted_${person.id}`,
        category: 'uncontacted',
        title: dfLabels.followup(displayName),
        subtitle: reason || dfLabels.uncontactedFallback,
        urgency: 'warning',
        targetTab: person._isSchool ? 'school' : 'all',
        entityId: person.id,
      });
    });

    // 3b. Naitei Visa Application Submission Deadline (<= 7 days / overdue)
    allPeopleCombined.forEach((p) => {
      const deadlineStr = String(p.visaSubmissionDeadline || '').trim();
      const st = String(p.progressStatus || '').trim();
      if (!deadlineStr || st === '入社済み' || PROPOSED_JITAI_STATUSES.includes(st)) return;
      const dlTs = parseFlexibleDateToTimestamp(deadlineStr, ty);
      if (dlTs > 0) {
        const d = new Date(dlTs + 9 * 60 * 60 * 1000);
        const ymd = d.toISOString().slice(0, 10);
        const [py, pm, pd] = ymd.split('-').map(Number);
        const diffDays = Math.round((Date.UTC(py, pm - 1, pd) - todayUtc) / 86400000);
        if (diffDays <= 7) {
          const displayName = getCandidateDisplayName(p, uiLang);
          const compLabel = p.naiteiCompany || getEffectiveProposedCompany(p) || '';
          const joinLabel = p.expectedJoinDate ? ` · ${p.expectedJoinDate}` : '';
          items.push({
            id: `focus_visa_sub_${p.id}`,
            category: 'visa_submission',
            title:
              diffDays < 0
                ? t.visaSubLabels.overdue(Math.abs(diffDays), displayName)
                : diffDays === 0
                ? t.visaSubLabels.dueToday(displayName)
                : t.visaSubLabels.dueInDays(diffDays, displayName),
            subtitle: `${deadlineStr}${compLabel ? ` · ${compLabel}` : ''}${joinLabel}`,
            urgency: diffDays <= 2 ? 'critical' : 'warning',
            targetTab: 'proposed',
            entityId: p.id,
          });
        }
      }
    });

    // 4. Residence card expiry < 90 days
    allPeopleCombined.forEach((p) => {
      if (!p.visaExpiry) return;
      const expTs = parseFlexibleDateToTimestamp(p.visaExpiry, ty);
      if (expTs > 0) {
        const diffDays = Math.round((expTs - todayUtc) / 86400000);
        if (diffDays >= -30 && diffDays <= 90) {
          const displayName = getCandidateDisplayName(p, uiLang);
          items.push({
            id: `focus_visa_${p.id}`,
            category: 'visa_expiry',
            title: dfLabels.visaExp(displayName),
            subtitle:
              diffDays < 0
                ? dfLabels.visaExpired(Math.abs(diffDays), p.visaExpiry)
                : dfLabels.visaRemain(diffDays, p.visaExpiry),
            urgency: diffDays <= 30 ? 'critical' : 'warning',
            targetTab: p._isSchool ? 'school' : 'all',
            entityId: p.id,
          });
        }
      }
    });

    return items.slice(0, 20);
  }, [
    allPeopleCombined,
    resignations,
    resignationAlerts,
    urgentResignationAlerts,
    urgentFollowUpList,
    t.dailyFocusLabels,
    t.visaSubLabels,
    uiLang,
  ]);

  const toggleDailyFocusItem = useCallback((id: string) => {
    setDailyFocusCompleted((prev) => ({
      ...prev,
      [id]: !prev[id],
    }));
  }, []);

  const industryFilterOptions = useMemo(() => {
    const set = new Set<string>();
    allPeopleCombined.forEach((p) => {
      const val = String(p.desiredIndustry || '').trim();
      if (val && val !== '未設定' && val !== 'Chưa rõ') {
        val
          .split(/[/・、,]+/)
          .map((s) => s.trim())
          .filter(Boolean)
          .forEach((part) => set.add(part));
      }
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [allPeopleCombined]);

  const nationalityFilterOptions = useMemo(() => {
    const set = new Set<string>();
    allPeopleCombined.forEach((p) => {
      const val = String(p.nationality || '').trim();
      if (val && val !== 'undefined' && val !== 'null') {
        set.add(val);
      }
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [allPeopleCombined]);

  const proposedCompanyFilterOptions = useMemo(() => {
    const set = new Set<string>();
    allPeopleCombined.forEach((p) => {
      const comp = String(getEffectiveProposedCompany(p) || '').trim();
      if (comp && comp !== '-' && comp !== '未設定') {
        set.add(comp);
      }
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [allPeopleCombined]);

  const filteredCandidates = useMemo(() => {
    let baseList: Candidate[] = [];
    if (activeTab === 'personal') {
      baseList = personalCandidates;
    } else if (activeTab === 'company') {
      baseList = companyCandidates;
    } else if (activeTab === 'line') {
      baseList = lineCandidates;
    } else if (activeTab === 'other') {
      baseList = otherCandidates;
    } else if (activeTab === 'school') {
      baseList = schoolCandidates;
    } else if (activeTab === 'mendan') {
      baseList = mendanPeopleSorted;
    } else if (activeTab === 'proposed') {
      baseList = allProposedPeople;
    } else {
      baseList = allPeopleCombined;
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

      if (industryFilter !== 'all') {
        const pInd = String(p.desiredIndustry || '').toLowerCase();
        if (!pInd.includes(industryFilter.toLowerCase())) {
          return false;
        }
      }

      if (nationalityFilter !== 'all') {
        const pNat = String(p.nationality || '').trim();
        if (pNat !== nationalityFilter) {
          return false;
        }
      }

      if (proposedCompanyFilter !== 'all') {
        const pComp = String(getEffectiveProposedCompany(p) || '').trim();
        if (pComp !== proposedCompanyFilter) {
          return false;
        }
      }

      if (activeTab === 'proposed' && proposedMonthFilter !== 'all') {
        const mInfo = extractNormalizedMendanInfo(p);
        const ym = String(mInfo.dateLabel || p.lastContactDate || '').slice(0, 7);
        if (ym !== proposedMonthFilter) return false;
      }

      if (activeTab === 'proposed' && proposedStageFilter !== 'all') {
        const grp = classifyProposedSubTab(p);
        if (grp !== proposedStageFilter) return false;
        if (proposedStageFilter === 'jitai' && jitaiSubFilter !== 'all') {
          const st = String(p.progressStatus || '').trim();
          if (jitaiSubFilter === '辞退・終了') {
            if (['内定辞退', '選考中辞退', '不合格'].includes(st)) return false;
          } else if (st !== jitaiSubFilter) {
            return false;
          }
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
        if (!getUrgentFollowUpReason(p, uiLang)) return false;
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
    lineCandidates,
    otherCandidates,
    schoolCandidates,
    allPeopleCombined,
    schoolStudents,
    mendanPeopleSorted,
    allProposedPeople,
    searchQuery,
    statusFilter,
    industryFilter,
    nationalityFilter,
    proposedCompanyFilter,
    proposedMonthFilter,
    proposedStageFilter,
    jitaiSubFilter,
    careFilter,
    declineReasonQuery,
    uiLang,
  ]);

  const kyujinStats = useMemo(() => {
    let totalSlots = 0;
    let openSlots = 0;
    let closedSlots = 0;
    let bonusSlots = 0;
    let baitoSlots = 0;
    let missingSalarySlots = 0;
    kyujins.forEach((c) => {
      (c.areaBreakdown || []).forEach((sl) => {
        totalSlots += 1;
        if (sl.recruitmentStatus === '募集終了') closedSlots += 1;
        else openSlots += 1;
        if (sl.bonus === '賞与あり') bonusSlots += 1;
        const isSlotBaito = /baito|アルバイト|パート/i.test(
          sl.employmentType || c.employmentType || ''
        );
        if (isSlotBaito) {
          baitoSlots += 1;
        }
        if (isKyujinSalaryMissing(sl, isSlotBaito)) {
          missingSalarySlots += 1;
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
      missingSalarySlots,
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
          const isSlotBaito = /baito|アルバイト|パート/i.test(slotEmp);

          if (kyujinMissingSalaryOnly && !isKyujinSalaryMissing(sl, isSlotBaito)) return false;

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

          if (kyujinShiftTypeFilter !== 'all') {
            const effectiveShiftType =
              sl.shiftType || detectKyujinShiftType(sl.shiftTime, sl.jobDescription);
            if (kyujinShiftTypeFilter === '日勤のみ' && !/日勤/.test(effectiveShiftType)) {
              return false;
            }
            if (kyujinShiftTypeFilter === 'シフト制' && !/シフト/.test(effectiveShiftType)) {
              return false;
            }
            if (kyujinShiftTypeFilter === '夜勤あり' && !/夜勤/.test(effectiveShiftType)) {
              return false;
            }
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
              sl.shiftType,
              sl.holidays,
              sl.overtimeHours,
              sl.salary,
              sl.hourlyWage,
              sl.estimatedNetSalary,
              sl.annualSalary,
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
    kyujinMissingSalaryOnly,
    kyujinSectorFilter,
    kyujinIndustryFilter,
    kyujinEmploymentTypeFilter,
    kyujinFacilityFilter,
    kyujinShiftTypeFilter,
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
    showToast(t.toasts.contactedUpdated(getCandidateDisplayName(person, uiLang), todayIso));
  };

  const handleQuickStatusChange = (personId: string, isSchool: boolean, newStatus: string) => {
    const todayIso = getJstTodayYmd();
    const updater = (list: Candidate[]) =>
      list.map((p) =>
        p.id === personId
          ? {
              ...p,
              progressStatus: newStatus,
              proposalDate:
                newStatus === '企業提案中' && !p.proposalDate ? todayIso : p.proposalDate,
              lastContactDate: todayIso,
            }
          : p
      );
    if (isSchool) setSchoolStudents(updater);
    else setCandidates(updater);
    showToast(t.toasts.statusUpdated(newStatus));
  };

  const handleStatusTransitionRequest = (
    person: Candidate,
    isSchool: boolean,
    newStatus: string
  ) => {
    const candidateName = getCandidateDisplayName(person, uiLang);
    if (PROPOSED_NAITEI_STATUSES.includes(newStatus)) {
      setNaiteiPromptModal({
        personId: person.id,
        isSchool,
        candidateName,
        targetStatus: newStatus,
        naiteiCompany: person.naiteiCompany || getEffectiveProposedCompany(person) || '',
        expectedJoinDate: person.expectedJoinDate || '',
        visaSubmissionDeadline: person.visaSubmissionDeadline || '',
      });
      return;
    }
    if (PROPOSED_JITAI_STATUSES.includes(newStatus)) {
      setJitaiPromptModal({
        personId: person.id,
        isSchool,
        candidateName,
        targetStatus: newStatus,
        reason: person.declineReason || '',
      });
      return;
    }
    handleQuickStatusChange(person.id, isSchool, newStatus);
  };

  const handleConfirmNaiteiPrompt = () => {
    if (!naiteiPromptModal) return;
    const todayIso = new Date().toISOString().slice(0, 10);
    const {
      personId,
      isSchool,
      targetStatus,
      naiteiCompany,
      expectedJoinDate,
      visaSubmissionDeadline,
    } = naiteiPromptModal;
    const cleanComp = naiteiCompany.trim();
    const noteLine = `[${todayIso} ${targetStatus}] ${cleanComp || '-'}${
      expectedJoinDate ? ` · 入社予定: ${expectedJoinDate}` : ''
    }${visaSubmissionDeadline ? ` · ビザ書類期限: ${visaSubmissionDeadline}` : ''}`;

    const updater = (list: Candidate[]) =>
      list.map((p) => {
        if (p.id !== personId) return p;
        return {
          ...p,
          progressStatus: targetStatus,
          naiteiCompany: cleanComp || p.naiteiCompany || p.proposedCompany || '',
          proposedCompany: p.proposedCompany || cleanComp,
          expectedJoinDate: expectedJoinDate.trim(),
          visaSubmissionDeadline: visaSubmissionDeadline.trim(),
          lastContactDate: todayIso,
          memo: p.memo ? `${noteLine}\n${p.memo}` : noteLine,
        };
      });
    if (isSchool) setSchoolStudents(updater);
    else setCandidates(updater);
    setNaiteiPromptModal(null);
    showToast(t.toasts.statusUpdated(targetStatus));
  };

  const handleConfirmJitaiPrompt = () => {
    if (!jitaiPromptModal) return;
    const todayIso = new Date().toISOString().slice(0, 10);
    const { personId, isSchool, targetStatus, reason } = jitaiPromptModal;
    const cleanReason = reason.trim();
    const noteLine = `[${todayIso} ${targetStatus}] ${
      cleanReason ? `辞退/不合格理由: ${cleanReason}` : targetStatus
    }`;

    const updater = (list: Candidate[]) =>
      list.map((p) => {
        if (p.id !== personId) return p;
        return {
          ...p,
          progressStatus: targetStatus,
          declineReason: cleanReason || p.declineReason || targetStatus,
          lastContactDate: todayIso,
          memo: p.memo ? `${noteLine}\n${p.memo}` : noteLine,
        };
      });
    if (isSchool) setSchoolStudents(updater);
    else setCandidates(updater);
    setJitaiPromptModal(null);
    showToast(t.toasts.statusUpdated(targetStatus));
  };

  const handleSaveKyujinConditions = () => {
    if (!kyujinConditionModal) return;
    const {
      companyId,
      slotId,
      excludedNationalitiesText,
      excludedGenders,
      requiresInJapan,
      requiresTokutei,
      extraRequiredText,
      extraNgText,
    } = kyujinConditionModal;
    const nats = excludedNationalitiesText
      .split(/[,、・\s]+/)
      .map((s) => s.trim())
      .filter(Boolean);
    const reqs = extraRequiredText
      .split(/[,、\n]+/)
      .map((s) => s.trim())
      .filter(Boolean);
    const ngs = extraNgText
      .split(/[,、\n]+/)
      .map((s) => s.trim())
      .filter(Boolean);

    setKyujins((prev) =>
      prev.map((comp) => {
        if (comp.id !== companyId) return comp;
        const updatedSlots = (comp.areaBreakdown || []).map((sl) => {
          if (sl.slotId !== slotId) return sl;
          const merged = extractNormalizedKyujinConditions({
            text: `${sl.jobDescription || ''} ${sl.experienceDetail || ''}`,
            minJlpt: sl.minJlpt,
            employmentType: sl.employmentType || comp.employmentType,
            existingExcludedNationalities: nats,
            existingExcludedGenders: excludedGenders,
            existingRequiresInJapan: requiresInJapan,
            existingRequiresTokutei: requiresTokutei,
            existingRequiredConditions: reqs,
            existingNgConditions: ngs,
          });
          return {
            ...sl,
            excludedNationalities: merged.excludedNationalities,
            excludedGenders: merged.excludedGenders,
            requiresInJapan: merged.requiresInJapan,
            requiresTokutei: merged.requiresTokutei,
            requiredConditions: merged.requiredConditions,
            ngConditions: merged.ngConditions,
          };
        });
        return normalizeKyujinItem({
          ...comp,
          excludedNationalities: undefined,
          excludedGenders: undefined,
          requiresInJapan: undefined,
          requiresTokutei: undefined,
          requiredConditions: undefined,
          ngConditions: undefined,
          areaBreakdown: updatedSlots,
        });
      })
    );
    setKyujinConditionModal(null);
    showToast(
      uiLang === 'jp'
        ? '✅ 求人の必須条件・NG条件を更新しました'
        : '✅ Đã cập nhật điều kiện bắt buộc & điều kiện NG cho đơn tuyển dụng'
    );
  };

  const handleProposeCompanyToCandidate = (
    person: Candidate,
    isSchool: boolean,
    companyName: string
  ) => {
    const todayIso = getJstTodayYmd();
    const updater = (list: Candidate[]) =>
      list.map((p) =>
        p.id === person.id
          ? {
              ...p,
              proposedCompany: companyName,
              progressStatus: '企業提案中',
              proposalDate: p.proposalDate || todayIso,
              lastContactDate: todayIso,
              memo: `${p.memo ? p.memo + '\n' : ''}[${todayIso}] ${companyName} へ提案`,
            }
          : p
      );
    if (isSchool) setSchoolStudents(updater);
    else setCandidates(updater);
    showToast(t.toasts.proposedSet(getCandidateDisplayName(person, uiLang), companyName));
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

    // Guarantee 100% verbatim preservation in negotiationLog and concise summary in memo
    let incomingLog = String(extracted.negotiationLog || '').trim();
    if (rawTrimmed) {
      if (!incomingLog) {
        incomingLog = rawArchive;
      } else if (!incomingLog.includes(rawTrimmed)) {
        incomingLog = `${incomingLog}\n\n${rawArchive}`;
      }
    }

    const incomingMemo = String(extracted.memoLine || rawArchive || incomingLog || '').trim();

    const shouldReplaceName =
      Boolean(extracted.name) &&
      (!base.name || base.name === 'NEW CANDIDATE' || base.name === base.fbName);
    const extractedJlpt = String(extracted.jlpt || '').trim();
    const resolvedJlpt =
      extractedJlpt && extractedJlpt !== '要確認'
        ? extractedJlpt
        : base.jlpt || extractedJlpt || '要確認';
    const normalizedExtractedNat = normalizeCandidateNationality(extracted.nationality);

    return {
      ...base,
      channel: extracted.channel || base.channel,
      name: shouldReplaceName ? extracted.name! : base.name,
      fbName: extracted.fbName || base.fbName,
      fbProfileUrl: extracted.fbProfileUrl || extracted.fbUrl || base.fbProfileUrl,
      nationality: normalizedExtractedNat || base.nationality,
      gender: extracted.gender || base.gender,
      birthDate: extracted.birthDate || base.birthDate,
      age: extracted.age || base.age,
      email: extracted.email || base.email,
      address: extracted.address || base.address,
      jlpt: resolvedJlpt,
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

  const runCandidateExtraction = async (options?: {
    imagePayload?: {
      mimeType: string;
      base64Data: string;
      dataUrl: string;
      fileName: string;
    } | null;
    rawTextOverride?: string;
  }) => {
    if (!editingPerson) return;
    const activeImage =
      options?.imagePayload !== undefined ? options.imagePayload : currentImageBase64;
    const activeText = (
      options?.rawTextOverride !== undefined ? options.rawTextOverride : editChatSnippet
    ).trim();

    if (!activeImage && !activeText) {
      const warnMsg = '⚠️ Vui lòng dán đoạn chat hoặc tải/dán ảnh chụp màn hình trước';
      setCandidateExtractBanner({ type: 'warning', text: warnMsg });
      showToast(warnMsg);
      return;
    }

    setIsExtractingCandidateAi(true);
    setCandidateExtractBanner({
      type: 'loading',
      text: activeImage
        ? '⏳ AI Vision đang đọc ảnh và bóc tách dữ liệu...'
        : '⏳ AI đang bóc tách đoạn hội thoại...',
    });

    let usedAi = false;
    let usedOcrFallback = false;

    try {
      let aiExtracted: (Partial<Candidate> & { memoLine?: string; fbUrl?: string }) | null = null;
      const parts: Array<{ text?: string; inlineData?: { mimeType: string; data: string } }> = [];
      if (activeImage) {
        parts.push({
          inlineData: {
            mimeType: activeImage.mimeType,
            data: activeImage.base64Data,
          },
        });
      }
      if (activeText) {
        parts.push({ text: activeText });
      }

      try {
        const token = await getFirebaseIdToken();
        const aiData = await extractCandidateViaServer(
          {
            text: activeText || undefined,
            parts: parts.length > 0 ? parts : undefined,
          },
          token
        );
        if (aiData && Object.keys(aiData).length > 0) {
          aiExtracted = aiData;
          usedAi = true;
        }
      } catch {
        // Vision AI unavailable or not configured -> fall through to Tesseract.js OCR / parser
      }

      let combinedRawSource = activeText;
      if (!usedAi && activeImage) {
        usedOcrFallback = true;
        try {
          const ocrText = await recognizeCandidateImageWithTesseract(
            activeImage.dataUrl,
            (progressText) => {
              setCandidateExtractBanner({
                type: 'loading',
                text: `⏳ ${progressText}`,
              });
            }
          );
          if (ocrText) {
            combinedRawSource = [activeText, ocrText].filter(Boolean).join('\n\n');
          }
        } catch (ocrErr) {
          const ocrMsg =
            ocrErr instanceof Error ? ocrErr.message : 'Không thể chạy OCR đọc ảnh';
          if (!activeText) {
            setCandidateExtractBanner({
              type: 'warning',
              text: `⚠️ Không đọc được ảnh (${ocrMsg}). Vui lòng thử ảnh rõ nét hơn hoặc dán nội dung chat.`,
            });
            showToast(`⚠️ Lỗi đọc ảnh: ${ocrMsg}`);
            return;
          }
        }
      }

      const fallbackSource =
        combinedRawSource || aiExtracted?.negotiationLog || aiExtracted?.memo || '';
      const fallback = parseChatIntoUpdates(fallbackSource);
      const verbatimRaw =
        combinedRawSource || aiExtracted?.negotiationLog || aiExtracted?.memo || '';

      const mergedExtract: Partial<Candidate> & { memoLine?: string; fbUrl?: string } = {
        ...fallback,
        ...(aiExtracted || {}),
        // Preserve Vision AI's Messenger vs ID-card distinction (do not invent fullName from chat if Vision AI left it empty)
        name: usedAi && activeImage ? aiExtracted?.name || '' : aiExtracted?.name || fallback.name || '',
        fbName: aiExtracted?.fbName || fallback.fbName || '',
        fbProfileUrl: aiExtracted?.fbProfileUrl || fallback.fbProfileUrl || '',
        visaStatus: aiExtracted?.visaStatus || fallback.visaStatus || '',
        visaExpiry: aiExtracted?.visaExpiry || fallback.visaExpiry || '',
        jlpt:
          usedAi && activeImage
            ? aiExtracted?.jlpt || '要確認'
            : aiExtracted?.jlpt || fallback.jlpt || '要確認',
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
        proposedCompany: aiExtracted?.proposedCompany || fallback.proposedCompany || '',
        mendanDate: aiExtracted?.mendanDate || fallback.mendanDate || '',
        interviewDate: aiExtracted?.interviewDate || fallback.interviewDate || '',
        progressStatus: aiExtracted?.progressStatus || fallback.progressStatus || '',
        negotiationLog: aiExtracted?.negotiationLog || fallback.negotiationLog || verbatimRaw,
        memoLine: fallback.memoLine,
      };

      setEditingPerson((prev) =>
        prev
          ? {
              ...prev,
              data: mergeExtractedCandidateData(prev.data, mergedExtract, verbatimRaw),
            }
          : prev
      );

      if (activeText) {
        setEditChatSnippet('');
      }

      const successText = activeImage
        ? usedOcrFallback
          ? '✅ Đã bóc tách thành công từ ảnh chụp! (OCR)'
          : '✅ Đã bóc tách thành công từ ảnh chụp!'
        : usedAi
          ? '✅ AI đã bóc tách & bảo toàn 100% nguyên văn vào hồ sơ!'
          : '✅ Đã bóc tách bằng bộ lọc dự phòng & lưu 100% nguyên văn!';

      setCandidateExtractBanner({ type: 'success', text: successText });
      showToast(successText);
    } finally {
      setIsExtractingCandidateAi(false);
    }
  };

  const handleApplyChatToEditingPerson = async () => {
    await runCandidateExtraction();
  };

  const handleExtractCandidateScreenshot = async (
    file: File,
    target: 'edit' | 'fbChat'
  ) => {
    if (!file.type.startsWith('image/')) {
      showToast('⚠️ Vui lòng chọn hoặc dán ảnh chụp màn hình hợp lệ');
      return;
    }

    if (target === 'edit') {
      try {
        setCandidateExtractBanner({
          type: 'loading',
          text: '⏳ Đang nén và chuẩn bị ảnh chụp...',
        });
        const compressed = await compressImageFileToBase64(file, 1600, 0.82);
        const imagePayload = {
          ...compressed,
          fileName: file.name || `screenshot_${Date.now()}.jpg`,
        };
        setCurrentImageBase64(imagePayload);
        await runCandidateExtraction({ imagePayload });
      } catch (err) {
        const msg = err instanceof Error ? err.message : 'Không thể đọc tệp ảnh';
        setCandidateExtractBanner({ type: 'warning', text: `⚠️ Lỗi đọc ảnh: ${msg}` });
        showToast(`⚠️ Lỗi đọc ảnh: ${msg}`);
      }
      return;
    }

    setIsExtractingFbChatAi(true);
    try {
      const compressed = await compressImageFileToBase64(file, 1600, 0.82);
      let extracted: Partial<Candidate> | null = null;
      let ocrFallbackText = '';

      try {
        const token = await getFirebaseIdToken();
        const aiRes = await extractCandidateViaServer(
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
        if (aiRes && Object.keys(aiRes).length > 0) {
          extracted = aiRes;
        }
      } catch {
        // Vision AI failed -> fallback to Tesseract.js OCR
      }

      if (!extracted) {
        ocrFallbackText = await recognizeCandidateImageWithTesseract(compressed.dataUrl);
        extracted = parseChatIntoUpdates(ocrFallbackText);
      }

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
        extracted.negotiationLog || extracted.memo || ocrFallbackText || '',
      ]
        .filter(Boolean)
        .join('\n');

      setFbChatRawText((prev) =>
        prev.trim() ? `${prev.trim()}\n\n${formattedSummary}` : formattedSummary
      );
      showToast('✅ Đã bóc tách thành công từ ảnh chụp!');
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Không thể đọc ảnh chụp lúc này';
      showToast(`⚠️ Lỗi đọc ảnh: ${msg}`);
    } finally {
      setIsExtractingFbChatAi(false);
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
        nationality: normalizeCandidateNationality(extracted.nationality) || '国籍不明',
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
        memo: String(extracted.memoLine || rawArchive || '').trim(),
      };

      if (fbChatChannel === 'school') {
        setSchoolStudents((prev) => [newPerson, ...prev]);
        setActiveTab('school');
      } else {
        setCandidates((prev) => [newPerson, ...prev]);
        setActiveTab(
          fbChatChannel === 'company'
            ? 'company'
            : fbChatChannel === 'line'
            ? 'line'
            : fbChatChannel === 'other'
            ? 'other'
            : 'personal'
        );
      }

      setFbChatModalOpen(false);
      setFbChatRawText('');
      const channelToastLabel =
        fbChatChannel === 'company'
          ? t.tabs.company
          : fbChatChannel === 'line'
          ? t.tabs.line
          : fbChatChannel === 'other'
          ? t.tabs.other
          : fbChatChannel === 'school'
          ? t.tabs.school
          : t.tabs.personal;
      showToast(
        uiLang === 'jp'
          ? `✅ ${channelToastLabel} に新規相談者を登録しました: ${newPerson.name}${usedAi ? '' : '（フォールバック解析）'}`
          : `✅ Đã tạo hồ sơ mới tại ${channelToastLabel}: ${newPerson.name}${usedAi ? '' : ' (bộ lọc dự phòng)'}`
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
      channelManuallyEdited: true,
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
        ? t.tabs.company
        : nextChannel === 'line'
        ? t.tabs.line
        : nextChannel === 'other'
        ? t.tabs.other
        : nextChannel === 'school'
        ? t.tabs.school
        : t.tabs.personal;
    const displayName = getCandidateDisplayName(person, uiLang);
    showToast(
      uiLang === 'jp'
        ? `🔄 「${displayName}」さんのチャネルを ${label} に変更しました（データ保持100%）`
        : `🔄 Đã chuyển "${displayName}" sang kênh ${label} (bảo toàn 100% dữ liệu)`
    );
  };

  const handleBulkAutoFixCandidateNationalities = () => {
    const fixedCand = autoFixCandidatesNationalities(candidates);
    const fixedStu = autoFixCandidatesNationalities(schoolStudents);
    const totalChanged = fixedCand.changedCount + fixedStu.changedCount;
    if (fixedCand.changedCount > 0) setCandidates(fixedCand.updated);
    if (fixedStu.changedCount > 0) setSchoolStudents(fixedStu.updated);
    showToast(t.autoFixNatToast(totalChanged));
  };

  const handleDataCleanupAndReclassify = () => {
    const res = cleanupAndReclassifyCandidates(candidates, schoolStudents);
    if (res.changedCount > 0) {
      setCandidates(res.candidates);
      setSchoolStudents(res.schoolStudents);
    }
    const dedupedKyujins = deduplicateAndNormalizeKyujins(kyujins);
    if (JSON.stringify(dedupedKyujins) !== JSON.stringify(kyujins)) {
      setKyujins(dedupedKyujins);
    }
    try {
      localStorage.setItem(CHANNEL_CLEANUP_VERSION_KEY, 'done');
    } catch {
      // ignore storage errors
    }
    const totalUnique = res.candidates.length + res.schoolStudents.length;
    showToast(
      t.dataCleanupToast(
        res.movedToOther,
        res.movedToCompany,
        res.movedToLine,
        res.movedToSchool,
        res.dedupedCount,
        totalUnique
      )
    );
  };

  const handleSaveEditingPerson = () => {
    if (!editingPerson) return;
    const { data, isSchool, isNew } = editingPerson;
    const trimmedName = String(data.name || '').trim();
    const trimmedFb = String(data.fbName || '').trim();
    if (!trimmedName && !trimmedFb) {
      showToast(
        uiLang === 'jp'
          ? '⚠️ 氏名またはFBアカウント名を入力してください'
          : '⚠️ Vui lòng nhập Họ tên khai sinh hoặc Tên Facebook'
      );
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
    const nextChannelValue = isSchoolTarget ? 'school' : data.channel || 'personal';
    const channelChangedByUser = Boolean(
      prevPerson && (prevPerson.channel || 'personal') !== nextChannelValue
    );

    const savedPerson: Candidate = {
      ...data,
      name: effectiveName,
      memo: fallbackNote,
      channel: nextChannelValue,
      channelManuallyEdited: channelChangedByUser ? true : data.channelManuallyEdited,
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
    showToast(isNew ? t.toasts.profileSavedNew : t.toasts.profileSavedEdit);

    const scheduleChanged =
      !prevPerson ||
      (prevPerson.mendanDate || '').trim() !==
        (savedPerson.mendanDate || '').trim() ||
      (prevPerson.interviewDate || '').trim() !==
        (savedPerson.interviewDate || '').trim() ||
      (prevPerson.proposedCompany || '').trim() !==
        (savedPerson.proposedCompany || '').trim();

    const nextDraft = buildCandidateCalendarEventDraft(savedPerson);
    const savedDisplayName = getCandidateDisplayName(savedPerson, uiLang);

    if (nextDraft && scheduleChanged) {
      const hasExistingEvent = Boolean(savedPerson.googleCalendarEventId);
      setConfirmDialog({
        title: hasExistingEvent
          ? uiLang === 'jp'
            ? 'Googleカレンダーの予定を更新しますか？'
            : 'Cập nhật lịch trên Google Calendar?'
          : uiLang === 'jp'
          ? 'Googleカレンダーに新規予定を作成しますか？'
          : 'Tạo sự kiện mới trên Google Calendar?',
        message:
          uiLang === 'jp'
            ? `${savedDisplayName} さんの予定をCRMに保存しました。\n• 件名: ${nextDraft.eventTitle}\n• 日時: ${nextDraft.displayDateTime}\n• リマインダー: 30分前・10分前\n\n今すぐGoogleカレンダーへ同期しますか？`
            : `Đã lưu lịch hẹn của ${savedDisplayName} trong CRM.\n• Tiêu đề: ${nextDraft.eventTitle}\n• Thời gian: ${nextDraft.displayDateTime}\n• Nhắc nhở: 30 phút & 10 phút trước.\n\nBạn có muốn đồng bộ lên Google Calendar ngay bây giờ không?`,
        confirmLabel: hasExistingEvent
          ? uiLang === 'jp'
            ? 'カレンダーを更新'
            : 'Cập nhật Calendar'
          : uiLang === 'jp'
          ? 'カレンダーに作成'
          : 'Tạo lịch trên Calendar',
        cancelLabel: uiLang === 'jp' ? 'CRMのみ保存' : 'Chỉ lưu CRM',
        tone: 'accent',
        onConfirm: () => {
          setConfirmDialog(null);
          void executeCalendarSyncForDraft(nextDraft);
        },
      });
    } else if (!nextDraft && prevPerson?.googleCalendarEventId) {
      const staleEventId = prevPerson.googleCalendarEventId;
      setConfirmDialog({
        title:
          uiLang === 'jp'
            ? 'Googleカレンダーの予定を削除しますか？'
            : 'Hủy sự kiện trên Google Calendar?',
        message:
          uiLang === 'jp'
            ? `${savedDisplayName} さんの面談日時が削除されました。連携していたGoogleカレンダーの予定も削除しますか？`
            : `Bạn đã xóa ngày giờ lịch hẹn của ${savedDisplayName}. Bạn có muốn xóa sự kiện tương ứng khỏi Google Calendar không?`,
        confirmLabel:
          uiLang === 'jp' ? 'Googleカレンダーから削除' : 'Xóa trên Google Calendar',
        cancelLabel: uiLang === 'jp' ? '残す' : 'Giữ lại',
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
    const targetDisplayName = getCandidateDisplayName(targetPerson, uiLang);

    setConfirmDialog({
      title: uiLang === 'jp' ? '削除の確認' : 'Xác nhận xóa hồ sơ',
      message: linkedEventId
        ? uiLang === 'jp'
          ? `この相談者データ（${targetDisplayName}）を削除しますか？関連付けられた Google Calendar の予定も同時に削除されます。`
          : `Bạn có chắc muốn xóa hồ sơ (${targetDisplayName})? Sự kiện liên kết trên Google Calendar cũng sẽ bị xóa.`
        : uiLang === 'jp'
        ? `この相談者データ（${targetDisplayName}）を削除してもよろしいですか？`
        : `Bạn có chắc chắn muốn xóa hồ sơ (${targetDisplayName}) không?`,
      confirmLabel: uiLang === 'jp' ? '削除を実行する' : 'Xác nhận xóa',
      cancelLabel: t.cancel,
      tone: 'danger',
      onConfirm: () => {
        if (isSchool) {
          setSchoolStudents((prev) => prev.filter((s) => s.id !== personId));
        } else {
          setCandidates((prev) => prev.filter((c) => c.id !== personId));
        }
        setConfirmDialog(null);
        showToast(t.toasts.deletedCandidate);
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

  const handleToggleCompanyRecruitmentStatus = (
    companyId: string,
    companyName: string,
    nextStatus: '募集中' | '募集終了'
  ) => {
    setKyujins((prev) =>
      prev.map((comp) => {
        if (comp.id !== companyId) return comp;
        const updatedSlots = (comp.areaBreakdown || []).map((sl) => ({
          ...sl,
          recruitmentStatus: nextStatus,
        }));
        return normalizeKyujinItem({
          ...comp,
          recruitmentStatus: nextStatus,
          areaBreakdown: updatedSlots,
        });
      })
    );
    showToast(
      uiLang === 'jp'
        ? nextStatus === '募集終了'
          ? `⏸ 「${companyName}」の全ポジションを一時停止（募集終了）にしました`
          : `▶️ 「${companyName}」の全ポジションを再開（募集中）にしました`
        : nextStatus === '募集終了'
        ? `⏸ Đã tạm dừng toàn bộ vị trí của "${companyName}"`
        : `▶️ Đã mở lại toàn bộ vị trí của "${companyName}"`
    );
  };

  const handleSaveInlineKyujinField = (
    companyId: string,
    slotId: string,
    field: 'shiftTime' | 'jobCode' | 'area',
    rawValue: string
  ) => {
    setEditingKyujinField(null);
    const trimmed = String(rawValue || '').trim();
    if (!trimmed) return;
    handleQuickUpdateSlotField(companyId, slotId, field, trimmed);
    if (field === 'shiftTime') {
      const detected = detectKyujinShiftType(trimmed, '');
      handleQuickUpdateSlotField(companyId, slotId, 'shiftType', detected);
    }
    showToast(
      uiLang === 'jp' ? '✅ 求人情報を更新しました' : '✅ Đã cập nhật thông tin tin tuyển dụng'
    );
  };

  const handleStartInlineKyujinSalaryEdit = (
    companyId: string,
    slotId: string,
    initialValue: string
  ) => {
    salaryEditCommitLockRef.current = null;
    setEditingKyujinSalary({
      companyId,
      slotId,
      value: initialValue,
    });
  };

  const handleSaveInlineKyujinSalary = (
    companyId: string,
    slotId: string,
    rawInput: string
  ) => {
    const lockKey = `${companyId}::${slotId}`;
    if (salaryEditCommitLockRef.current === lockKey) return;
    salaryEditCommitLockRef.current = lockKey;
    setEditingKyujinSalary(null);

    const trimmed = String(rawInput || '').trim();
    if (!trimmed) {
      return;
    }

    const parsed = normalizeManualKyujinSalaryInput(trimmed);
    const targetComp = kyujins.find((c) => c.id === companyId);
    const targetSlot = targetComp?.areaBreakdown?.find((s) => s.slotId === slotId);
    const existingMeta = targetSlot
      ? getKyujinSlotSalaryDisplayMeta(targetSlot, targetComp?.employmentType || '')
      : null;

    if (
      existingMeta &&
      !existingMeta.isMissing &&
      parsed.salary === targetSlot?.salary &&
      parsed.hourlyWage === (targetSlot?.hourlyWage || '')
    ) {
      return;
    }

    const updatedList = kyujins.map((comp) => {
      if (comp.id !== companyId) return comp;
      const updatedSlots = (comp.areaBreakdown || []).map((sl) => {
        if (sl.slotId !== slotId) return sl;
        const protectedSet = new Set<string>([
          ...(Array.isArray(sl.manualProtectedFields) ? sl.manualProtectedFields : []),
          'salary',
          'hourlyWage',
        ]);
        return {
          ...sl,
          salary: parsed.salary,
          hourlyWage: parsed.hourlyWage,
          salaryManuallyEdited: true,
          manualProtectedFields: Array.from(protectedSet),
        };
      });
      return normalizeKyujinItem({
        ...comp,
        areaBreakdown: updatedSlots,
      });
    });

    setKyujins(updatedList);
    try {
      localStorage.setItem(STORAGE_KEYS.KYUJINS, JSON.stringify(updatedList));
    } catch {
      // ignore localStorage quota errors
    }

    const displayLabel =
      parsed.salaryType === 'hourly'
        ? `時給: ${parsed.hourlyWage || parsed.salary}`
        : `Lương: ${parsed.salary}`;

    const trimmedGasUrl = gasWebAppUrl.trim();
    if (trimmedGasUrl && isValidSafeGasUrl(trimmedGasUrl)) {
      void fetch(trimmedGasUrl, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'updateKyujinSalary',
          respectManualProtection: true,
          salaryManuallyEdited: true,
          manualProtectedFields: ['salary', 'hourlyWage'],
          companyId,
          companyName: targetComp?.companyName || '',
          slotId,
          jobCode: targetSlot?.jobCode || '',
          salary: parsed.salary,
          hourlyWage: parsed.hourlyWage,
        }),
      }).catch(() => {
        // ignore network error in background GAS sync
      });
      showToast(`✅ Đã lưu ${displayLabel} & đồng bộ GAS (đã bật bảo vệ ghi đè thủ công)`);
    } else {
      showToast(`✅ Đã lưu ${displayLabel} vào LocalStorage (đã bật bảo vệ ghi đè thủ công)`);
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
          const normalizedIncomingName =
            normalizeKyujinCompanyName(entry.companyName) || '新規取込企業';
          const incomingNameKey = normalizedIncomingName
            .replace(/[\s　]+/g, '')
            .toLowerCase();
          const existingIdx = updated.findIndex(
            (c) =>
              normalizeKyujinCompanyName(c.companyName)
                .replace(/[\s　]+/g, '')
                .toLowerCase() === incomingNameKey
          );
          if (existingIdx >= 0) {
            const target = updated[existingIdx];
            const existingSlots = [...(target.areaBreakdown || [])];
            const incomingCode = stripKyujinPdfNoise(String(entry.slot.jobCode || '')).trim();
            const matchingSlotIdx =
              incomingCode && !/^新規枠$/i.test(incomingCode)
                ? existingSlots.findIndex(
                    (sl) =>
                      sl.slotId === entry.slot.slotId ||
                      (stripKyujinPdfNoise(String(sl.jobCode || '')).trim() === incomingCode &&
                        (sl.salaryManuallyEdited ||
                          sl.manualProtectedFields?.includes('salary')))
                  )
                : -1;

            if (matchingSlotIdx >= 0) {
              const protectedSlot = existingSlots[matchingSlotIdx];
              const isSalaryProtected =
                protectedSlot.salaryManuallyEdited ||
                protectedSlot.manualProtectedFields?.includes('salary');
              existingSlots[matchingSlotIdx] = {
                ...entry.slot,
                slotId: protectedSlot.slotId,
                ...(isSalaryProtected
                  ? {
                      salary: protectedSlot.salary,
                      hourlyWage: protectedSlot.hourlyWage,
                      salaryManuallyEdited: protectedSlot.salaryManuallyEdited,
                      manualProtectedFields: protectedSlot.manualProtectedFields,
                    }
                  : {}),
              };
            } else {
              existingSlots.push(entry.slot);
            }

            updated[existingIdx] = normalizeKyujinItem({
              ...target,
              companyName: normalizedIncomingName,
              areaBreakdown: existingSlots,
            });
          } else {
            updated.unshift(
              normalizeKyujinItem({
                id: `kj_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
                companyName: normalizedIncomingName,
                sector: entry.sector || entry.slot.sector,
                industry: entry.industry,
                employmentType: entry.employmentType || entry.slot.employmentType,
                recruitmentStatus: '募集中',
                areaBreakdown: [entry.slot],
              })
            );
          }
        });
        return deduplicateAndNormalizeKyujins(updated);
      });
    },
    []
  );

  const mapServerKyujinItemsToEntries = (
    items: Array<Record<string, unknown>>,
    sourceName: string
  ) => {
    const fallbackComp = sourceName
      .replace(/^求人票[_・\-\s]*/i, '')
      .replace(/様.*$/, '')
      .replace(/\.[^.]+$/, '')
      .trim();

    return items.map((parsed, idx) => {
      const str = (key: string) => String(parsed[key] ?? '').trim();
      const strArr = (key: string) =>
        Array.isArray(parsed[key])
          ? (parsed[key] as unknown[]).map((v) => String(v ?? '').trim()).filter(Boolean)
          : [];
      const cleanArea = formatPublicArea(str('area')) || str('area');
      const cleanDesc = sanitizeJobDescriptionForCandidate(str('jobDescription'));
      const shiftLabel = str('shiftTime');
      const industry = detectIndustryFromText(
        `${str('industry')} ${cleanDesc}`,
        `${sourceName} ${str('industry')}`
      );
      const sector =
        str('sector') ||
        detectSectorFromIndustryOrText(industry, `${str('industry')} ${cleanDesc}`);

      const rawSalaryHint = [
        str('salary') && !/要確認/.test(str('salary')) ? str('salary') : '',
        str('hourlyWage') && !/要確認/.test(str('hourlyWage'))
          ? `時給 ${str('hourlyWage')}`
          : '',
        cleanDesc,
      ]
        .filter(Boolean)
        .join('\n');
      const parsedSalary = extractKyujinSalaryInfo(rawSalaryHint, {
        assumeBareNumberIsSalary: Boolean(str('salary') && !/要確認/.test(str('salary'))),
      });
      const resolvedSalary =
        parsedSalary.salaryType !== 'unknown'
          ? parsedSalary.salary
          : str('salary') || (str('hourlyWage') ? `時給 ${str('hourlyWage')}` : '給与要確認');
      const resolvedHourlyWage =
        parsedSalary.salaryType === 'hourly' ? parsedSalary.hourlyWage : str('hourlyWage');

      const rawEmpType = str('employmentType');
      const fullEmpContext = `${rawEmpType} ${resolvedSalary} ${str('annualSalary')} ${str('targetVisa')} ${cleanDesc} ${sourceName}`;
      const hasMonthlyOrFullTimeSignals =
        parsedSalary.salaryType === 'monthly' ||
        /(?:月給|基本給|想定月収|総支給|想定年収|モデル年収|正社員|正職員|社員登用)/.test(
          fullEmpContext
        );

      let employmentType =
        rawEmpType ||
        detectEmploymentTypeFromText(
          `${resolvedSalary} ${resolvedHourlyWage} ${cleanDesc}`,
          sourceName
        );
      if (/baito|アルバイト|パート/i.test(employmentType) && hasMonthlyOrFullTimeSignals) {
        employmentType = /特定技能|tokutei/i.test(fullEmpContext)
          ? 'Tokutei Ginou (特定技能)'
          : 'Nhân viên chính thức / Chuyển đổi visa (正社員・登用)';
      } else if (
        !rawEmpType &&
        !hasMonthlyOrFullTimeSignals &&
        parsedSalary.salaryType === 'hourly' &&
        !/tokutei|特定技能|正社員/i.test(fullEmpContext)
      ) {
        employmentType = 'Baito / Làm thêm (アルバイト・パート)';
      }

      const rawJlpt = str('minJlpt');
      const normalizedJlpt = /^(N[1-5]|日常会話|不問)$/.test(rawJlpt)
        ? rawJlpt
        : '要確認';
      const derivedCode =
        str('jobCode') || cleanArea || `${sourceName.replace(/\.[^.]+$/, '')}-${idx + 1}`;
      const bonusNorm = normalizeBonusInfo(str('bonus'), str('bonusDetail'), cleanDesc);
      const overseasRecruitOk = str('overseasRecruitOk');
      const conds = extractNormalizedKyujinConditions({
        text: `${cleanDesc} ${str('experienceDetail')} ${str('experienceRequirement')} ${overseasRecruitOk ? `海外招へい ${overseasRecruitOk}` : ''}`,
        minJlpt: normalizedJlpt,
        employmentType,
        existingExcludedNationalities: strArr('excludedNationalities'),
        existingExcludedGenders: strArr('excludedGenders'),
        existingRequiresInJapan:
          Boolean(parsed.requiresInJapan) || overseasRecruitOk === '不可',
        existingRequiresTokutei: Boolean(parsed.requiresTokutei),
        existingRequiredConditions: strArr('requiredConditions'),
        existingNgConditions: strArr('ngConditions'),
      });

      const cleanCompName =
        normalizeKyujinCompanyName(str('companyName') || fallbackComp) || '新規取込企業';
      const cleanedJobCode = stripKyujinPdfNoise(derivedCode) || derivedCode;
      const cleanTargetVisa = stripKyujinPdfNoise(str('targetVisa'));
      const cleanTrialPeriodNote = stripKyujinPdfNoise(str('trialPeriodNote'));

      return {
        companyName: cleanCompName,
        sector,
        industry,
        employmentType,
        recruitmentStatus: '募集中',
        slot: {
          slotId: `slot_vis_${Date.now()}_${idx}_${Math.random().toString(36).slice(2, 7)}`,
          sector,
          industry,
          employmentType,
          targetVisa: cleanTargetVisa || undefined,
          trialPeriodNote: cleanTrialPeriodNote || undefined,
          overseasRecruitOk: overseasRecruitOk || undefined,
          jobCode: cleanedJobCode,
          area: stripKyujinPdfNoise(cleanArea) || '勤務地要確認',
          shiftTime: stripKyujinPdfNoise(shiftLabel),
          shiftType: str('shiftType') || detectKyujinShiftType(shiftLabel, cleanDesc),
          holidays: stripKyujinPdfNoise(str('holidays')),
          overtimeHours: stripKyujinPdfNoise(str('overtimeHours')),
          salary: resolvedSalary,
          nightShiftAllowance: str('nightShiftAllowance'),
          estimatedNetSalary: str('estimatedNetSalary'),
          annualSalary: str('annualSalary'),
          hourlyWage: resolvedHourlyWage,
          commuteAllowance: stripKyujinPdfNoise(str('commuteAllowance')),
          shiftsPerWeek: str('shiftsPerWeek'),
          bonus: bonusNorm.bonus,
          bonusDetail: stripKyujinPdfNoise(bonusNorm.bonusDetail),
          minJlpt: normalizedJlpt,
          housingSupport: normalizeHousingDisplay(str('housingSupport')),
          recruitmentStatus: '募集中',
          recruitCount: str('recruitCount'),
          interviewProcess: stripKyujinPdfNoise(str('interviewProcess')),
          specialBenefits: stripKyujinPdfNoise(str('specialBenefits')),
          careQualification: str('careQualification'),
          experienceRequirement: str('experienceRequirement'),
          experienceDetail: stripKyujinPdfNoise(str('experienceDetail')),
          jobDescription: cleanDesc,
          excludedNationalities: conds.excludedNationalities,
          excludedGenders: conds.excludedGenders,
          requiresInJapan: conds.requiresInJapan,
          requiresTokutei: conds.requiresTokutei,
          requiredConditions: conds.requiredConditions,
          ngConditions: conds.ngConditions,
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
    const fileErrors: string[] = [];

    try {
      for (let fIdx = 0; fIdx < files.length; fIdx++) {
        const file = files[fIdx];
        setKyujinImportProgress({
          current: fIdx + 1,
          total: files.length,
          label: `${file.name} — Đang chuẩn bị...`,
        });
        // Yield briefly to keep UI responsive on large multi-file batches
        await new Promise((r) => setTimeout(r, 10));

        const lower = file.name.toLowerCase();
        try {
          if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) {
            setKyujinImportProgress({
              current: fIdx + 1,
              total: files.length,
              label: `${file.name} — Đang đọc dữ liệu Excel...`,
            });
            const excelEntries = await parseExcelFileToKyujinEntries(file);
            extractedEntries.push(...excelEntries);
          } else {
            const token = await getFirebaseIdToken();
            try {
              const items = await extractKyujinViaServer(file, token, (statusText) => {
                setKyujinImportProgress({
                  current: fIdx + 1,
                  total: files.length,
                  label: `${file.name} — ${statusText}`,
                });
              });
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
        } catch (singleErr) {
          const msg =
            singleErr instanceof Error
              ? singleErr.message
              : `Lỗi đọc file "${file.name}"`;
          fileErrors.push(msg);
        }
      }

      if (extractedEntries.length > 0) {
        mergeExtractedKyujinEntries(extractedEntries);
        if (fileErrors.length > 0) {
          showToast(
            `⚠️ Đã nhập ${extractedEntries.length} đơn, ${fileErrors.length} file lỗi: ${fileErrors[0]}`
          );
        } else {
          showToast(`✅ Đã đọc và nhập ${extractedEntries.length} đơn / vị trí tuyển dụng mới!`);
        }
      } else if (fileErrors.length > 0) {
        showToast(`⚠️ ${fileErrors[0]}`);
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
      let items: Array<Record<string, unknown>> = [];
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
    const matchEval = evaluateCandidateKyujinMatches(person, kyujins, uiLang);
    const matchingKyujins = matchEval.matches;
    const isJobMatchExpanded = Boolean(expandedJobMatchIds[person.id]);
    const visibleMatches = isJobMatchExpanded ? matchingKyujins : matchingKyujins.slice(0, 2);
    const proposedComp = getEffectiveProposedCompany(person);
    const declineReason = extractDeclineReason(person);
    const displayName = getCandidateDisplayName(person, uiLang);
    const now = new Date();
    const todayIso = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(
      now.getDate()
    ).padStart(2, '0')}`;
    const isContactedToday = String(person.lastContactDate || '').trim() === todayIso;

    return (
      <div
        key={person.id}
        className="rounded-lg border border-[#E2E8F0] bg-white p-4 shadow-[0_1px_3px_rgba(0,0,0,0.05)] transition-colors hover:border-[#CBD5E1]"
      >
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div>
            <div className="flex flex-wrap items-center gap-2.5">
              <span className="text-base font-semibold tracking-tight text-[#0F172A]">
                {displayName}
              </span>
              {person.fbName && person.fbName !== displayName && (
                <span className="font-mono text-[11px] font-medium text-sky-700">
                  FB: {person.fbName}
                </span>
              )}
              <select
                value={
                  isSchool
                    ? 'school'
                    : person.channel === 'company'
                    ? 'company'
                    : person.channel === 'line'
                    ? 'line'
                    : person.channel === 'other'
                    ? 'other'
                    : 'personal'
                }
                onChange={(e) =>
                  handleQuickChannelChange(person, e.target.value as ChannelType)
                }
                title={t.quickChannelTitle}
                className="rounded border border-[#E2E8F0] bg-[#F8FAFC] px-2 py-0.5 text-[10px] font-semibold text-[#475569] hover:border-[#16A34A] hover:text-[#0F172A] focus:border-[#16A34A] focus:outline-none cursor-pointer"
              >
                <option value="personal">{t.channelPersonal}</option>
                <option value="company">{t.channelCompany}</option>
                <option value="line">{t.channelLine}</option>
                <option value="other">{t.channelOther}</option>
                <option value="school">{t.channelSchool}</option>
              </select>
              {isContactedToday ? (
                <span className="rounded bg-emerald-50 border border-emerald-200 px-1.5 py-0.5 font-mono text-[10px] font-semibold text-emerald-700 tabular-nums">
                  {t.contactedPrefix} {todayIso}
                </span>
              ) : (
                person.lastContactDate && (
                  <span className="font-mono text-[10px] text-[#64748B] tabular-nums">
                    {t.lastContactPrefix} {person.lastContactDate}
                  </span>
                )
              )}
              {proposedComp && (
                <span
                  className={`rounded border px-1.5 py-0.5 font-mono text-[10px] font-semibold ${
                    isActualProposedRecord(person)
                      ? 'border-sky-200 bg-sky-50 text-sky-700'
                      : 'border-amber-200 bg-amber-50 text-amber-800'
                  }`}
                >
                  {isActualProposedRecord(person)
                    ? `${t.proposedPrefix} ${proposedComp}`
                    : `${t.proposedCandidateBadge} ${proposedComp}`}
                </span>
              )}
              {isStale7DayProposal(person) && (
                <span className="rounded border border-rose-300 bg-rose-50 px-1.5 py-0.5 font-mono text-[10px] font-bold text-rose-700">
                  {t.stale7DaysBadge}
                </span>
              )}
            </div>
            <div className="mt-1 flex flex-wrap items-center gap-1.5 font-mono text-[10px] font-semibold uppercase tracking-wider">
              <span
                className={`rounded px-1.5 py-0.5 ${
                  person.progressStatus === '内定辞退'
                    ? 'border border-orange-300 bg-orange-50 text-orange-800'
                    : PROPOSED_NAITEI_STATUSES.includes(person.progressStatus || '')
                    ? 'border border-emerald-200 bg-emerald-50 text-emerald-700'
                    : PROPOSED_JITAI_STATUSES.includes(person.progressStatus || '')
                    ? 'border border-rose-200 bg-rose-50 text-rose-700'
                    : 'text-emerald-700'
                }`}
              >
                {t.statusPrefix} {person.progressStatus || '新規相談'}
              </span>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-1.5">
            <select
              value={person.progressStatus || '新規相談'}
              onChange={(e) => handleStatusTransitionRequest(person, isSchool, e.target.value)}
              title={t.quickStatusTitle}
              className={`rounded-md border px-2.5 py-1.5 text-[11px] font-semibold focus:border-[#16A34A] focus:outline-none whitespace-nowrap shrink-0 cursor-pointer ${
                person.progressStatus === '内定辞退'
                  ? 'border-orange-300 bg-orange-50 text-orange-800'
                  : PROPOSED_NAITEI_STATUSES.includes(person.progressStatus || '')
                  ? 'border-emerald-200 bg-emerald-50 text-emerald-700'
                  : PROPOSED_JITAI_STATUSES.includes(person.progressStatus || '')
                  ? 'border-rose-200 bg-rose-50 text-rose-700'
                  : 'border-[#E2E8F0] bg-[#F8FAFC] text-[#0F172A]'
              }`}
            >
              {PROGRESS_STATUSES.map((st) => (
                <option key={st} value={st}>
                  {st}
                </option>
              ))}
            </select>

            <button
              type="button"
              onClick={() => handleMarkContactedToday(person, isSchool)}
              className={`rounded-md border px-2.5 py-1.5 text-[11px] font-semibold transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                isContactedToday
                  ? 'border-emerald-200 bg-emerald-50 text-emerald-700'
                  : 'border-[#E2E8F0] bg-white text-[#475569] hover:border-[#CBD5E1] hover:bg-[#F8FAFC] hover:text-[#0F172A]'
              }`}
            >
              {isContactedToday ? t.contactedTodayBtn : t.markContactedTodayBtn}
            </button>

            <button
              type="button"
              onClick={() =>
                setReplyModal({
                  person,
                  scenario: missingItems.length > 0 ? 'first_mendan' : 'recommend_job',
                })
              }
              className="rounded-md bg-[#16A34A] px-3 py-1.5 text-[11px] font-semibold text-white transition-colors hover:bg-[#15803D] whitespace-nowrap shrink-0 cursor-pointer"
            >
              {t.suggestReplyBtn}
            </button>

            {(() => {
              const hasSavedCv = Boolean(person.savedRirekisho?.draft);
              return (
                <button
                  type="button"
                  onClick={() => openRirekishoModalForCandidate(person)}
                  className={`inline-flex items-center gap-1 rounded-md border px-3 py-1.5 text-[11px] font-semibold whitespace-nowrap shrink-0 cursor-pointer transition-colors ${
                    hasSavedCv
                      ? 'border-emerald-300 bg-emerald-50 text-emerald-800 hover:bg-emerald-100'
                      : 'border-[#E2E8F0] bg-white text-[#0F172A] hover:border-[#CBD5E1] hover:bg-[#F8FAFC]'
                  }`}
                  title={
                    hasSavedCv && person.savedRirekisho?.updatedAt
                      ? `${t.editSavedRirekishoBtn} (${person.savedRirekisho.updatedAt
                          .slice(0, 16)
                          .replace('T', ' ')})`
                      : t.createRirekishoBtn
                  }
                >
                  {hasSavedCv ? `✅ ${t.editSavedRirekishoBtn}` : t.createRirekishoBtn}
                </button>
              );
            })()}

            <button
              type="button"
              disabled={syncingCalendarId === person.id}
              onClick={() => requestSyncCandidateCalendar(person)}
              className={`inline-flex items-center gap-1 rounded-md border px-2.5 py-1.5 text-[11px] font-semibold transition-colors whitespace-nowrap shrink-0 cursor-pointer disabled:opacity-50 ${
                person.googleCalendarEventId
                  ? 'border-emerald-200 bg-emerald-50 text-emerald-700 hover:bg-emerald-100'
                  : 'border-[#E2E8F0] bg-white text-[#475569] hover:border-emerald-300 hover:text-emerald-700'
              }`}
              title={
                person.googleCalendarEventId
                  ? t.syncCalendarTitleSynced
                  : t.syncCalendarTitleNew
              }
            >
              <Calendar className="h-3 w-3" aria-hidden="true" />
              {syncingCalendarId === person.id
                ? t.syncingCalendarBtn
                : person.googleCalendarEventId
                ? t.syncedCalendarBtn
                : t.syncCalendarBtn}
            </button>

            {(() => {
              const fbMessengerUrl = getSafeFacebookMessengerUrl(person);
              if (fbMessengerUrl) {
                return (
                  <a
                    href={fbMessengerUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-1 rounded-md border border-sky-200 bg-sky-50 px-2.5 py-1.5 text-[11px] font-semibold text-sky-700 hover:bg-sky-100 whitespace-nowrap shrink-0"
                    title={t.openMessengerBtn}
                  >
                    <MessageCircle className="h-3 w-3" aria-hidden="true" />
                    {t.openMessengerBtn}
                  </a>
                );
              }
              if (person.fbSenderId) {
                return (
                  <a
                    href="https://business.facebook.com/latest/inbox/all"
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-1 rounded-md border border-sky-200 bg-sky-50 px-2.5 py-1.5 text-[11px] font-semibold text-sky-700 hover:bg-sky-100 whitespace-nowrap shrink-0"
                    title={`Meta Business (Sender ID: ${person.fbSenderId})`}
                  >
                    <MessageCircle className="h-3 w-3" aria-hidden="true" />
                    {t.fbInboxBtn}
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
                  className="inline-flex items-center gap-1 rounded-md border border-[#E2E8F0] bg-white px-2 py-1.5 text-[11px] font-semibold text-[#0F172A] hover:border-emerald-300 hover:text-emerald-700 whitespace-nowrap shrink-0"
                  title={t.openCalendarBtn}
                >
                  <ExternalLink className="h-3 w-3" aria-hidden="true" />
                  {t.openCalendarBtn}
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
              className="rounded-md border border-[#E2E8F0] bg-white px-2.5 py-1.5 text-[11px] font-semibold text-[#475569] hover:border-[#CBD5E1] hover:bg-[#F8FAFC] hover:text-[#0F172A] whitespace-nowrap shrink-0 cursor-pointer"
            >
              {t.editBtn}
            </button>

            <button
              type="button"
              onClick={() => handleDeletePerson(person.id, isSchool)}
              className="rounded-md border border-rose-200 bg-rose-50 px-2.5 py-1.5 text-[11px] font-semibold text-rose-700 hover:bg-rose-100 whitespace-nowrap shrink-0 cursor-pointer"
            >
              {t.deleteBtn}
            </button>
          </div>
        </div>

        {/* Structured Key-Value Grid */}
        <div className="mt-3 grid grid-cols-2 gap-3 border-t border-dashed border-[#E2E8F0] pt-3 sm:grid-cols-3 lg:grid-cols-5">
          <div>
            <div className="font-mono text-[10px] font-medium uppercase tracking-wider text-[#64748B]">
              {t.cardLabels.fbName}
            </div>
            <div className="mt-0.5 break-words text-xs font-medium text-[#0F172A]">
              {person.fbName || displayName}
            </div>
          </div>
          <div>
            <div className="font-mono text-[10px] font-medium uppercase tracking-wider text-[#64748B]">
              {t.cardLabels.natVisa}
            </div>
            <div className="mt-0.5 text-xs font-medium text-[#0F172A] tabular-nums">
              {person.nationality || t.cardLabels.unset}
              {person.age ? ` / ${person.age} ${t.cardLabels.ageUnit}` : ''}
            </div>
            {person.visaStatus && (
              <div className="mt-0.5 text-[11px] text-[#64748B]">{person.visaStatus}</div>
            )}
          </div>
          <div>
            <div className="font-mono text-[10px] font-medium uppercase tracking-wider text-[#64748B]">
              {t.cardLabels.langCert}
            </div>
            <div className="mt-0.5 font-mono text-xs font-semibold text-emerald-700">
              {person.jlpt || t.cardLabels.unset}
            </div>
            {(person.passedExam || person.japanExperience) && (
              <div className="mt-0.5 text-[11px] text-[#64748B]">
                {[person.japanExperience, person.passedExam].filter(Boolean).join(' · ')}
              </div>
            )}
          </div>
          <div>
            <div className="font-mono text-[10px] font-medium uppercase tracking-wider text-[#64748B]">
              {t.cardLabels.fieldArea}
            </div>
            <div className="mt-0.5 break-words text-xs font-medium text-[#0F172A]">
              {person.desiredIndustry || t.cardLabels.unset}
              {person.desiredArea ? ` (${person.desiredArea})` : ''}
            </div>
          </div>
          <div>
            <div className="font-mono text-[10px] font-medium uppercase tracking-wider text-[#64748B]">
              {person.visaExpiry
                ? t.cardLabels.visaExp
                : mendanInfo.hasMendan
                ? t.cardLabels.mendan
                : t.cardLabels.school}
            </div>
            <div
              className={`mt-0.5 font-mono text-xs font-medium tabular-nums ${
                person.visaExpiry
                  ? 'text-rose-700'
                  : mendanInfo.hasMendan
                  ? 'text-amber-700'
                  : 'text-[#64748B]'
              }`}
            >
              {person.visaExpiry || mendanInfo.dateLabel || person.schoolName || '-'}
            </div>
          </div>
        </div>

        {(person.naiteiCompany || person.expectedJoinDate || person.visaSubmissionDeadline) && (
          <div className="mt-2.5 flex flex-wrap items-center gap-x-4 gap-y-1.5 rounded-md border border-emerald-200 bg-emerald-50/70 px-3 py-2 text-xs">
            <span className="font-mono text-[10px] font-bold uppercase tracking-wider text-emerald-800">
              🎉 {t.drawer.naiteiBannerTitle}:
            </span>
            {person.naiteiCompany && (
              <span className="font-semibold text-[#0F172A]">
                {t.drawer.naiteiComp} {person.naiteiCompany}
              </span>
            )}
            {person.expectedJoinDate && (
              <span className="font-mono font-medium text-emerald-800 tabular-nums">
                {t.drawer.expectedJoin} {person.expectedJoinDate}
              </span>
            )}
            {person.visaSubmissionDeadline && (
              <span className="font-mono font-semibold text-amber-800 tabular-nums">
                {t.drawer.visaDeadline} {person.visaSubmissionDeadline}
              </span>
            )}
          </div>
        )}

        {(person.conversionGoal || person.salaryExpectation) && (
          <div className="mt-2.5 grid grid-cols-1 gap-2 rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-2 text-xs sm:grid-cols-2">
            {person.conversionGoal && (
              <div>
                <span className="font-mono text-[10px] font-semibold uppercase text-sky-700">
                  {t.cardLabels.goal}{' '}
                </span>
                <span className="text-[#0F172A]">{person.conversionGoal}</span>
              </div>
            )}
            {person.salaryExpectation && (
              <div>
                <span className="font-mono text-[10px] font-semibold uppercase text-emerald-700">
                  {t.cardLabels.salaryExp}{' '}
                </span>
                <span className="text-[#0F172A]">{person.salaryExpectation}</span>
              </div>
            )}
          </div>
        )}

        {(person.negotiationLog || person.memo) && (
          <div className="mt-2.5 whitespace-pre-line rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-2.5 text-xs leading-relaxed text-[#334155]">
            {person.negotiationLog && person.negotiationLog !== person.memo
              ? `${person.negotiationLog}${
                  person.memo ? `\n\n--- ${t.cardLabels.extraNotes} ---\n${person.memo}` : ''
                }`
              : person.memo || person.negotiationLog}
          </div>
        )}

        {/* Next Best Action Bar */}
        {(() => {
          const isInterviewPassed =
            Boolean(person.interviewDate && parseFlexibleDateToTimestamp(person.interviewDate) < Date.now()) ||
            /面接|内定/.test(String(person.progressStatus || ''));
          return (
            <div className="mt-2.5 flex flex-wrap items-center justify-between gap-2 rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-2 text-xs">
              <div className="flex flex-wrap items-center gap-2">
                <span className="font-mono text-[10px] font-semibold uppercase tracking-wider text-emerald-700">
                  {t.cardLabels.nextAction}
                </span>
                {missingItems.length > 0 && (
                  <span className="font-medium text-amber-800">
                    {t.cardLabels.missingCount(
                      missingItems.length,
                      missingItems.map((m) => (uiLang === 'jp' ? m.ja : m.vi)).join(', ')
                    )}
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
                    className="rounded border border-amber-300 bg-amber-50 px-2.5 py-1 text-[11px] font-semibold text-amber-800 hover:bg-amber-100 whitespace-nowrap cursor-pointer"
                  >
                    {t.cardLabels.askMissingBtn}
                  </button>
                )}
                <button
                  type="button"
                  onClick={() => setReplyModal({ person, scenario: 'recommend_job' })}
                  className="rounded border border-emerald-200 bg-emerald-50 px-2.5 py-1 text-[11px] font-semibold text-emerald-700 hover:bg-emerald-100 whitespace-nowrap cursor-pointer"
                >
                  {t.cardLabels.send3JobsBtn}
                </button>
                {isInterviewPassed && (
                  <button
                    type="button"
                    onClick={() =>
                      setReplyModal({ person, scenario: 'post_interview_followup' })
                    }
                    className="rounded border border-sky-200 bg-sky-50 px-2.5 py-1 text-[11px] font-semibold text-sky-700 hover:bg-sky-100 whitespace-nowrap cursor-pointer"
                  >
                    {t.cardLabels.postInterviewBtn}
                  </button>
                )}
              </div>
            </div>
          );
        })()}

        {declineReason && (
          <div className="mt-2.5 flex flex-wrap items-center justify-between gap-2 rounded-md border border-rose-200 bg-rose-50 px-3 py-2 text-xs">
            <span className="font-semibold text-rose-700">
              {t.cardLabels.declineReason}{' '}
              <span className="font-normal text-[#0F172A]">{declineReason}</span>
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
              className="rounded border border-rose-200 bg-white px-2 py-0.5 text-[11px] font-semibold text-rose-700 hover:bg-rose-100 whitespace-nowrap shrink-0 cursor-pointer"
            >
              {t.cardLabels.editDeclineReason}
            </button>
          </div>
        )}

        {(matchingKyujins.length > 0 || Boolean(matchEval.noticeMessage)) &&
          (() => {
            const isInterviewOrWaiting = [
              '1次面接予定',
              '2次面接予定',
              '面接調整中',
              '結果待ち',
            ].includes(person.progressStatus || '');
            const showMatches = !isInterviewOrWaiting || isJobMatchExpanded;

            return (
              <div className="mt-3 border-t border-[#E2E8F0] pt-2.5">
                <div className="mb-1.5 flex flex-wrap items-center justify-between gap-2">
                  <span className="font-mono text-[10px] font-semibold uppercase tracking-wider text-[#64748B]">
                    {t.jobMatch.title} ({matchingKyujins.length})
                  </span>
                  <div className="flex items-center gap-1.5">
                    {isInterviewOrWaiting && matchingKyujins.length > 0 && (
                      <button
                        type="button"
                        onClick={() =>
                          setExpandedJobMatchIds((prev) => ({
                            ...prev,
                            [person.id]: !prev[person.id],
                          }))
                        }
                        className="rounded border border-sky-200 bg-sky-50 px-2 py-0.5 text-[11px] font-semibold text-sky-700 hover:bg-sky-100 whitespace-nowrap cursor-pointer"
                      >
                        {isJobMatchExpanded
                          ? t.jobMatch.hideMatchesBtn
                          : t.jobMatch.showMatchesBtn(matchingKyujins.length)}
                      </button>
                    )}
                    {!isInterviewOrWaiting && matchingKyujins.length > 2 && (
                      <button
                        type="button"
                        onClick={() =>
                          setExpandedJobMatchIds((prev) => ({
                            ...prev,
                            [person.id]: !prev[person.id],
                          }))
                        }
                        className="rounded border border-[#E2E8F0] bg-[#F8FAFC] px-2 py-0.5 text-[11px] font-semibold text-sky-700 hover:border-sky-300 hover:bg-sky-50 whitespace-nowrap cursor-pointer"
                      >
                        {isJobMatchExpanded
                          ? t.jobMatch.collapse
                          : t.jobMatch.showMore(matchingKyujins.length - 2)}
                      </button>
                    )}
                  </div>
                </div>
                {showMatches && (
                  <>
                    {matchEval.noticeMessage && (
                      <div
                        className={`mb-2 rounded-md border px-3 py-2 text-[11px] ${
                          matchingKyujins.length > 0
                            ? 'border-amber-300 bg-amber-50 text-amber-800'
                            : 'border-[#E2E8F0] bg-[#F8FAFC] text-[#64748B]'
                        }`}
                      >
                        {matchEval.noticeMessage}
                      </div>
                    )}
                    {visibleMatches.length > 0 && (
                      <div className="space-y-1.5">
                        {visibleMatches.map(({ company, slot, score, reasons }) => {
                          const isBaito = /baito|アルバイト|パート/i.test(
                            slot.employmentType || company.employmentType || ''
                          );
                          const isAlreadyProposedHere =
                            isActualProposedRecord(person) &&
                            String(proposedComp || '')
                              .toLowerCase()
                              .includes(company.companyName.toLowerCase());

                          return (
                            <div
                              key={slot.slotId}
                              className="flex flex-wrap items-center justify-between gap-2 rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-2 text-xs"
                            >
                              <div className="min-w-0 flex-1 space-y-1">
                                <div className="flex flex-wrap items-center gap-2">
                                  <span
                                    className={`rounded border px-1.5 py-0.5 font-mono text-[10px] font-bold tabular-nums ${
                                      score >= 75
                                        ? 'border-emerald-200 bg-emerald-50 text-emerald-700'
                                        : score >= 55
                                        ? 'border-sky-200 bg-sky-50 text-sky-700'
                                        : 'border-amber-300 bg-amber-50 text-amber-800'
                                    }`}
                                  >
                                    {t.jobMatch.matchPct(score)}
                                  </span>
                                  <span className="font-semibold text-[#0F172A]">
                                    {company.companyName}
                                  </span>
                                  {isAlreadyProposedHere && (
                                    <span className="rounded border border-emerald-200 bg-emerald-50 px-1.5 py-0.5 text-[10px] font-bold text-emerald-700">
                                      {t.jobMatch.alreadyProposedTag}
                                    </span>
                                  )}
                                  <span
                                    className={`rounded px-1.5 py-0.5 font-mono text-[10px] font-medium ${
                                      isBaito
                                        ? 'bg-sky-50 text-sky-700 border border-sky-200'
                                        : 'bg-slate-200/70 text-[#475569]'
                                    }`}
                                  >
                                    {isBaito ? 'Baito' : slot.employmentType || 'Tokutei'}
                                  </span>
                                  <span className="text-[#475569]">
                                    {formatPublicArea(slot.area)}
                                  </span>
                                  <span className="font-mono font-semibold text-emerald-700 tabular-nums">
                                    {(() => {
                                      const m = getKyujinSlotSalaryDisplayMeta(
                                        slot,
                                        slot.employmentType || company.employmentType || ''
                                      );
                                      return m.isHourly
                                        ? `${t.jobMatch.hourlyPrefix} ${m.displayValue}`
                                        : m.displayValue;
                                    })()}
                                  </span>
                                  <span className="font-mono text-[11px] text-[#64748B]">
                                    {slot.minJlpt}
                                  </span>
                                </div>
                                {reasons && reasons.length > 0 && (
                                  <div className="flex flex-wrap gap-1.5">
                                    {reasons.map((rs, rIdx) => (
                                      <span
                                        key={rIdx}
                                        className="rounded border border-[#E2E8F0] bg-white px-1.5 py-0.5 text-[10px] text-[#475569]"
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
                                    const post = buildAnonymousJobPostText(company, slot, uiLang);
                                    const ok = await copyTextSafely(post);
                                    showToast(
                                      ok ? t.toasts.copiedInvite(score) : t.toasts.copyFailed
                                    );
                                  }}
                                  className="rounded border border-sky-200 bg-sky-50 px-2.5 py-1 text-[11px] font-semibold text-sky-700 hover:bg-sky-100 whitespace-nowrap cursor-pointer"
                                >
                                  {t.jobMatch.createInviteBtn}
                                </button>
                                <button
                                  type="button"
                                  onClick={() => setPreviewModal({ company, slot })}
                                  className="rounded border border-[#E2E8F0] bg-white px-2.5 py-1 text-[11px] font-semibold text-[#475569] hover:border-[#CBD5E1] hover:text-[#0F172A] whitespace-nowrap cursor-pointer"
                                >
                                  {t.jobMatch.previewAnonBtn}
                                </button>
                                <button
                                  type="button"
                                  disabled={isAlreadyProposedHere}
                                  onClick={() =>
                                    handleProposeCompanyToCandidate(
                                      person,
                                      isSchool,
                                      company.companyName
                                    )
                                  }
                                  className={`rounded border px-2.5 py-1 text-[11px] font-semibold whitespace-nowrap ${
                                    isAlreadyProposedHere
                                      ? 'border-emerald-200 bg-emerald-50 text-emerald-700 opacity-70 cursor-default'
                                      : 'border-emerald-200 bg-emerald-50 text-emerald-700 hover:bg-emerald-100 cursor-pointer'
                                  }`}
                                >
                                  {isAlreadyProposedHere
                                    ? t.jobMatch.alreadyProposedTag
                                    : t.jobMatch.proposeBtn}
                                </button>
                              </div>
                            </div>
                          );
                        })}
                        {!isInterviewOrWaiting &&
                          matchingKyujins.length > 2 &&
                          !isJobMatchExpanded && (
                            <button
                              type="button"
                              onClick={() =>
                                setExpandedJobMatchIds((prev) => ({
                                  ...prev,
                                  [person.id]: true,
                                }))
                              }
                              className="w-full rounded-md border border-dashed border-[#CBD5E1] bg-white py-1.5 text-center text-[11px] font-semibold text-sky-700 hover:border-sky-300 hover:bg-sky-50 cursor-pointer"
                            >
                              {t.jobMatch.showMore(matchingKyujins.length - 2)}
                            </button>
                          )}
                      </div>
                    )}
                  </>
                )}
              </div>
            );
          })()}
      </div>
    );
  };

  const crmSummaryForAi = useMemo(() => {
    const statusMap: Record<string, number> = {};
    allPeopleCombined.forEach((p) => {
      const st = String(p.progressStatus || '新規相談').trim();
      statusMap[st] = (statusMap[st] || 0) + 1;
    });
    const statusSummary = Object.entries(statusMap)
      .sort((a, b) => b[1] - a[1])
      .map(([st, cnt]) => `${st}:${cnt}`)
      .join(', ');

    const unconfirmedResigCount = resignations
      .map((r) => normalizeAndAutoComputeResignationDates(r))
      .filter(hasUnconfirmedResignationDate).length;

    const openCompanies = kyujins
      .filter((c) => c.recruitmentStatus !== '募集終了')
      .slice(0, 10)
      .map(
        (c) =>
          `${c.companyName} (${c.industry || '未設定'}, ${c.areaBreakdown?.[0]?.area || '全国'}, ${c.areaBreakdown?.length || 0}枠)`
      )
      .join('; ');

    const upcomingList = upcomingMendanAlerts
      .slice(0, 6)
      .map(
        ({ person, mendan }) =>
          `${getCandidateDisplayName(person, uiLang)} (${person.jlpt || '要確認'}, ${mendan.dateLabel || person.progressStatus || '調整中'})`
      )
      .join('; ');

    return [
      `=== CRM LIVE EXACT STATS (DO NOT GUESS OR ALTER NUMBERS) ===`,
      `総登録人材: ${allPeopleCombined.length}名 (FB・SNS・その他: ${nonSchoolCandidates.length}名 [FB個人:${personalCandidates.length}, FB会社公式:${companyCandidates.length}, LINE:${lineCandidates.length}, その他:${otherCandidates.length}] / 日本語学校・留学生: ${schoolCandidates.length}名).`,
      `進捗ステータス内訳: ${statusSummary || 'なし'}. 推薦済み(実質): ${allProposedPeople.length}名 (進行中:${proposedStageCounts.in_progress}, 内定:${proposedStageCounts.naitei}, 辞退・終了:${proposedStageCounts.jitai}).`,
      `求人データベース: ${kyujinStats.totalCompanies}社・計${kyujinStats.totalSlots}ポジション (募集中:${kyujinStats.openSlots}枠, 募集終了:${kyujinStats.closedSlots}枠, 賞与あり:${kyujinStats.bonusSlots}枠, アルバイト:${kyujinStats.baitoSlots}枠).`,
      `退職者管理: 計${resignations.length}名 (期限アラート:${resignationAlerts.length}名, 期限<=3日/超過:${urgentResignationAlerts.length}名, 退職日未確定:${unconfirmedResigCount}名).`,
      `提携PT機関数: ${ptPartners.length}社. 要至急フォロー対象: ${urgentFollowUpList.length}名.`,
      `直近の面談・面接予定 (${upcomingMendanAlerts.length}名): ${upcomingList || 'なし'}.`,
      `主な募集中企業: ${openCompanies || 'なし'}.`,
    ].join('\n');
  }, [
    allPeopleCombined,
    nonSchoolCandidates.length,
    personalCandidates.length,
    companyCandidates.length,
    lineCandidates.length,
    otherCandidates.length,
    schoolCandidates.length,
    allProposedPeople.length,
    proposedStageCounts,
    kyujinStats,
    kyujins,
    resignations,
    resignationAlerts.length,
    urgentResignationAlerts.length,
    ptPartners.length,
    urgentFollowUpList.length,
    upcomingMendanAlerts,
    uiLang,
  ]);

  return (
    <div className="min-h-screen bg-[#F8FAFC] text-[#0F172A]">
      {toastMsg && (
        <div className="fixed bottom-20 right-5 z-[10000] max-w-md rounded-lg border border-emerald-200 bg-white px-4 py-3 text-xs font-bold text-emerald-700 shadow-lg">
          {toastMsg}
        </div>
      )}

      {backupRestoreError && (
        <div
          role="alert"
          className="fixed bottom-5 right-5 z-[10001] max-w-lg rounded-xl border border-rose-300 bg-white p-4 shadow-xl"
        >
          <div className="flex items-start justify-between gap-3">
            <div className="flex items-start gap-2.5">
              <AlertTriangle className="mt-0.5 h-4 w-4 shrink-0 text-rose-600" aria-hidden="true" />
              <div className="space-y-1">
                <p className="text-xs font-bold text-rose-700">
                  {uiLang === 'jp'
                    ? 'データ復元（データ復元）に失敗しました'
                    : 'Lỗi chi tiết khi khôi phục file backup (データ復元)'}
                </p>
                <p className="whitespace-pre-wrap break-words text-[11px] leading-relaxed text-[#334155]">
                  {backupRestoreError}
                </p>
              </div>
            </div>
            <button
              type="button"
              onClick={() => setBackupRestoreError(null)}
              className="rounded p-1 text-[#64748B] hover:bg-[#F1F5F9] hover:text-[#0F172A] cursor-pointer"
              aria-label="Close"
            >
              <X className="h-3.5 w-3.5" />
            </button>
          </div>
        </div>
      )}

      {/* Hidden File Inputs */}
      <input
        ref={backupInputRef}
        type="file"
        accept=".json,.csv,application/json,text/csv"
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
      <input
        ref={rirekishoPhotoInputRef}
        type="file"
        accept="image/*"
        onChange={handleUploadRirekishoPhoto}
        className="hidden"
      />

      {viewMode === 'home' ? (
        <HomePage
          candidateCount={nonSchoolCandidates.length}
          schoolCount={schoolCandidates.length}
          companyCount={kyujinStats.totalCompanies}
          slotCount={kyujinStats.totalSlots}
          ptCount={ptPartners.length}
          upcomingMendanCount={upcomingMendanAlerts.length}
          userEmail={authUser?.email || null}
          authLoading={authLoading}
          lang={uiLang}
          onChangeLang={setUiLang}
          onStartWorkspace={(targetTab) => {
            if (targetTab) setActiveTab(targetTab);
            setViewMode('workspace');
          }}
          onSignInGoogle={handleSignInGoogle}
          onSignOut={handleSignOut}
          onOpenChatbot={() => setChatbotOpen(true)}
        />
      ) : (
        <div className="flex min-h-screen flex-col lg:grid lg:h-screen lg:grid-cols-[240px_1fr_320px] lg:grid-rows-[64px_1fr] lg:overflow-hidden print:block print:h-auto print:overflow-visible">
          {/* TOP HEADER BAR */}
          <header className="flex min-h-[64px] flex-wrap items-center justify-between gap-3 border-b border-[#E2E8F0] bg-white px-5 py-2 shadow-[0_1px_2px_rgba(0,0,0,0.03)] lg:col-span-3 print:hidden">
            <div className="flex flex-wrap items-center gap-3">
              <button
                type="button"
                onClick={() => setViewMode('home')}
                className="flex items-center gap-2 text-sm font-bold tracking-tight text-[#0F172A] hover:text-[#16A34A] transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                title={t.homeTooltip}
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
                {t.brandTitle}
              </button>

              <button
                type="button"
                onClick={() => {
                  const next = !gasWriteLocked;
                  setGasWriteLocked(next);
                  showToast(next ? t.toasts.gasLockedOn : t.toasts.gasLockedOff);
                }}
                className={`rounded px-2 py-0.5 font-mono text-[10px] font-bold uppercase tracking-wider transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                  gasWriteLocked
                    ? 'bg-emerald-50 border border-emerald-200 text-emerald-700'
                    : 'border border-amber-300 bg-amber-50 text-amber-800'
                }`}
              >
                {gasWriteLocked ? t.gasLocked : t.gasUnlocked}
              </button>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              <span className="hidden font-mono text-[10px] uppercase tracking-wider text-[#64748B] xl:inline tabular-nums">
                {t.headerSummary(
                  candidates.length,
                  schoolStudents.length,
                  kyujinStats.totalCompanies,
                  kyujinStats.totalSlots
                )}
              </span>

              <button
                type="button"
                onClick={() => backupInputRef.current?.click()}
                className="rounded-md border border-[#E2E8F0] bg-white px-3 py-1.5 text-[11px] font-semibold text-[#475569] hover:border-[#CBD5E1] hover:bg-[#F8FAFC] hover:text-[#0F172A] whitespace-nowrap shrink-0 cursor-pointer"
              >
                {t.restoreData}
              </button>

              <button
                type="button"
                onClick={handleExportBackupJson}
                className="rounded-md border border-[#E2E8F0] bg-white px-3 py-1.5 text-[11px] font-semibold text-[#475569] hover:border-[#CBD5E1] hover:bg-[#F8FAFC] hover:text-[#0F172A] whitespace-nowrap shrink-0 cursor-pointer"
              >
                {t.backupJson}
              </button>

              <button
                type="button"
                onClick={openMessengerLinkModal}
                className={`inline-flex items-center gap-1.5 rounded-md border px-3 py-1.5 text-[11px] font-semibold transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                  webhookStatus?.configured || fbPageLink
                    ? 'border-sky-200 bg-sky-50 text-sky-700 hover:bg-sky-100'
                    : 'border-[#E2E8F0] bg-white text-[#0F172A] hover:border-sky-300'
                }`}
                title={t.linkMessenger}
              >
                <span
                  className={`inline-block h-1.5 w-1.5 rounded-full ${
                    webhookStatus?.configured
                      ? 'bg-[#16A34A]'
                      : fbPageLink
                      ? 'bg-sky-500'
                      : 'bg-amber-500'
                  }`}
                />
                <Link2 className="h-3.5 w-3.5" aria-hidden="true" />
                {t.linkMessenger}
              </button>

              {/* Language Switcher [VI | JP] */}
              <div
                className="inline-flex items-center rounded-md border border-[#E2E8F0] bg-[#F8FAFC] p-0.5 shrink-0"
                role="group"
                aria-label="Language switcher"
              >
                <button
                  type="button"
                  onClick={() => setUiLang('vi')}
                  className={`rounded px-2 py-1 text-[11px] font-semibold transition-colors whitespace-nowrap cursor-pointer ${
                    uiLang === 'vi'
                      ? 'bg-white text-[#0F172A] shadow-sm'
                      : 'text-[#64748B] hover:text-[#0F172A]'
                  }`}
                >
                  🇻🇳 VI
                </button>
                <button
                  type="button"
                  onClick={() => setUiLang('jp')}
                  className={`rounded px-2 py-1 text-[11px] font-semibold transition-colors whitespace-nowrap cursor-pointer ${
                    uiLang === 'jp'
                      ? 'bg-white text-[#0F172A] shadow-sm'
                      : 'text-[#64748B] hover:text-[#0F172A]'
                  }`}
                >
                  🇯🇵 JP
                </button>
              </div>

              <button
                type="button"
                onClick={() => setChatbotOpen(true)}
                className="inline-flex items-center gap-1.5 rounded-md bg-[#16A34A] px-3 py-1.5 text-[11px] font-semibold text-white hover:bg-[#15803D] whitespace-nowrap shrink-0 cursor-pointer"
              >
                <MessageSquareText className="h-3.5 w-3.5" aria-hidden="true" />
                {t.aiAssistant}
              </button>
            </div>
          </header>

          {/* MOBILE / TABLET TAB STRIP (< lg) */}
          <div className="flex items-center gap-1 overflow-x-auto border-b border-[#E2E8F0] bg-white px-4 py-2 lg:hidden print:hidden">
            {(
              [
                { id: 'all', label: `${t.tabs.allMobile} (${allPeopleCombined.length})` },
                { id: 'personal', label: `${t.tabs.personal} (${personalCandidates.length})` },
                { id: 'company', label: `${t.tabs.company} (${companyCandidates.length})` },
                { id: 'line', label: `${t.tabs.line} (${lineCandidates.length})` },
                { id: 'other', label: `${t.tabs.other} (${otherCandidates.length})` },
                { id: 'school', label: `${t.tabs.schoolMobile} (${schoolCandidates.length})` },
                { id: 'mendan', label: `${t.tabs.mendan} (${mendanPeopleSorted.length})` },
                { id: 'proposed', label: `${t.tabs.proposed} (${allProposedPeople.length})` },
                { id: 'resignation', label: `${t.tabs.resignation} (${resignations.length})` },
                { id: 'pt', label: `${t.tabs.pt} (${ptPartners.length})` },
                { id: 'kyujin', label: `${t.tabs.kyujin} (${kyujinStats.totalCompanies})` },
                { id: 'report', label: t.tabs.report },
              ] as Array<{ id: MainTabId; label: string }>
            ).map((tab) => (
              <button
                key={tab.id}
                type="button"
                onClick={() => setActiveTab(tab.id)}
                className={`rounded-md px-2.5 py-1 text-xs font-semibold whitespace-nowrap shrink-0 tabular-nums cursor-pointer ${
                  activeTab === tab.id
                    ? 'bg-[#16A34A] text-white'
                    : 'text-[#64748B] hover:bg-[#F8FAFC] hover:text-[#0F172A]'
                }`}
              >
                {tab.label}
              </button>
            ))}
          </div>

          {/* LEFT SIDEBAR (DESKTOP) */}
          <aside className="hidden border-r border-[#E2E8F0] bg-white py-4 lg:flex lg:flex-col lg:overflow-y-auto print:hidden">
            <div className="mb-6">
              <div className="px-5 pb-2 font-mono text-[10px] font-semibold uppercase tracking-widest text-[#64748B]">
                {t.sidebarDatabase}
              </div>
              {(
                [
                  { id: 'all', label: t.tabs.allDesktop, count: `${allPeopleCombined.length}` },
                  { id: 'personal', label: t.tabs.personal, count: `${personalCandidates.length}` },
                  { id: 'company', label: t.tabs.company, count: `${companyCandidates.length}` },
                  { id: 'line', label: t.tabs.line, count: `${lineCandidates.length}` },
                  { id: 'other', label: t.tabs.other, count: `${otherCandidates.length}` },
                  { id: 'school', label: t.tabs.school, count: `${schoolCandidates.length}` },
                ] as Array<{ id: MainTabId; label: string; count: string }>
              ).map((item) => (
                <button
                  key={item.id}
                  type="button"
                  onClick={() => setActiveTab(item.id)}
                  className={`flex w-full items-center justify-between px-5 py-2.5 text-left text-xs transition-colors cursor-pointer ${
                    activeTab === item.id
                      ? 'border-r-2 border-[#16A34A] bg-emerald-50 font-semibold text-emerald-700'
                      : 'text-[#64748B] hover:bg-[#F8FAFC] hover:text-[#0F172A]'
                  }`}
                >
                  <span>{item.label}</span>
                  <span className="font-mono text-[11px] tabular-nums">{item.count}</span>
                </button>
              ))}
            </div>

            <div className="mb-6">
              <div className="px-5 pb-2 font-mono text-[10px] font-semibold uppercase tracking-widest text-[#64748B]">
                {t.sidebarWorkflow}
              </div>
              {(
                [
                  { id: 'mendan', label: t.tabs.mendan, count: `${mendanPeopleSorted.length}` },
                  { id: 'proposed', label: t.tabs.proposed, count: `${allProposedPeople.length}` },
                  {
                    id: 'resignation',
                    label: t.tabs.resignation,
                    count:
                      resignationAlerts.length > 0
                        ? `${resignationAlerts.length} ${t.urgentShort} / ${resignations.length}`
                        : `${resignations.length}`,
                  },
                  { id: 'pt', label: t.tabs.pt, count: `${ptPartners.length}` },
                  { id: 'kyujin', label: t.tabs.kyujin, count: `${kyujinStats.totalCompanies}` },
                  { id: 'report', label: t.tabs.reportDesktop, count: `${kyujinStats.totalSlots}` },
                ] as Array<{ id: MainTabId; label: string; count: string }>
              ).map((item) => (
                <button
                  key={item.id}
                  type="button"
                  onClick={() => setActiveTab(item.id)}
                  className={`flex w-full items-center justify-between px-5 py-2.5 text-left text-xs transition-colors cursor-pointer ${
                    activeTab === item.id
                      ? 'border-r-2 border-[#16A34A] bg-emerald-50 font-semibold text-emerald-700'
                      : 'text-[#64748B] hover:bg-[#F8FAFC] hover:text-[#0F172A]'
                  }`}
                >
                  <span>{item.label}</span>
                  <span className="font-mono text-[11px] tabular-nums">{item.count}</span>
                </button>
              ))}
            </div>

            <div className="mt-auto space-y-2 border-t border-[#E2E8F0] p-4">
              {authUser ? (
                <button
                  type="button"
                  onClick={handleSignOut}
                  title={authUser.email || ''}
                  className="w-full rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-2 text-[11px] font-semibold text-[#0F172A] hover:border-[#CBD5E1] truncate cursor-pointer"
                >
                  {authUser.email ? `${authUser.email.split('@')[0]} · ${t.signOut}` : t.signOut}
                </button>
              ) : (
                <button
                  type="button"
                  disabled={authLoading}
                  onClick={handleSignInGoogle}
                  className="w-full rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-2 text-[11px] font-semibold text-[#0F172A] hover:border-[#16A34A] disabled:opacity-60 cursor-pointer"
                >
                  {authLoading ? t.signingIn : t.signInGoogle}
                </button>
              )}

              <button
                type="button"
                onClick={() => setViewMode('home')}
                className="w-full rounded-md border border-transparent px-3 py-1.5 text-[11px] font-medium text-[#64748B] hover:border-[#E2E8F0] hover:text-[#0F172A] cursor-pointer"
              >
                {t.homeLink}
              </button>
            </div>
          </aside>

          {/* CENTER MAIN PANE */}
          <main className="min-w-0 bg-[#F8FAFC] p-5 lg:overflow-y-auto lg:border-r lg:border-[#E2E8F0] print:overflow-visible print:p-0 print:border-0">
            {['all', 'personal', 'company', 'line', 'other', 'school', 'mendan', 'proposed'].includes(
              activeTab
            ) && (
              <div className="space-y-4">
                {/* Control Row + Filter Bar */}
                <div className="space-y-3 rounded-lg border border-[#E2E8F0] bg-white p-3.5 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
                  <div className="flex flex-wrap items-center gap-2">
                    <input
                      type="text"
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      placeholder={t.searchPlaceholder}
                      className="min-w-[220px] flex-1 rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3.5 py-2 text-xs text-[#0F172A] placeholder:text-[#94A3B8] focus:border-[#16A34A] focus:bg-white focus:outline-none"
                    />

                    <select
                      value={statusFilter}
                      onChange={(e) => setStatusFilter(e.target.value)}
                      className="rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-2 text-xs font-semibold text-[#0F172A] focus:border-[#16A34A] focus:outline-none"
                    >
                      <option value="all">{t.allStatuses}</option>
                      {PROGRESS_STATUSES.map((st) => (
                        <option key={st} value={st}>
                          {st}
                        </option>
                      ))}
                    </select>

                    <select
                      value={industryFilter}
                      onChange={(e) => setIndustryFilter(e.target.value)}
                      className="rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-2 text-xs font-semibold text-[#0F172A] focus:border-[#16A34A] focus:outline-none max-w-[200px]"
                    >
                      <option value="all">{t.allIndustries}</option>
                      {industryFilterOptions.map((ind) => (
                        <option key={ind} value={ind}>
                          {ind}
                        </option>
                      ))}
                    </select>

                    <select
                      value={nationalityFilter}
                      onChange={(e) => setNationalityFilter(e.target.value)}
                      className="rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-2 text-xs font-semibold text-[#0F172A] focus:border-[#16A34A] focus:outline-none max-w-[180px]"
                    >
                      <option value="all">{t.allNationalities}</option>
                      {nationalityFilterOptions.map((nat) => (
                        <option key={nat} value={nat}>
                          {nat}
                        </option>
                      ))}
                    </select>

                    <select
                      value={proposedCompanyFilter}
                      onChange={(e) => setProposedCompanyFilter(e.target.value)}
                      className="rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-2 text-xs font-semibold text-[#0F172A] focus:border-[#16A34A] focus:outline-none max-w-[210px]"
                    >
                      <option value="all">{t.allProposedCompanies}</option>
                      {proposedCompanyFilterOptions.map((comp) => (
                        <option key={comp} value={comp}>
                          {comp}
                        </option>
                      ))}
                    </select>

                    {activeTab === 'proposed' && (
                      <select
                        value={proposedMonthFilter}
                        onChange={(e) => setProposedMonthFilter(e.target.value)}
                        className="rounded-md border border-emerald-200 bg-emerald-50 px-3 py-2 text-xs font-semibold text-emerald-700"
                      >
                        <option value="all">{t.allProposedMonths}</option>
                        {proposedMonthOptions.map((m) => (
                          <option key={m} value={m}>
                            {m}
                          </option>
                        ))}
                      </select>
                    )}

                    {(searchQuery ||
                      statusFilter !== 'all' ||
                      industryFilter !== 'all' ||
                      nationalityFilter !== 'all' ||
                      proposedCompanyFilter !== 'all' ||
                      proposedMonthFilter !== 'all' ||
                      careFilter !== 'all' ||
                      proposedStageFilter !== 'all') && (
                      <button
                        type="button"
                        onClick={() => {
                          setSearchQuery('');
                          setStatusFilter('all');
                          setIndustryFilter('all');
                          setNationalityFilter('all');
                          setProposedCompanyFilter('all');
                          setProposedMonthFilter('all');
                          setCareFilter('all');
                          setProposedStageFilter('all');
                          setDeclineReasonQuery('');
                        }}
                        className="rounded-md border border-[#E2E8F0] bg-white px-2.5 py-2 text-xs font-semibold text-[#64748B] hover:bg-[#F8FAFC] hover:text-[#0F172A] whitespace-nowrap shrink-0 cursor-pointer"
                      >
                        {t.resetFilters}
                      </button>
                    )}

                    <button
                      type="button"
                      onClick={handleBulkAutoFixCandidateNationalities}
                      className="inline-flex items-center gap-1 rounded-md border border-amber-300 bg-amber-50 px-2.5 py-2 text-xs font-semibold text-amber-800 hover:bg-amber-100 whitespace-nowrap shrink-0 cursor-pointer"
                    >
                      <Sparkles className="h-3.5 w-3.5" aria-hidden="true" />
                      {t.autoFixNatBtn}
                    </button>

                    <button
                      type="button"
                      onClick={handleDataCleanupAndReclassify}
                      className="inline-flex items-center gap-1 rounded-md border border-sky-300 bg-sky-50 px-2.5 py-2 text-xs font-semibold text-sky-800 hover:bg-sky-100 whitespace-nowrap shrink-0 cursor-pointer"
                    >
                      <RefreshCw className="h-3.5 w-3.5" aria-hidden="true" />
                      {t.dataCleanupBtn}
                    </button>

                    <button
                      type="button"
                      onClick={() => {
                        setFbChatMode('new');
                        setFbChatChannel(
                          activeTab === 'company'
                            ? 'company'
                            : activeTab === 'line'
                            ? 'line'
                            : activeTab === 'other'
                            ? 'other'
                            : activeTab === 'school'
                            ? 'school'
                            : 'personal'
                        );
                        setFbChatRawText('');
                        setFbChatModalOpen(true);
                      }}
                      className="rounded-md border border-[#E2E8F0] bg-white px-3 py-2 text-xs font-semibold text-[#0F172A] hover:border-[#CBD5E1] hover:bg-[#F8FAFC] whitespace-nowrap shrink-0 cursor-pointer"
                    >
                      {t.fbChatRegister}
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
                                : activeTab === 'line'
                                ? 'line'
                                : activeTab === 'other'
                                ? 'other'
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
                      className="rounded-md bg-[#16A34A] px-3.5 py-2 text-xs font-semibold text-white hover:bg-[#15803D] whitespace-nowrap shrink-0 cursor-pointer"
                    >
                      {t.newCandidate}
                    </button>
                  </div>

                  {/* Quick Care Filter Row + List/Card View Switcher */}
                  <div className="flex flex-wrap items-center justify-between gap-2 border-t border-[#E2E8F0] pt-2.5">
                    <div className="flex flex-wrap items-center gap-1.5">
                      {(
                        [
                          { id: 'all', label: t.careFilters.all },
                          { id: 'recent7', label: t.careFilters.recent7 },
                          { id: 'scheduled', label: t.careFilters.scheduled },
                          { id: 'recentMendan', label: t.careFilters.recentMendan },
                          { id: 'urgent', label: t.careFilters.urgent },
                          { id: 'declined', label: `${t.careFilters.declined} (${declinedPeopleCount})` },
                        ] as Array<{ id: CareFilterType; label: string }>
                      ).map((pill) => (
                        <button
                          key={pill.id}
                          type="button"
                          onClick={() => setCareFilter(pill.id)}
                          className={`rounded px-2.5 py-1 text-[11px] font-semibold transition-colors whitespace-nowrap shrink-0 tabular-nums cursor-pointer ${
                            careFilter === pill.id
                              ? pill.id === 'declined'
                                ? 'bg-rose-600 text-white'
                                : 'bg-[#16A34A] text-white'
                              : 'border border-[#E2E8F0] bg-[#F8FAFC] text-[#64748B] hover:text-[#0F172A]'
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
                          placeholder={t.declineReasonPlaceholder}
                          className="ml-1 w-60 rounded border border-rose-200 bg-rose-50/40 px-2.5 py-1 text-xs text-[#0F172A] placeholder:text-[#94A3B8] focus:outline-none"
                        />
                      )}
                    </div>

                    {/* Segmented List / Card View Switcher */}
                    <div className="inline-flex items-center rounded-md border border-[#E2E8F0] bg-[#F8FAFC] p-0.5 shrink-0">
                      <button
                        type="button"
                        onClick={() => setCandidateViewMode('list')}
                        className={`rounded px-2.5 py-1 text-[11px] font-semibold transition-colors whitespace-nowrap cursor-pointer ${
                          candidateViewMode === 'list'
                            ? 'bg-white text-[#0F172A] shadow-sm'
                            : 'text-[#64748B] hover:text-[#0F172A]'
                        }`}
                      >
                        {t.viewMode.list}
                      </button>
                      <button
                        type="button"
                        onClick={() => setCandidateViewMode('card')}
                        className={`rounded px-2.5 py-1 text-[11px] font-semibold transition-colors whitespace-nowrap cursor-pointer ${
                          candidateViewMode === 'card'
                            ? 'bg-white text-[#0F172A] shadow-sm'
                            : 'text-[#64748B] hover:text-[#0F172A]'
                        }`}
                      >
                        {t.viewMode.card}
                      </button>
                    </div>
                  </div>

                  {/* Proposed Tab 3-Subtab Pipeline Bar */}
                  {activeTab === 'proposed' && (
                    <div className="space-y-2 border-t border-[#E2E8F0] pt-2.5">
                      <div className="flex flex-wrap items-center gap-1.5">
                        <span className="mr-1 font-mono text-[10px] font-semibold uppercase tracking-wider text-[#64748B]">
                          {t.stageLabel}
                        </span>
                        {(
                          [
                            {
                              id: 'in_progress',
                              label: `${t.proposedGroups.inProgress} (${proposedStageCounts.in_progress})`,
                              activeClass: 'bg-sky-600 text-white',
                            },
                            {
                              id: 'naitei',
                              label: `${t.proposedGroups.naitei} (${proposedStageCounts.naitei})`,
                              activeClass: 'bg-[#16A34A] text-white',
                            },
                            {
                              id: 'jitai',
                              label: `${t.proposedGroups.jitai} (${proposedStageCounts.jitai})`,
                              activeClass: 'bg-rose-600 text-white',
                            },
                            {
                              id: 'all',
                              label: `${t.proposedGroups.all} (${proposedStageCounts.all})`,
                              activeClass: 'bg-[#0F172A] text-white',
                            },
                          ] as Array<{
                            id: ProposedGroupFilter;
                            label: string;
                            activeClass: string;
                          }>
                        ).map((stPill) => (
                          <button
                            key={stPill.id}
                            type="button"
                            onClick={() => {
                              setProposedStageFilter(stPill.id);
                              if (stPill.id !== 'jitai') setJitaiSubFilter('all');
                            }}
                            className={`rounded px-3 py-1.5 text-[11px] font-semibold transition-colors whitespace-nowrap shrink-0 tabular-nums cursor-pointer ${
                              proposedStageFilter === stPill.id
                                ? stPill.activeClass
                                : 'border border-[#E2E8F0] bg-[#F8FAFC] text-[#475569] hover:text-[#0F172A]'
                            }`}
                          >
                            {stPill.label}
                          </button>
                        ))}
                      </div>

                      {proposedStageFilter === 'jitai' && (
                        <div className="flex flex-wrap items-center gap-1.5 pl-1">
                          {(
                            [
                              {
                                id: 'all',
                                label: `${t.jitaiSub.all} (${proposedStageCounts.jitai})`,
                                tone: 'neutral',
                              },
                              {
                                id: '内定辞退',
                                label: `⚠️ ${t.jitaiSub.naiteiJitai} (${proposedStageCounts.jitaiBreakdown['内定辞退']})`,
                                tone: 'orange',
                              },
                              {
                                id: '選考中辞退',
                                label: `${t.jitaiSub.midJitai} (${proposedStageCounts.jitaiBreakdown['選考中辞退']})`,
                                tone: 'amber',
                              },
                              {
                                id: '不合格',
                                label: `${t.jitaiSub.fail} (${proposedStageCounts.jitaiBreakdown['不合格']})`,
                                tone: 'rose',
                              },
                              ...(proposedStageCounts.jitaiBreakdown['辞退・終了'] > 0
                                ? [
                                    {
                                      id: '辞退・終了' as JitaiSubFilter,
                                      label: `${t.jitaiSub.legacyDecline} (${proposedStageCounts.jitaiBreakdown['辞退・終了']})`,
                                      tone: 'neutral',
                                    },
                                  ]
                                : []),
                            ] as Array<{ id: JitaiSubFilter; label: string; tone: string }>
                          ).map((sub) => (
                            <button
                              key={sub.id}
                              type="button"
                              onClick={() => setJitaiSubFilter(sub.id)}
                              className={`rounded border px-2.5 py-1 text-[11px] font-semibold transition-colors whitespace-nowrap shrink-0 tabular-nums cursor-pointer ${
                                jitaiSubFilter === sub.id
                                  ? sub.tone === 'orange'
                                    ? 'border-orange-500 bg-orange-600 text-white'
                                    : sub.tone === 'rose'
                                    ? 'border-rose-600 bg-rose-600 text-white'
                                    : sub.tone === 'amber'
                                    ? 'border-amber-600 bg-amber-600 text-white'
                                    : 'border-[#0F172A] bg-[#0F172A] text-white'
                                  : sub.tone === 'orange'
                                  ? 'border-orange-300 bg-orange-50 text-orange-800 hover:bg-orange-100'
                                  : 'border-[#E2E8F0] bg-white text-[#475569] hover:border-[#CBD5E1]'
                              }`}
                            >
                              {sub.label}
                            </button>
                          ))}
                        </div>
                      )}
                    </div>
                  )}
                </div>

                {filteredCandidates.length === 0 ? (
                  <div className="rounded-lg border border-dashed border-[#CBD5E1] bg-white p-10 text-center text-sm font-semibold text-[#64748B]">
                    {t.noCandidateMatch}
                  </div>
                ) : candidateViewMode === 'list' ? (
                  <div className="overflow-x-auto rounded-lg border border-[#E2E8F0] bg-white shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
                    <table className="w-full min-w-[820px] border-collapse text-left text-xs">
                      <thead className="border-b border-[#E2E8F0] bg-[#F8FAFC] text-[10px] font-semibold uppercase tracking-wider text-[#64748B]">
                        <tr>
                          <th className="px-3.5 py-2.5">{t.listCols.nameNat}</th>
                          <th className="px-3 py-2.5">{t.listCols.companyIndustry}</th>
                          <th className="px-3 py-2.5">{t.listCols.status}</th>
                          <th className="px-3 py-2.5">{t.listCols.nextDateVisa}</th>
                          <th className="px-3.5 py-2.5">{t.listCols.latestNote}</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-[#E2E8F0]">
                        {filteredCandidates.map((person) => {
                          const isSchool = Boolean(person._isSchool || person.channel === 'school');
                          const displayName = getCandidateDisplayName(person, uiLang);
                          const proposedComp = getEffectiveProposedCompany(person);
                          const naiteiOrProposed = person.naiteiCompany || proposedComp;
                          const mendanInfo = extractNormalizedMendanInfo(person);
                          const fbMessengerUrl = getSafeFacebookMessengerUrl(person);
                          const st = person.progressStatus || '新規相談';
                          const isNaiteiJitai = st === '内定辞退';
                          const isNaitei = PROPOSED_NAITEI_STATUSES.includes(st);
                          const isJitai = PROPOSED_JITAI_STATUSES.includes(st);
                          const rawNote =
                            extractDeclineReason(person) ||
                            person.negotiationLog ||
                            person.memo ||
                            '';
                          const firstNoteLine = rawNote
                            .split('\n')
                            .map((s) => s.trim())
                            .find(Boolean) || '-';
                          const nextDateDisplay =
                            person.visaSubmissionDeadline
                              ? `🛂 ${person.visaSubmissionDeadline}`
                              : person.expectedJoinDate
                              ? `🎉 ${person.expectedJoinDate}`
                              : person.interviewDate ||
                                mendanInfo.dateLabel ||
                                person.lastContactDate ||
                                '-';

                          return (
                            <tr
                              key={person.id}
                              className={`transition-colors hover:bg-[#F8FAFC] ${
                                isNaiteiJitai ? 'bg-orange-50/30' : ''
                              }`}
                            >
                              <td className="px-3.5 py-2.5 align-top">
                                <div className="flex items-center gap-1.5">
                                  <button
                                    type="button"
                                    onClick={() => setDrawerCandidateId(person.id)}
                                    className="font-semibold text-[#0F172A] hover:text-[#16A34A] hover:underline text-left cursor-pointer"
                                  >
                                    {displayName}
                                  </button>
                                  {fbMessengerUrl ? (
                                    <a
                                      href={fbMessengerUrl}
                                      target="_blank"
                                      rel="noopener noreferrer"
                                      title={t.openMessengerBtn}
                                      className="inline-flex items-center rounded border border-sky-200 bg-sky-50 p-1 text-sky-700 hover:bg-sky-100 shrink-0"
                                    >
                                      <MessageCircle className="h-3 w-3" aria-hidden="true" />
                                    </a>
                                  ) : person.fbSenderId ? (
                                    <a
                                      href="https://business.facebook.com/latest/inbox/all"
                                      target="_blank"
                                      rel="noopener noreferrer"
                                      title={t.fbInboxBtn}
                                      className="inline-flex items-center rounded border border-sky-200 bg-sky-50 p-1 text-sky-700 hover:bg-sky-100 shrink-0"
                                    >
                                      <MessageCircle className="h-3 w-3" aria-hidden="true" />
                                    </a>
                                  ) : null}
                                </div>
                                <div className="mt-0.5 flex flex-wrap items-center gap-1.5 text-[11px] text-[#64748B]">
                                  <select
                                    aria-label={uiLang === 'jp' ? '国籍変更' : 'Sửa quốc tịch'}
                                    value={
                                      (CANDIDATE_NATIONALITIES as readonly string[]).includes(
                                        person.nationality || ''
                                      )
                                        ? person.nationality || ''
                                        : 'その他'
                                    }
                                    onChange={(e) => {
                                      const nextNat = e.target.value;
                                      const updater = (list: Candidate[]) =>
                                        list.map((c) =>
                                          c.id === person.id
                                            ? {
                                                ...c,
                                                nationality: nextNat,
                                                nationalityManuallyEdited: true,
                                              }
                                            : c
                                        );
                                      if (isSchool) setSchoolStudents(updater);
                                      else setCandidates(updater);
                                    }}
                                    className="rounded border border-transparent bg-transparent py-0.5 text-[11px] text-[#475569] hover:border-[#E2E8F0] hover:bg-white focus:border-[#16A34A] focus:bg-white focus:outline-none cursor-pointer"
                                  >
                                    {!(CANDIDATE_NATIONALITIES as readonly string[]).includes(
                                      person.nationality || ''
                                    ) &&
                                      person.nationality && (
                                        <option value={person.nationality}>
                                          {person.nationality}
                                        </option>
                                      )}
                                    {CANDIDATE_NATIONALITIES.map((nat) => (
                                      <option key={nat} value={nat}>
                                        {nat}
                                      </option>
                                    ))}
                                  </select>
                                  {person.jlpt && person.jlpt !== '要確認' && (
                                    <span className="rounded bg-emerald-50 px-1 py-0.2 font-mono text-[10px] font-semibold text-emerald-700">
                                      {person.jlpt}
                                    </span>
                                  )}
                                  {isStale7DayProposal(person) && (
                                    <span className="rounded border border-rose-300 bg-rose-50 px-1 py-0.5 font-mono text-[9px] font-bold text-rose-700">
                                      {t.stale7DaysBadge}
                                    </span>
                                  )}
                                </div>
                              </td>

                              <td className="px-3 py-2.5 align-top">
                                <div className="flex flex-wrap items-center gap-1.5">
                                  <span className="font-medium text-[#0F172A]">
                                    {naiteiOrProposed || '-'}
                                  </span>
                                  {isNaiteiJitai && (
                                    <span className="rounded border border-orange-300 bg-orange-100 px-1.5 py-0.5 text-[10px] font-bold text-orange-800 whitespace-nowrap">
                                      ⚠️ 内定辞退
                                    </span>
                                  )}
                                </div>
                                <div className="mt-0.5 text-[11px] text-[#64748B]">
                                  {person.desiredIndustry || t.cardLabels.unset}
                                  {person.desiredArea ? ` · ${person.desiredArea}` : ''}
                                </div>
                              </td>

                              <td className="px-3 py-2.5 align-top">
                                <select
                                  value={st}
                                  onChange={(e) =>
                                    handleStatusTransitionRequest(person, isSchool, e.target.value)
                                  }
                                  className={`rounded border px-2 py-1 text-[11px] font-semibold focus:border-[#16A34A] focus:outline-none cursor-pointer ${
                                    isNaiteiJitai
                                      ? 'border-orange-300 bg-orange-50 text-orange-800'
                                      : isNaitei
                                      ? 'border-emerald-200 bg-emerald-50 text-emerald-700'
                                      : isJitai
                                      ? 'border-rose-200 bg-rose-50 text-rose-700'
                                      : 'border-[#E2E8F0] bg-[#F8FAFC] text-[#0F172A]'
                                  }`}
                                >
                                  {PROGRESS_STATUSES.map((opt) => (
                                    <option key={opt} value={opt}>
                                      {opt}
                                    </option>
                                  ))}
                                </select>
                              </td>

                              <td className="px-3 py-2.5 align-top font-mono text-[11px] tabular-nums">
                                <div className="font-medium text-[#0F172A]">{nextDateDisplay}</div>
                                {person.visaExpiry && (
                                  <div className="mt-0.5 text-[10px] text-rose-700">
                                    Visa: {person.visaExpiry}
                                  </div>
                                )}
                              </td>

                              <td className="px-3.5 py-2.5 align-top">
                                <div className="flex items-start justify-between gap-2">
                                  <span
                                    title={rawNote}
                                    className="line-clamp-2 max-w-[240px] text-[11px] text-[#475569]"
                                  >
                                    {firstNoteLine}
                                  </span>
                                  <div className="flex items-center gap-1 shrink-0">
                                    {person.savedRirekisho?.draft && (
                                      <button
                                        type="button"
                                        onClick={() => openRirekishoModalForCandidate(person)}
                                        title={t.editSavedRirekishoBtn}
                                        className="rounded border border-emerald-300 bg-emerald-50 px-1.5 py-1 text-[10px] font-semibold text-emerald-800 hover:bg-emerald-100 whitespace-nowrap cursor-pointer"
                                      >
                                        ✅ CV
                                      </button>
                                    )}
                                    <button
                                      type="button"
                                      onClick={() => setDrawerCandidateId(person.id)}
                                      className="rounded border border-[#E2E8F0] bg-white px-2 py-1 text-[11px] font-semibold text-[#0F172A] hover:border-[#16A34A] hover:text-[#16A34A] whitespace-nowrap shrink-0 cursor-pointer"
                                    >
                                      {t.listCols.detailBtn}
                                    </button>
                                  </div>
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
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
                <div className="space-y-3 rounded-lg border border-[#E2E8F0] bg-white p-4 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <div className="flex flex-wrap items-center gap-2">
                      <button
                        type="button"
                        disabled={isImportingKyujin}
                        onClick={() => kyujinFileInputRef.current?.click()}
                        className="inline-flex items-center gap-1.5 rounded-md bg-[#16A34A] px-4 py-2 text-xs font-semibold text-white hover:bg-[#15803D] disabled:opacity-60 whitespace-nowrap shrink-0 cursor-pointer"
                      >
                        <Upload className="h-3.5 w-3.5" aria-hidden="true" />
                        {isImportingKyujin ? t.kyujin.analyzing : t.kyujin.importMultiBtn}
                      </button>

                      <button
                        type="button"
                        disabled={isImportingKyujin}
                        onClick={() => setKyujinRawModalOpen(true)}
                        className="inline-flex items-center gap-1.5 rounded-md border border-sky-200 bg-sky-50 px-3 py-2 text-xs font-semibold text-sky-700 hover:bg-sky-100 whitespace-nowrap shrink-0 cursor-pointer"
                      >
                        <ClipboardPaste className="h-3.5 w-3.5" aria-hidden="true" />
                        {t.kyujin.pasteTextBtn}
                      </button>

                      {selectedSlotIds.length > 0 && (
                        <button
                          type="button"
                          onClick={handleDeleteSelectedSlots}
                          className="rounded-md bg-rose-600 px-3 py-2 text-xs font-semibold text-white hover:bg-rose-700 whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          {t.kyujin.deleteSelected(selectedSlotIds.length)}
                        </button>
                      )}

                      {kyujins.length > 0 && (
                        <button
                          type="button"
                          onClick={handleDeleteAllKyujins}
                          className="rounded-md border border-rose-200 bg-rose-50 px-3 py-2 text-xs font-semibold text-rose-700 hover:bg-rose-100 whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          {t.kyujin.deleteAll}
                        </button>
                      )}
                    </div>

                    <input
                      type="text"
                      value={kyujinSearch}
                      onChange={(e) => setKyujinSearch(e.target.value)}
                      placeholder={t.kyujin.searchPlaceholder}
                      className="w-72 rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-1.5 text-xs text-[#0F172A] placeholder:text-[#94A3B8] focus:border-[#16A34A] focus:bg-white focus:outline-none"
                    />
                  </div>

                  {/* Screenshot Drag/Paste hint & progress bar */}
                  <div className="flex flex-wrap items-center justify-between gap-2 rounded-md border border-dashed border-[#CBD5E1] bg-[#F8FAFC] px-3 py-2 text-[11px] text-[#64748B]">
                    <span className="inline-flex items-center gap-1.5">
                      <ImageIcon className="h-3.5 w-3.5 text-sky-600" aria-hidden="true" />
                      {t.kyujin.pasteTip}
                    </span>
                    {kyujinImportProgress && (
                      <span
                        className={`font-mono font-semibold tabular-nums ${
                          /OCR|scan/i.test(kyujinImportProgress.label)
                            ? 'text-sky-700'
                            : 'text-emerald-700'
                        }`}
                      >
                        {t.kyujin.processing(
                          kyujinImportProgress.current,
                          kyujinImportProgress.total,
                          kyujinImportProgress.label
                        )}
                      </span>
                    )}
                  </div>

                  {kyujinImportProgress && (
                    <div className="h-1.5 w-full overflow-hidden rounded-full bg-slate-200">
                      <div
                        className="h-full bg-[#16A34A] transition-all"
                        style={{
                          width: `${Math.round(
                            (kyujinImportProgress.current / Math.max(1, kyujinImportProgress.total)) * 100
                          )}%`,
                        }}
                      />
                    </div>
                  )}

                  <div className="flex flex-wrap items-center gap-2 border-t border-[#E2E8F0] pt-3">
                    <div className="flex items-center gap-1 rounded-md border border-[#E2E8F0] bg-[#F8FAFC] p-1">
                      <button
                        type="button"
                        onClick={() => setKyujinStatusFilter('all')}
                        className={`rounded px-2.5 py-1 text-xs font-semibold tabular-nums whitespace-nowrap shrink-0 cursor-pointer ${
                          kyujinStatusFilter === 'all'
                            ? 'bg-[#16A34A] text-white'
                            : 'text-[#64748B] hover:text-[#0F172A]'
                        }`}
                      >
                        {t.kyujin.statusAll(kyujinStats.totalSlots)}
                      </button>
                      <button
                        type="button"
                        onClick={() => setKyujinStatusFilter('open')}
                        className={`rounded px-2.5 py-1 text-xs font-semibold tabular-nums whitespace-nowrap shrink-0 cursor-pointer ${
                          kyujinStatusFilter === 'open'
                            ? 'bg-[#16A34A] text-white'
                            : 'text-[#64748B] hover:text-[#0F172A]'
                        }`}
                      >
                        {t.kyujin.statusOpen(kyujinStats.openSlots)}
                      </button>
                      <button
                        type="button"
                        onClick={() => setKyujinStatusFilter('closed')}
                        className={`rounded px-2.5 py-1 text-xs font-semibold tabular-nums whitespace-nowrap shrink-0 cursor-pointer ${
                          kyujinStatusFilter === 'closed'
                            ? 'bg-rose-600 text-white'
                            : 'text-[#64748B] hover:text-[#0F172A]'
                        }`}
                      >
                        {t.kyujin.statusClosed(kyujinStats.closedSlots)}
                      </button>
                    </div>

                    <button
                      type="button"
                      onClick={() => setKyujinBonusOnly((prev) => !prev)}
                      className={`rounded-md border px-3 py-1.5 text-xs font-semibold transition-colors tabular-nums whitespace-nowrap shrink-0 cursor-pointer ${
                        kyujinBonusOnly
                          ? 'border-amber-300 bg-amber-50 text-amber-800'
                          : 'border-[#E2E8F0] bg-[#F8FAFC] text-[#64748B] hover:text-[#0F172A]'
                      }`}
                    >
                      {t.kyujin.hasBonus(kyujinStats.bonusSlots)}
                    </button>

                    <button
                      type="button"
                      onClick={() => setKyujinMissingSalaryOnly((prev) => !prev)}
                      className={`rounded-md border px-3 py-1.5 text-xs font-semibold transition-colors tabular-nums whitespace-nowrap shrink-0 cursor-pointer ${
                        kyujinMissingSalaryOnly
                          ? 'border-amber-400 bg-amber-100 text-amber-900'
                          : kyujinStats.missingSalarySlots > 0
                          ? 'border-amber-300 bg-amber-50 text-amber-800 hover:bg-amber-100'
                          : 'border-[#E2E8F0] bg-[#F8FAFC] text-[#64748B] hover:text-[#0F172A]'
                      }`}
                    >
                      {t.kyujin.missingSalary(kyujinStats.missingSalarySlots)}
                    </button>

                    <select
                      value={kyujinSectorFilter}
                      onChange={(e) => setKyujinSectorFilter(e.target.value)}
                      className="rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-1.5 text-xs font-semibold text-[#0F172A]"
                    >
                      <option value="all">{t.kyujin.allSectors}</option>
                      {SECTORS.map((sec) => (
                        <option key={sec} value={sec}>
                          {sec}
                        </option>
                      ))}
                    </select>

                    <select
                      value={kyujinIndustryFilter}
                      onChange={(e) => setKyujinIndustryFilter(e.target.value)}
                      className="rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-1.5 text-xs font-semibold text-[#0F172A]"
                    >
                      <option value="all">{t.kyujin.allIndustries}</option>
                      {INDUSTRIES.map((ind) => (
                        <option key={ind} value={ind}>
                          {ind}
                        </option>
                      ))}
                    </select>

                    <select
                      value={kyujinEmploymentTypeFilter}
                      onChange={(e) => setKyujinEmploymentTypeFilter(e.target.value)}
                      className="rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-1.5 text-xs font-semibold text-sky-700"
                    >
                      <option value="all">
                        {t.kyujin.allEmploymentTypes(kyujinStats.baitoSlots)}
                      </option>
                      {EMPLOYMENT_TYPES.map((emp) => (
                        <option key={emp} value={emp}>
                          {emp}
                        </option>
                      ))}
                    </select>

                    <select
                      value={kyujinFacilityFilter}
                      onChange={(e) => setKyujinFacilityFilter(e.target.value)}
                      className="rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-1.5 text-xs font-semibold text-[#0F172A]"
                    >
                      <option value="all">{t.kyujin.allFacilities}</option>
                      {KAIGO_FACILITY_TYPES.map((ft) => (
                        <option key={ft} value={ft}>
                          {ft}
                        </option>
                      ))}
                    </select>

                    <select
                      value={kyujinShiftTypeFilter}
                      onChange={(e) =>
                        setKyujinShiftTypeFilter(
                          e.target.value as 'all' | '日勤のみ' | 'シフト制' | '夜勤あり'
                        )
                      }
                      className="rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-1.5 text-xs font-semibold text-[#0F172A]"
                    >
                      <option value="all">{t.kyujin.allShiftTypes}</option>
                      <option value="日勤のみ">{t.kyujin.shiftDayOnly}</option>
                      <option value="シフト制">{t.kyujin.shiftRotating}</option>
                      <option value="夜勤あり">{t.kyujin.shiftHasNight}</option>
                    </select>
                  </div>
                </div>

                {filteredKyujinCompanies.length === 0 ? (
                  <div className="rounded-lg border border-dashed border-[#CBD5E1] bg-white p-10 text-center text-sm font-semibold text-[#64748B]">
                    {t.kyujin.emptyKyujin}
                  </div>
                ) : (
                  filteredKyujinCompanies.map((comp) => {
                    const allClosed = (comp.areaBreakdown || []).every(
                      (sl) => sl.recruitmentStatus === '募集終了'
                    );
                    return (
                      <div
                        key={comp.id}
                        className="rounded-lg border border-[#E2E8F0] bg-white p-4 shadow-[0_1px_3px_rgba(0,0,0,0.05)]"
                      >
                        <div className="mb-3 flex flex-wrap items-center justify-between gap-2 border-b border-[#E2E8F0] pb-2.5">
                          <div className="flex flex-wrap items-center gap-2.5">
                            <span className="text-base font-bold text-[#0F172A]">
                              {comp.companyName}
                            </span>
                            <span className="rounded bg-slate-100 border border-[#E2E8F0] px-2 py-0.5 font-mono text-[10px] font-medium text-[#475569]">
                              {comp.sector || detectSectorFromIndustryOrText(comp.industry)}
                            </span>
                            <span className="font-mono text-xs font-semibold text-emerald-700">
                              {comp.industry}
                            </span>
                            <span className="font-mono text-xs text-[#64748B] tabular-nums">
                              {t.kyujin.positionsCount(comp.areaBreakdown.length)}
                            </span>
                          </div>

                          <div className="flex flex-wrap items-center gap-1.5 shrink-0">
                            <button
                              type="button"
                              onClick={() =>
                                handleToggleCompanyRecruitmentStatus(
                                  comp.id,
                                  comp.companyName,
                                  allClosed ? '募集中' : '募集終了'
                                )
                              }
                              className={`rounded border px-2.5 py-1 text-xs font-semibold whitespace-nowrap shrink-0 cursor-pointer ${
                                allClosed
                                  ? 'border-emerald-200 bg-emerald-50 text-emerald-700 hover:bg-emerald-100'
                                  : 'border-amber-300 bg-amber-50 text-amber-800 hover:bg-amber-100'
                              }`}
                            >
                              {allClosed ? t.kyujin.bulkResumeBtn : t.kyujin.bulkPauseBtn}
                            </button>
                            <button
                              type="button"
                              onClick={() => handleDeleteSingleCompany(comp.id, comp.companyName)}
                              className="rounded border border-rose-200 bg-rose-50 px-2.5 py-1 text-xs font-semibold text-rose-700 hover:bg-rose-100 whitespace-nowrap shrink-0 cursor-pointer"
                            >
                              {t.kyujin.deleteCompany}
                            </button>
                          </div>
                        </div>

                        <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
                          {comp.areaBreakdown.map((sl) => {
                            const isChecked = selectedSlotIds.includes(sl.slotId);
                            const empType =
                              sl.employmentType ||
                              comp.employmentType ||
                              'Tokutei Ginou (特定技能)';
                            const isBaito = /baito|アルバイト|パート/i.test(empType);
                            const resolvedShiftType =
                              sl.shiftType ||
                              detectKyujinShiftType(sl.shiftTime, sl.jobDescription);
                            const isEditingJobCode =
                              editingKyujinField?.companyId === comp.id &&
                              editingKyujinField?.slotId === sl.slotId &&
                              editingKyujinField?.field === 'jobCode';
                            const isEditingArea =
                              editingKyujinField?.companyId === comp.id &&
                              editingKyujinField?.slotId === sl.slotId &&
                              editingKyujinField?.field === 'area';
                            const isEditingShift =
                              editingKyujinField?.companyId === comp.id &&
                              editingKyujinField?.slotId === sl.slotId &&
                              editingKyujinField?.field === 'shiftTime';

                            return (
                              <div
                                key={sl.slotId}
                                className={`rounded-md border p-3 transition-colors ${
                                  sl.recruitmentStatus === '募集終了'
                                    ? 'border-rose-200 bg-rose-50/50 opacity-85'
                                    : isBaito
                                    ? 'border-sky-200 bg-sky-50/30'
                                    : 'border-[#E2E8F0] bg-[#F8FAFC]'
                                }`}
                              >
                                <div className="mb-2 flex items-start justify-between gap-2">
                                  <div className="flex flex-wrap items-center gap-2">
                                    <input
                                      type="checkbox"
                                      checked={isChecked}
                                      onChange={() => toggleSelectSlot(sl.slotId)}
                                      className="h-4 w-4 rounded border-[#CBD5E1] bg-white accent-[#16A34A] cursor-pointer"
                                    />
                                    <span
                                      className={`rounded px-1.5 py-0.5 font-mono text-[10px] font-bold ${
                                        isBaito
                                          ? 'bg-sky-100 text-sky-800'
                                          : 'bg-emerald-100 text-emerald-800'
                                      }`}
                                    >
                                      {empType}
                                    </span>
                                    {isEditingJobCode ? (
                                      <input
                                        type="text"
                                        autoFocus
                                        value={editingKyujinField.value}
                                        onChange={(e) =>
                                          setEditingKyujinField((prev) =>
                                            prev ? { ...prev, value: e.target.value } : null
                                          )
                                        }
                                        onKeyDown={(e) => {
                                          if (e.key === 'Enter') {
                                            e.preventDefault();
                                            handleSaveInlineKyujinField(
                                              comp.id,
                                              sl.slotId,
                                              'jobCode',
                                              editingKyujinField.value
                                            );
                                          } else if (e.key === 'Escape') {
                                            setEditingKyujinField(null);
                                          }
                                        }}
                                        onBlur={() =>
                                          handleSaveInlineKyujinField(
                                            comp.id,
                                            sl.slotId,
                                            'jobCode',
                                            editingKyujinField.value
                                          )
                                        }
                                        className="w-28 rounded border border-[#16A34A] bg-white px-1.5 py-0.5 font-mono text-xs font-bold text-[#0F172A] outline-none"
                                      />
                                    ) : (
                                      <button
                                        type="button"
                                        onClick={() =>
                                          setEditingKyujinField({
                                            companyId: comp.id,
                                            slotId: sl.slotId,
                                            field: 'jobCode',
                                            value: sl.jobCode || '',
                                          })
                                        }
                                        className="group inline-flex items-center gap-0.5 rounded px-1 py-0.5 font-mono text-xs font-bold text-[#0F172A] hover:bg-slate-200/70 cursor-pointer"
                                      >
                                        <span>#{sl.jobCode || 'Slot'}</span>
                                        <span className="text-[10px] opacity-0 group-hover:opacity-75">
                                          ✎
                                        </span>
                                      </button>
                                    )}
                                    {isEditingArea ? (
                                      <input
                                        type="text"
                                        autoFocus
                                        value={editingKyujinField.value}
                                        onChange={(e) =>
                                          setEditingKyujinField((prev) =>
                                            prev ? { ...prev, value: e.target.value } : null
                                          )
                                        }
                                        onKeyDown={(e) => {
                                          if (e.key === 'Enter') {
                                            e.preventDefault();
                                            handleSaveInlineKyujinField(
                                              comp.id,
                                              sl.slotId,
                                              'area',
                                              editingKyujinField.value
                                            );
                                          } else if (e.key === 'Escape') {
                                            setEditingKyujinField(null);
                                          }
                                        }}
                                        onBlur={() =>
                                          handleSaveInlineKyujinField(
                                            comp.id,
                                            sl.slotId,
                                            'area',
                                            editingKyujinField.value
                                          )
                                        }
                                        className="w-32 rounded border border-[#16A34A] bg-white px-1.5 py-0.5 text-xs font-semibold text-[#334155] outline-none"
                                      />
                                    ) : (
                                      <button
                                        type="button"
                                        onClick={() =>
                                          setEditingKyujinField({
                                            companyId: comp.id,
                                            slotId: sl.slotId,
                                            field: 'area',
                                            value: sl.area || '',
                                          })
                                        }
                                        className="group inline-flex items-center gap-0.5 rounded px-1 py-0.5 text-xs font-semibold text-[#334155] hover:bg-slate-200/70 cursor-pointer"
                                      >
                                        <span>
                                          {formatPublicArea(sl.area) || sl.area || 'Nhật Bản'}
                                        </span>
                                        <span className="text-[10px] opacity-0 group-hover:opacity-75">
                                          ✎
                                        </span>
                                      </button>
                                    )}
                                    {sl.facilityType && (
                                      <span className="text-[11px] font-semibold text-purple-700">
                                        · {sl.facilityType}
                                      </span>
                                    )}
                                  </div>

                                  <button
                                    type="button"
                                    onClick={() => setPreviewModal({ company: comp, slot: sl })}
                                    className="rounded bg-[#16A34A] px-2.5 py-1 text-xs font-semibold text-white hover:bg-[#15803D] whitespace-nowrap shrink-0 cursor-pointer"
                                  >
                                    {t.kyujin.exportAnon}
                                  </button>
                                </div>

                                {/* Shift Time / Holidays / Overtime / Annual Salary Metadata Strip */}
                                <div className="mb-2 flex flex-wrap items-center gap-x-2.5 gap-y-1 rounded border border-[#E2E8F0] bg-white px-2.5 py-1.5 text-[11px]">
                                  {isEditingShift ? (
                                    <input
                                      type="text"
                                      autoFocus
                                      value={editingKyujinField.value}
                                      placeholder="例: ①08:00〜17:00 ②09:00〜18:00"
                                      onChange={(e) =>
                                        setEditingKyujinField((prev) =>
                                          prev ? { ...prev, value: e.target.value } : null
                                        )
                                      }
                                      onKeyDown={(e) => {
                                        if (e.key === 'Enter') {
                                          e.preventDefault();
                                          handleSaveInlineKyujinField(
                                            comp.id,
                                            sl.slotId,
                                            'shiftTime',
                                            editingKyujinField.value
                                          );
                                        } else if (e.key === 'Escape') {
                                          setEditingKyujinField(null);
                                        }
                                      }}
                                      onBlur={() =>
                                        handleSaveInlineKyujinField(
                                          comp.id,
                                          sl.slotId,
                                          'shiftTime',
                                          editingKyujinField.value
                                        )
                                      }
                                      className="min-w-[220px] flex-1 rounded border border-[#16A34A] bg-white px-2 py-0.5 font-mono text-[11px] text-[#0F172A] outline-none"
                                    />
                                  ) : (
                                    <button
                                      type="button"
                                      onClick={() =>
                                        setEditingKyujinField({
                                          companyId: comp.id,
                                          slotId: sl.slotId,
                                          field: 'shiftTime',
                                          value: sl.shiftTime || '',
                                        })
                                      }
                                      className={`group inline-flex items-center gap-1 rounded px-1 py-0.5 text-left font-medium cursor-pointer ${
                                        sl.shiftTime
                                          ? 'text-sky-800 hover:bg-sky-50'
                                          : 'text-amber-800 bg-amber-50 hover:bg-amber-100'
                                      }`}
                                    >
                                      <span>
                                        🕒 {sl.shiftTime || (uiLang === 'jp' ? '⚠️ 勤務時間: 要確認' : '⚠️ Giờ làm: Cần xác nhận')}
                                      </span>
                                      <span className="text-[10px] opacity-0 group-hover:opacity-80">
                                        ✎
                                      </span>
                                    </button>
                                  )}
                                  {resolvedShiftType && (
                                    <span className="rounded bg-slate-100 px-1.5 py-0.5 text-[10px] font-semibold text-[#475569]">
                                      {resolvedShiftType}
                                    </span>
                                  )}
                                  {sl.holidays && (
                                    <span className="text-[#475569]">
                                      📅 {uiLang === 'jp' ? '休日' : 'Nghỉ'}: {sl.holidays}
                                    </span>
                                  )}
                                  {sl.overtimeHours && (
                                    <span className="text-[#475569]">
                                      ⏱ {uiLang === 'jp' ? '残業' : 'OT'}: {sl.overtimeHours}
                                    </span>
                                  )}
                                  {sl.annualSalary && (
                                    <span className="font-mono font-semibold text-emerald-700">
                                      💴 {uiLang === 'jp' ? '想定年収' : 'Thu nhập năm'}: {sl.annualSalary}
                                    </span>
                                  )}
                                  {sl.recruitCount && (
                                    <span className="text-[#475569]">
                                      👥 {sl.recruitCount}
                                    </span>
                                  )}
                                  {sl.interviewProcess && (
                                    <span className="text-[#64748B]">
                                      🎯 {sl.interviewProcess}
                                    </span>
                                  )}
                                  {sl.specialBenefits && (
                                    <span className="text-emerald-700">
                                      ✨ {sl.specialBenefits}
                                    </span>
                                  )}
                                </div>

                              {/* Compensation & Experience Strip */}
                              {(() => {
                                const salaryMeta = getKyujinSlotSalaryDisplayMeta(sl, empType);
                                const isEditingThisSalary =
                                  editingKyujinSalary?.companyId === comp.id &&
                                  editingKyujinSalary?.slotId === sl.slotId;

                                return (
                                  <div className="mb-2 flex flex-wrap items-center gap-2.5 text-xs">
                                    {isEditingThisSalary ? (
                                      <div className="flex items-center gap-1.5">
                                        <span className="font-mono text-[11px] font-semibold text-[#64748B]">
                                          {salaryMeta.labelPrefix}:
                                        </span>
                                        <input
                                          type="text"
                                          autoFocus
                                          value={editingKyujinSalary.value}
                                          placeholder="VD: 230,000円 hoặc 1,250円"
                                          onChange={(e) =>
                                            setEditingKyujinSalary((prev) =>
                                              prev ? { ...prev, value: e.target.value } : null
                                            )
                                          }
                                          onKeyDown={(e) => {
                                            if (e.key === 'Enter') {
                                              e.preventDefault();
                                              handleSaveInlineKyujinSalary(
                                                comp.id,
                                                sl.slotId,
                                                editingKyujinSalary.value
                                              );
                                            } else if (e.key === 'Escape') {
                                              e.preventDefault();
                                              salaryEditCommitLockRef.current = `${comp.id}::${sl.slotId}`;
                                              setEditingKyujinSalary(null);
                                            }
                                          }}
                                          onBlur={() =>
                                            handleSaveInlineKyujinSalary(
                                              comp.id,
                                              sl.slotId,
                                              editingKyujinSalary.value
                                            )
                                          }
                                          className="w-48 rounded border border-[#16A34A] bg-white px-2 py-0.5 font-mono text-xs font-bold text-[#0F172A] outline-none ring-1 ring-emerald-200"
                                        />
                                        <span className="text-[10px] text-[#64748B]">
                                          {t.kyujin.enterToSave}
                                        </span>
                                      </div>
                                    ) : (
                                      <button
                                        type="button"
                                        onClick={() =>
                                          handleStartInlineKyujinSalaryEdit(
                                            comp.id,
                                            sl.slotId,
                                            salaryMeta.editInitialValue
                                          )
                                        }
                                        className={`group inline-flex cursor-pointer items-center gap-1 rounded px-1.5 py-0.5 font-mono font-bold tabular-nums transition-colors ${
                                          salaryMeta.isMissing
                                            ? 'border border-amber-300 bg-amber-50 text-amber-800 hover:bg-amber-100'
                                            : salaryMeta.isHourly
                                            ? 'text-sky-700 hover:bg-sky-50'
                                            : 'text-emerald-700 hover:bg-emerald-50'
                                        }`}
                                      >
                                        <span>
                                          {salaryMeta.labelPrefix}: {salaryMeta.displayValue}
                                        </span>
                                        {sl.salaryManuallyEdited && (
                                          <span className="text-[10px] opacity-80">🔒</span>
                                        )}
                                        <span className="text-[10px] opacity-0 transition-opacity group-hover:opacity-80">
                                          ✎
                                        </span>
                                      </button>
                                    )}

                                    {isBaito ? (
                                      <>
                                        {sl.commuteAllowance && (
                                          <span className="text-[#64748B]">
                                            {uiLang === 'jp' ? '交通費' : 'Đi lại'}: {sl.commuteAllowance}
                                          </span>
                                        )}
                                        {sl.shiftsPerWeek && (
                                          <span className="text-[#64748B]">
                                            {uiLang === 'jp' ? 'シフト' : 'Ca'}: {sl.shiftsPerWeek}
                                          </span>
                                        )}
                                      </>
                                    ) : (
                                      <>
                                        {sl.estimatedNetSalary && (
                                          <span className="font-mono text-[11px] font-medium text-sky-700">
                                            {uiLang === 'jp' ? '手取り' : 'Thực nhận'}: {sl.estimatedNetSalary}
                                          </span>
                                        )}
                                        {sl.nightShiftAllowance && (
                                          <span className="text-[11px] text-[#64748B]">
                                            {uiLang === 'jp' ? '夜勤手当' : 'Ca đêm'}: {sl.nightShiftAllowance}
                                          </span>
                                        )}
                                        {sl.bonus === '賞与あり' && (
                                          <span className="font-semibold text-amber-700">
                                            {uiLang === 'jp' ? '賞与' : 'Thưởng'}
                                            {sl.bonusDetail ? ` (${sl.bonusDetail})` : ''}
                                          </span>
                                        )}
                                      </>
                                    )}
                                  </div>
                                );
                              })()}

                              {sl.experienceRequirement && (
                                <div className="mb-2 flex flex-wrap items-center gap-1.5 text-[11px]">
                                  <span
                                    className={`rounded px-1.5 py-0.5 font-semibold ${
                                      /未経験/.test(sl.experienceRequirement)
                                        ? 'bg-emerald-50 text-emerald-700 border border-emerald-200'
                                        : 'bg-amber-50 text-amber-800 border border-amber-300'
                                    }`}
                                  >
                                    {sl.experienceRequirement}
                                  </span>
                                  {sl.experienceDetail && (
                                    <span className="text-[#64748B]">{sl.experienceDetail}</span>
                                  )}
                                </div>
                              )}

                              <div className="mb-2 flex flex-wrap items-center gap-1.5 text-[10px]">
                                {(sl.requiredConditions || []).map((reqTag) => (
                                  <span
                                    key={reqTag}
                                    className="rounded border border-sky-200 bg-sky-50 px-1.5 py-0.5 font-semibold text-sky-700"
                                  >
                                    ✓ {reqTag}
                                  </span>
                                ))}
                                {(sl.ngConditions || []).map((ngTag) => (
                                  <span
                                    key={ngTag}
                                    className="rounded border border-rose-200 bg-rose-50 px-1.5 py-0.5 font-semibold text-rose-700"
                                  >
                                    🚫 {ngTag}
                                  </span>
                                ))}
                                <button
                                  type="button"
                                  onClick={() =>
                                    setKyujinConditionModal({
                                      companyId: comp.id,
                                      slotId: sl.slotId,
                                      companyName: comp.companyName,
                                      excludedNationalitiesText: (
                                        sl.excludedNationalities || []
                                      ).join(', '),
                                      excludedGenders: sl.excludedGenders || [],
                                      requiresInJapan: Boolean(sl.requiresInJapan),
                                      requiresTokutei: Boolean(sl.requiresTokutei),
                                      extraRequiredText: (sl.requiredConditions || []).join(', '),
                                      extraNgText: (sl.ngConditions || []).join(', '),
                                    })
                                  }
                                  className="rounded border border-[#E2E8F0] bg-white px-1.5 py-0.5 font-semibold text-[#64748B] hover:border-[#16A34A] hover:text-[#16A34A] cursor-pointer"
                                >
                                  ✎ {t.kyujinConds.editNgBtn}
                                </button>
                              </div>

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
                                      ? 'border-rose-200 bg-rose-50 text-rose-700'
                                      : 'border-emerald-200 bg-emerald-50 text-emerald-700'
                                  }`}
                                >
                                  <option value="募集中">{t.kyujin.recruitingOpt}</option>
                                  <option value="募集終了">{t.kyujin.closedOpt}</option>
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
                                  className="rounded border border-[#E2E8F0] bg-white px-2 py-1 text-[11px] font-semibold text-sky-700"
                                >
                                  {EMPLOYMENT_TYPES.map((emp) => (
                                    <option key={emp} value={emp}>
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
                                  className="rounded border border-[#E2E8F0] bg-white px-2 py-1 text-[11px] font-semibold text-[#0F172A]"
                                >
                                  {INDUSTRIES.map((ind) => (
                                    <option key={ind} value={ind}>
                                      {ind}
                                    </option>
                                  ))}
                                  <option value="特定技能">特定技能</option>
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
                                  className="rounded border border-[#E2E8F0] bg-white px-2 py-1 text-[11px] font-semibold text-[#0F172A]"
                                >
                                  {JLPT_LEVELS.map((lvl) => (
                                    <option key={lvl} value={lvl}>
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
                                  className="rounded border border-[#E2E8F0] bg-white px-2 py-1 text-[11px] font-semibold text-[#0F172A]"
                                >
                                  <option value="🏠 自社寮あり">
                                    {uiLang === 'jp' ? '🏠 自社寮あり' : '🏠 KTX công ty'}
                                  </option>
                                  <option value="🏠 借上げ宿舎あり">
                                    {uiLang === 'jp' ? '🏠 借上げ宿舎あり' : '🏠 Nhà thuê hỗ trợ'}
                                  </option>
                                  <option value="🏠 寮・社宅あり">
                                    {uiLang === 'jp' ? '🏠 寮・社宅あり' : '🏠 Có KTX/Nhà ở'}
                                  </option>
                                  <option value="❌ 提供無">
                                    {uiLang === 'jp' ? '❌ 提供無' : '❌ Tự túc nhà'}
                                  </option>
                                  <option value="⚠️ 要確認">
                                    {uiLang === 'jp' ? '⚠️ 住居要確認' : '⚠️ Nhà: Cần xác nhận'}
                                  </option>
                                </select>
                              </div>

                              {sl.jobDescription && (
                                <p className="text-xs leading-relaxed text-[#475569]">
                                  {sl.jobDescription}
                                </p>
                              )}
                            </div>
                          );
                        })}
                      </div>
                    </div>
                    );
                  })
                )}
              </div>
            )}

            {/* RESIGNATION MANAGEMENT TAB */}
            {activeTab === 'resignation' && (
              <ResignationTab
                records={resignations}
                lang={uiLang}
                onSaveRecord={handleSaveResignation}
                onDeleteRecord={handleDeleteResignation}
                onStatusChange={handleResignationStatusChange}
                onBulkImport={handleBulkImportResignations}
                onBulkUpdateRecords={handleBulkUpdateResignations}
                onSyncCalendarSingle={handleSyncResignationCalendarSingle}
                onSyncCalendarBulk={handleBulkSyncResignationsCalendar}
                onConvertToCandidate={handleConvertResignationToCandidate}
                syncingCalendarIds={syncingResignationCalIds}
                isBulkCalendarSyncing={isBulkSyncingResignationsCal}
                onNotify={(msg) => showToast(msg)}
                hasGasUrl={Boolean(gasWebAppUrl.trim())}
                notificationPermission={notificationPermission}
                onRequestBrowserNotifications={() => void handleRequestBrowserNotifications()}
              />
            )}

            {/* PT PARTNERSHIP TAB */}
            {activeTab === 'pt' && (
              <div className="space-y-4">
                <div className="flex flex-wrap items-center justify-between gap-2 rounded-lg border border-[#E2E8F0] bg-white p-4 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
                  <div>
                    <h2 className="text-base font-bold text-[#0F172A]">{t.pt.title}</h2>
                    <p className="text-xs text-[#64748B]">{t.pt.subtitle}</p>
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
                    className="rounded-md bg-[#16A34A] px-3.5 py-2 text-xs font-semibold text-white hover:bg-[#15803D] whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    {t.pt.addBtn}
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
                        className="rounded-lg border border-[#E2E8F0] bg-white p-4 shadow-[0_1px_3px_rgba(0,0,0,0.05)]"
                      >
                        <div className="mb-2 flex items-start justify-between gap-2">
                          <div>
                            <div className="flex items-center gap-2">
                              <span className="text-base font-bold text-[#0F172A]">
                                {pt.partnerName}
                              </span>
                              <span className="text-[#CBD5E1]" aria-hidden="true">
                                ·
                              </span>
                              <span className="font-mono text-xs font-semibold text-emerald-700">
                                {pt.status}
                              </span>
                            </div>
                            <p className="mt-1 text-xs text-[#64748B]">
                              {t.pt.contactLabel} {pt.contactPerson || t.cardLabels.unset} / TEL:{' '}
                              {pt.phone || '-'} / Email: {pt.email || '-'}
                            </p>
                          </div>
                          <div className="flex gap-1.5">
                            <button
                              type="button"
                              onClick={() => setPtModal({ isNew: false, data: { ...pt } })}
                              className="rounded border border-[#E2E8F0] bg-white px-2.5 py-1 text-xs font-semibold text-[#475569] hover:bg-[#F8FAFC] hover:text-[#0F172A] whitespace-nowrap shrink-0 cursor-pointer"
                            >
                              {t.editBtn}
                            </button>
                            <button
                              type="button"
                              onClick={() =>
                                setConfirmDialog({
                                  title: uiLang === 'jp' ? 'PT提携先の削除' : 'Xóa đối tác PT',
                                  message:
                                    uiLang === 'jp'
                                      ? `「${pt.partnerName}」を削除しますか？`
                                      : `Bạn có chắc muốn xóa đối tác "${pt.partnerName}"?`,
                                  onConfirm: () => {
                                    setPtPartners((prev) => prev.filter((x) => x.id !== pt.id));
                                    setConfirmDialog(null);
                                    showToast(
                                      uiLang === 'jp'
                                        ? '🗑️ PT提携先を削除しました'
                                        : '🗑️ Đã xóa đối tác PT'
                                    );
                                  },
                                })
                              }
                              className="rounded border border-rose-200 bg-rose-50 px-2.5 py-1 text-xs font-semibold text-rose-700 hover:bg-rose-100 whitespace-nowrap shrink-0 cursor-pointer"
                            >
                              {t.deleteBtn}
                            </button>
                          </div>
                        </div>

                        <div className="space-y-1.5 border-t border-[#E2E8F0] pt-2.5 text-xs text-[#64748B]">
                          <div>
                            <span className="font-semibold text-[#0F172A]">{t.pt.specialtyLabel} </span>
                            {pt.specialty || (uiLang === 'jp' ? '全分野対応' : 'Tất cả ngành')}
                          </div>
                          <div>
                            <span className="font-semibold text-[#0F172A]">{t.pt.termsLabel} </span>
                            {pt.commissionTerms || (uiLang === 'jp' ? '別途協議' : 'Thỏa thuận riêng')}
                          </div>
                          {pt.memo && <div className="text-[#64748B]">{pt.memo}</div>}
                        </div>

                        <div className="mt-2.5 font-mono text-xs font-semibold text-emerald-700 tabular-nums">
                          {t.pt.linkedCandidates(relatedPeople.length)}
                          {relatedPeople.length > 0 && (
                            <span className="ml-1 font-sans font-normal text-[#64748B]">
                              （{relatedPeople.map((r) => getCandidateDisplayName(r, uiLang)).join('、')}）
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
                    label: t.tabs.personal,
                    count: personalCandidates.length,
                    pct: Math.round((personalCandidates.length / totalPeopleCount) * 100),
                    barClass: 'bg-[#16A34A]',
                  },
                  {
                    label: t.tabs.company,
                    count: companyCandidates.length,
                    pct: Math.round((companyCandidates.length / totalPeopleCount) * 100),
                    barClass: 'bg-sky-500',
                  },
                  {
                    label: t.report.lineChannel,
                    count: lineCandidates.length,
                    pct: Math.round((lineCandidates.length / totalPeopleCount) * 100),
                    barClass: 'bg-emerald-400',
                  },
                  {
                    label: t.report.otherChannel,
                    count: otherCandidates.length,
                    pct: Math.round((otherCandidates.length / totalPeopleCount) * 100),
                    barClass: 'bg-amber-500',
                  },
                  {
                    label: t.tabs.school,
                    count: schoolCandidates.length,
                    pct: Math.round((schoolCandidates.length / totalPeopleCount) * 100),
                    barClass: 'bg-purple-500',
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
                  const ind = String(p.desiredIndustry || t.cardLabels.unset).trim() || t.cardLabels.unset;
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
                  const nat = String(p.nationality || t.cardLabels.unset).trim() || t.cardLabels.unset;
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
                      <div className="rounded-lg border border-[#E2E8F0] bg-white p-4 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
                        <div className="font-mono text-[10px] font-semibold uppercase tracking-wider text-[#64748B]">
                          {t.report.totalTalent}
                        </div>
                        <div className="mt-1 font-mono text-2xl font-bold text-[#0F172A] tabular-nums">
                          {allPeopleCombined.length} {t.report.peopleUnit}
                        </div>
                        <div className="mt-1 text-[11px] text-[#64748B] tabular-nums">
                          {uiLang === 'jp'
                            ? `SNS・紹介等 ${nonSchoolCandidates.length}名 · 留学生 ${schoolCandidates.length}名`
                            : `FB/LINE/Khác ${nonSchoolCandidates.length} · DHS ${schoolCandidates.length}`}
                        </div>
                      </div>
                      <div className="rounded-lg border border-[#E2E8F0] bg-white p-4 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
                        <div className="font-mono text-[10px] font-semibold uppercase tracking-wider text-[#64748B]">
                          {t.report.proposedCount}
                        </div>
                        <div className="mt-1 font-mono text-2xl font-bold text-emerald-700 tabular-nums">
                          {allProposedPeople.length} {t.report.peopleUnit}
                        </div>
                        <div className="mt-1 text-[11px] text-[#64748B] tabular-nums">
                          {uiLang === 'jp'
                            ? `面談予定 ${mendanPeopleSorted.length}名`
                            : `Mendan ${mendanPeopleSorted.length} ${t.report.peopleUnit}`}
                        </div>
                      </div>
                      <div className="rounded-lg border border-[#E2E8F0] bg-white p-4 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
                        <div className="font-mono text-[10px] font-semibold uppercase tracking-wider text-[#64748B]">
                          {t.report.openSlots}
                        </div>
                        <div className="mt-1 font-mono text-2xl font-bold text-amber-700 tabular-nums">
                          {kyujinStats.openSlots} {t.report.slotsUnit}
                        </div>
                        <div className="mt-1 text-[11px] text-[#64748B] tabular-nums">
                          {kyujinStats.totalCompanies} {t.report.companiesUnit} · {t.report.bonusLabel} {kyujinStats.bonusSlots}
                        </div>
                      </div>
                      <div className="rounded-lg border border-[#E2E8F0] bg-white p-4 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
                        <div className="font-mono text-[10px] font-semibold uppercase tracking-wider text-[#64748B]">
                          {t.report.ptCount}
                        </div>
                        <div className="mt-1 font-mono text-2xl font-bold text-[#0F172A] tabular-nums">
                          {ptPartners.length} {t.report.companiesUnit}
                        </div>
                        <div className="mt-1 text-[11px] text-[#64748B]">
                          {uiLang === 'jp' ? '継続連携・案件共有中' : 'Đang hợp tác & chia sẻ đơn'}
                        </div>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
                      <div className="rounded-lg border border-[#E2E8F0] bg-white p-4 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
                        <div className="mb-3 flex items-center justify-between">
                          <h3 className="text-sm font-bold text-[#0F172A]">
                            {t.report.channelDist(allPeopleCombined.length)}
                          </h3>
                          <span className="font-mono text-[10px] font-semibold uppercase tracking-wider text-emerald-700">
                            Summary
                          </span>
                        </div>

                        <div className="mb-3 flex h-3 w-full overflow-hidden rounded-full bg-slate-100">
                          {channelBreakdown.map((ch) =>
                            ch.count > 0 ? (
                              <div
                                key={ch.label}
                                style={{ width: `${Math.max(4, ch.pct)}%` }}
                                className={`${ch.barClass} transition-all`}
                                title={`${ch.label}: ${ch.count} (${ch.pct}%)`}
                              />
                            ) : null
                          )}
                        </div>

                        <div className="space-y-2.5">
                          {channelBreakdown.map((ch) => (
                            <div key={ch.label} className="space-y-1">
                              <div className="flex items-center justify-between text-xs">
                                <span className="font-semibold text-[#0F172A]">{ch.label}</span>
                                <span className="font-mono font-bold text-[#0F172A] tabular-nums">
                                  {ch.count} {t.report.peopleUnit}{' '}
                                  <span className="text-[#64748B]">({ch.pct}%)</span>
                                </span>
                              </div>
                              <div className="h-2 w-full overflow-hidden rounded-full bg-slate-100">
                                <div
                                  style={{ width: `${ch.pct}%` }}
                                  className={`h-full rounded-full ${ch.barClass}`}
                                />
                              </div>
                            </div>
                          ))}
                        </div>

                        <div className="mt-4 border-t border-[#E2E8F0] pt-3">
                          <div className="mb-2 font-mono text-[10px] font-semibold uppercase tracking-wider text-[#64748B]">
                            {t.report.nationalityBreakdown}
                          </div>
                          <div className="flex flex-wrap gap-2 text-xs">
                            {nationalityChartRows.map((item) => (
                              <div
                                key={item.nat}
                                className="flex items-center gap-1.5 rounded border border-[#E2E8F0] bg-[#F8FAFC] px-2.5 py-1"
                              >
                                <span className="font-semibold text-[#0F172A]">{item.nat}</span>
                                <span className="font-mono font-bold text-emerald-700 tabular-nums">
                                  {item.count} ({item.pct}%)
                                </span>
                              </div>
                            ))}
                          </div>
                        </div>
                      </div>

                      <div className="rounded-lg border border-[#E2E8F0] bg-white p-4 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
                        <div className="mb-3 flex items-center justify-between">
                          <h3 className="text-sm font-bold text-[#0F172A]">
                            {t.report.statusPipeline}
                          </h3>
                        </div>
                        <div className="space-y-2">
                          {statusChartRows.map((row) => (
                            <div key={row.status} className="space-y-1">
                              <div className="flex items-center justify-between text-xs">
                                <span className="font-semibold text-[#0F172A]">{row.status}</span>
                                <span className="font-mono font-bold text-emerald-700 tabular-nums">
                                  {row.count} {t.report.peopleUnit} ({row.pct}%)
                                </span>
                              </div>
                              <div className="h-2 w-full overflow-hidden rounded-full bg-slate-100">
                                <div
                                  style={{ width: `${Math.max(3, row.pct)}%` }}
                                  className={`h-full rounded-full ${
                                    row.status.includes('内定') || row.status.includes('入社')
                                      ? 'bg-[#16A34A]'
                                      : row.status.includes('面接') || row.status.includes('提案')
                                      ? 'bg-sky-500'
                                      : row.status.includes('面談')
                                      ? 'bg-teal-500'
                                      : row.status.includes('保留') || row.status.includes('辞退')
                                      ? 'bg-rose-500'
                                      : 'bg-slate-400'
                                  }`}
                                />
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
                      <div className="space-y-4 rounded-lg border border-[#E2E8F0] bg-white p-4 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
                        <div>
                          <h3 className="mb-2.5 text-sm font-bold text-[#0F172A]">
                            {t.report.industryDist}
                          </h3>
                          <div className="space-y-2">
                            {industryChartRows.map((row) => (
                              <div key={row.industry} className="space-y-1">
                                <div className="flex items-center justify-between text-xs">
                                  <span className="font-semibold text-[#0F172A]">{row.industry}</span>
                                  <span className="font-mono font-bold text-emerald-700 tabular-nums">
                                    {row.count} {t.report.peopleUnit} ({row.pct}%)
                                  </span>
                                </div>
                                <div className="h-2 w-full overflow-hidden rounded-full bg-slate-100">
                                  <div
                                    style={{ width: `${Math.max(3, row.pct)}%` }}
                                    className="h-full rounded-full bg-[#16A34A]"
                                  />
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>

                        <div className="border-t border-[#E2E8F0] pt-3">
                          <h3 className="mb-2.5 text-sm font-bold text-[#0F172A]">
                            {t.report.jlptDist}
                          </h3>
                          <div className="grid grid-cols-2 gap-2 sm:grid-cols-4">
                            {jlptChartRows.map((row) => (
                              <div
                                key={row.level}
                                className="rounded border border-[#E2E8F0] bg-[#F8FAFC] p-2.5 text-center"
                              >
                                <div className="font-mono text-xs font-bold text-emerald-700">
                                  {row.level}
                                </div>
                                <div className="mt-0.5 font-mono text-lg font-bold text-[#0F172A] tabular-nums">
                                  {row.count}
                                </div>
                                <div className="font-mono text-[10px] font-semibold text-[#64748B] tabular-nums">
                                  {row.pct}%
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>
                      </div>

                      <div className="space-y-4 rounded-lg border border-[#E2E8F0] bg-white p-4 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
                        <div className="flex items-center justify-between">
                          <h3 className="text-sm font-bold text-[#0F172A]">
                            {t.report.kyujinStatsTitle}
                          </h3>
                          <span className="font-mono text-[11px] font-semibold text-[#64748B] tabular-nums">
                            {kyujinStats.totalCompanies} {t.report.companiesUnit} /{' '}
                            {kyujinStats.totalSlots} {t.report.slotsUnit}
                          </span>
                        </div>

                        <div className="grid grid-cols-3 gap-2 text-center">
                          <div className="rounded border border-emerald-200 bg-emerald-50 p-2.5">
                            <div className="text-[11px] font-semibold text-emerald-700">
                              {t.report.openLabel}
                            </div>
                            <div className="mt-0.5 font-mono text-lg font-bold text-[#0F172A] tabular-nums">
                              {kyujinStats.openSlots}
                            </div>
                            <div className="font-mono text-[10px] text-emerald-700 tabular-nums">
                              {Math.round((kyujinStats.openSlots / totalSlotsSafe) * 100)}%
                            </div>
                          </div>
                          <div className="rounded border border-amber-300 bg-amber-50 p-2.5">
                            <div className="text-[11px] font-semibold text-amber-800">
                              {t.report.bonusLabel}
                            </div>
                            <div className="mt-0.5 font-mono text-lg font-bold text-[#0F172A] tabular-nums">
                              {kyujinStats.bonusSlots}
                            </div>
                            <div className="font-mono text-[10px] text-amber-800 tabular-nums">
                              {Math.round((kyujinStats.bonusSlots / totalSlotsSafe) * 100)}%
                            </div>
                          </div>
                          <div className="rounded border border-sky-200 bg-sky-50 p-2.5">
                            <div className="text-[11px] font-semibold text-sky-700">
                              {t.report.housingLabel}
                            </div>
                            <div className="mt-0.5 font-mono text-lg font-bold text-[#0F172A] tabular-nums">
                              {housingAvailableSlots}
                            </div>
                            <div className="font-mono text-[10px] text-sky-700 tabular-nums">
                              {Math.round((housingAvailableSlots / totalSlotsSafe) * 100)}%
                            </div>
                          </div>
                        </div>

                        <div>
                          <div className="mb-2 text-xs font-bold text-[#64748B]">
                            {t.report.kyujinByIndTitle}
                          </div>
                          <div className="space-y-2.5">
                            {kyujinIndustryChartRows.map((row) => {
                              const pct = Math.round((row.total / totalSlotsSafe) * 100);
                              return (
                                <div key={row.industry} className="space-y-1">
                                  <div className="flex flex-wrap items-center justify-between gap-1 text-xs">
                                    <span className="font-semibold text-[#0F172A]">{row.industry}</span>
                                    <span className="font-mono font-bold text-[#0F172A] tabular-nums">
                                      {row.total} ({t.report.openLabel} {row.open} / {t.report.bonusLabel}{' '}
                                      {row.bonus})
                                    </span>
                                  </div>
                                  <div className="h-2 w-full overflow-hidden rounded-full bg-slate-100">
                                    <div
                                      style={{ width: `${Math.max(4, pct)}%` }}
                                      className="h-full rounded-full bg-amber-500"
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
                    <div className="rounded-lg border border-[#E2E8F0] bg-white p-4 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <div className="flex items-center gap-2">
                          <ShieldCheck className="h-4 w-4 text-emerald-700" aria-hidden="true" />
                          <h3 className="text-sm font-bold text-[#0F172A]">
                            Facebook Messenger Webhook (Meta Developers)
                          </h3>
                        </div>
                        <span className="rounded border border-emerald-200 bg-emerald-50 px-2 py-0.5 font-mono text-[10px] font-bold text-emerald-700">
                          {t.readyBadge}
                        </span>
                      </div>

                      <p className="mt-1.5 text-xs text-[#64748B]">
                        {uiLang === 'jp'
                          ? 'Fanpageからの新規メッセージは自動解析され「FB会社公式」に保存されます。'
                          : 'Tin nhắn mới từ ứng viên qua Fanpage sẽ được bóc tách tự động và lưu vào mục '}
                        {uiLang === 'vi' && <strong className="text-[#0F172A]">Fanpage cty</strong>}
                      </p>

                      <div className="mt-3 flex flex-wrap items-center gap-2">
                        <input
                          type="text"
                          readOnly
                          value={liveMessengerWebhookUrl}
                          aria-label="Webhook URL"
                          className="min-w-[260px] flex-1 rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-2 font-mono text-xs text-[#0F172A]"
                        />
                        <button
                          type="button"
                          onClick={handleCopyWebhookUrl}
                          className="inline-flex items-center gap-1.5 rounded-md bg-[#16A34A] px-4 py-2 text-xs font-semibold text-white hover:bg-[#15803D] whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          <Copy className="h-3.5 w-3.5" aria-hidden="true" />
                          {t.copyUrl}
                        </button>
                        <button
                          type="button"
                          onClick={openMessengerLinkModal}
                          className="inline-flex items-center gap-1.5 rounded-md border border-sky-200 bg-sky-50 px-3 py-2 text-xs font-semibold text-sky-700 hover:bg-sky-100 whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          <Link2 className="h-3.5 w-3.5" aria-hidden="true" />
                          {t.linkMessenger}
                        </button>
                        <button
                          type="button"
                          disabled={isRefreshingSql}
                          onClick={() => void refreshCrmFromSql(false)}
                          className="inline-flex items-center gap-1.5 rounded-md border border-[#E2E8F0] bg-white px-3 py-2 text-xs font-semibold text-[#0F172A] hover:bg-[#F8FAFC] disabled:opacity-50 whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          <RefreshCw
                            className={`h-3.5 w-3.5 ${isRefreshingSql ? 'animate-spin' : ''}`}
                            aria-hidden="true"
                          />
                          {uiLang === 'jp' ? 'Messenger同期' : 'Đồng bộ Messenger'}
                        </button>
                      </div>

                      <div className="mt-3 flex flex-wrap gap-3 font-mono text-[10px] text-[#64748B]">
                        <span>
                          Verify Token:{' '}
                          <strong
                            className={
                              webhookStatus?.hasVerifyToken ? 'text-emerald-700' : 'text-sky-700'
                            }
                          >
                            {webhookStatus?.hasVerifyToken ? 'Configured' : 'Optional'}
                          </strong>
                        </span>
                        <span>·</span>
                        <span>
                          App Secret:{' '}
                          <strong
                            className={
                              webhookStatus?.hasAppSecret ? 'text-emerald-700' : 'text-sky-700'
                            }
                          >
                            {webhookStatus?.hasAppSecret ? 'HMAC Protected' : 'Optional'}
                          </strong>
                        </span>
                        <span>·</span>
                        <span>
                          Account:{' '}
                          <strong className="text-emerald-700">
                            {authUser?.email || 'Active'}
                          </strong>
                        </span>
                      </div>
                    </div>

                    <div className="rounded-lg border border-[#E2E8F0] bg-white p-4 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
                      <div className="flex flex-wrap items-center justify-between gap-2">
                        <h3 className="text-sm font-bold text-[#0F172A]">{t.report.gasTitle}</h3>
                        <button
                          type="button"
                          onClick={async () => {
                            const ok = await copyTextSafely(GAS_CALENDAR_APP_SNIPPET);
                            showToast(ok ? `📋 ${t.report.copyGasCode}` : t.toasts.copyFailed);
                          }}
                          className="inline-flex items-center gap-1 rounded border border-sky-200 bg-sky-50 px-2.5 py-1 text-[11px] font-semibold text-sky-700 hover:bg-sky-100 cursor-pointer"
                        >
                          <Copy className="h-3 w-3" aria-hidden="true" />
                          {t.report.copyGasCode}
                        </button>
                      </div>
                      <p className="mt-1 text-xs text-[#64748B]">
                        {uiLang === 'jp'
                          ? '誤上書き防止のため初期状態は保護中です。同期が必要な場合のみヘッダーの保護ボタンを解除して送信してください。'
                          : 'Mặc định hệ thống khóa ghi đè GAS để bảo toàn dữ liệu. Chỉ mở khóa trên thanh tiêu đề khi cần đẩy đồng bộ.'}
                      </p>
                      <div className="mt-3 flex flex-wrap items-center gap-2">
                        <input
                          type="text"
                          value={gasWebAppUrl}
                          onChange={(e) => setGasWebAppUrl(e.target.value)}
                          placeholder="https://script.google.com/macros/s/..."
                          className="min-w-[280px] flex-1 rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-2 text-xs text-[#0F172A] placeholder:text-[#94A3B8] focus:border-[#16A34A] focus:bg-white focus:outline-none"
                        />
                        <button
                          type="button"
                          disabled={gasWriteLocked || isSyncingGas}
                          onClick={handleSyncToGas}
                          className="rounded-md bg-[#16A34A] px-4 py-2 text-xs font-semibold text-white hover:bg-[#15803D] disabled:opacity-40 whitespace-nowrap shrink-0 cursor-pointer"
                        >
                          {isSyncingGas ? t.report.sending : t.report.syncToGas}
                        </button>
                      </div>
                    </div>
                  </div>
                );
              })()}
          </main>

          {/* RIGHT SYSTEM ALERTS PANE */}
          <aside className="border-t border-[#E2E8F0] bg-white p-5 lg:overflow-y-auto lg:border-t-0 print:hidden">
            <div className="mb-4 font-mono text-[10px] font-semibold uppercase tracking-widest text-[#64748B]">
              {t.systemAlerts}
            </div>

            {/* Messenger Webhook Quick Status Card */}
            <div className="mb-4 rounded-lg border border-[#E2E8F0] bg-[#F8FAFC] p-3.5">
              <div className="flex items-center justify-between gap-2">
                <span className="font-mono text-[10px] font-semibold uppercase tracking-wider text-[#64748B]">
                  {t.messengerWebhook}
                </span>
                <span className="rounded bg-emerald-50 border border-emerald-200 px-1.5 py-0.5 font-mono text-[10px] font-bold text-emerald-700">
                  {t.readyBadge}
                </span>
              </div>
              <div
                className="mt-2 truncate font-mono text-[10px] text-[#64748B]"
                title={liveMessengerWebhookUrl}
              >
                {liveMessengerWebhookUrl}
              </div>
              <div className="mt-2.5 flex flex-col gap-1.5">
                <button
                  type="button"
                  onClick={openMessengerLinkModal}
                  className="inline-flex w-full items-center justify-center gap-1.5 rounded bg-sky-600 px-2.5 py-1.5 text-[11px] font-semibold text-white hover:bg-sky-700 whitespace-nowrap cursor-pointer"
                >
                  <Link2 className="h-3 w-3" aria-hidden="true" />
                  {t.linkMessenger}
                </button>
                <div className="flex gap-1.5">
                  <button
                    type="button"
                    onClick={handleCopyWebhookUrl}
                    className="inline-flex flex-1 items-center justify-center gap-1 rounded border border-[#E2E8F0] bg-white px-2.5 py-1 text-[10px] font-semibold text-[#0F172A] hover:border-emerald-300 whitespace-nowrap cursor-pointer"
                  >
                    <Copy className="h-3 w-3" aria-hidden="true" />
                    {t.copyUrl}
                  </button>
                  {fbPageLink && (
                    <a
                      href={fbPageLink}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="inline-flex items-center gap-1 rounded border border-[#E2E8F0] bg-white px-2.5 py-1 text-[10px] font-semibold text-sky-700 hover:border-sky-300 whitespace-nowrap"
                    >
                      <ExternalLink className="h-3 w-3" aria-hidden="true" />
                      {t.openPage}
                    </a>
                  )}
                </div>
              </div>
            </div>

            {/* Alert Card 1: Upcoming PV / Mendan + Google Calendar Sync */}
            <div className="mb-4 space-y-3">
              <div className="flex items-center justify-between gap-2">
                <span className="font-mono text-[10px] font-semibold uppercase tracking-wider text-[#64748B]">
                  {t.nextPvMendan}
                </span>
                <button
                  type="button"
                  disabled={isBulkSyncingCalendar}
                  onClick={requestBulkSyncUpcomingCalendar}
                  className="inline-flex items-center gap-1 rounded border border-emerald-200 bg-emerald-50 px-2 py-1 text-[10px] font-semibold text-emerald-700 hover:bg-emerald-100 disabled:opacity-50 whitespace-nowrap shrink-0 cursor-pointer"
                  title={
                    hasCalendarToken
                      ? t.syncAllUpcoming
                      : t.syncAllUpcoming
                  }
                >
                  <Calendar className="h-3 w-3" aria-hidden="true" />
                  {isBulkSyncingCalendar ? t.syncingCalendarBtn : t.syncAllUpcoming}
                </button>
              </div>

              {upcomingMendanAlerts.length > 0 ? (
                upcomingMendanAlerts.slice(0, 3).map(({ person, mendan }) => {
                  const isSchool = Boolean(person._isSchool || person.channel === 'school');
                  const hasSyncedCal = Boolean(person.googleCalendarEventId);
                  return (
                    <div
                      key={`mendan_${person.id}`}
                      className="rounded-lg border border-emerald-200 bg-white p-3.5 shadow-[0_1px_2px_rgba(0,0,0,0.03)]"
                    >
                      <div className="mb-1.5 flex items-center justify-between gap-2">
                        <span className="font-mono text-[10px] font-semibold uppercase tracking-wider text-emerald-700">
                          {t.nextPvMendan}
                        </span>
                        {hasSyncedCal && (
                          <span className="rounded bg-emerald-50 border border-emerald-200 px-1.5 py-0.5 font-mono text-[9px] font-bold text-emerald-700">
                            {t.calendarOk}
                          </span>
                        )}
                      </div>
                      <div className="flex items-center gap-2 text-xs font-semibold text-[#0F172A]">
                        <span className="inline-block h-2 w-2 rounded-full bg-[#16A34A]" />
                        <span className="truncate">{getCandidateDisplayName(person, uiLang)}</span>
                      </div>
                      <div className="mt-1 font-mono text-[11px] text-[#64748B] tabular-nums">
                        {person.interviewDate || mendan.dateLabel || person.progressStatus}
                      </div>
                      <div className="mt-3 flex flex-wrap gap-1.5">
                        <button
                          type="button"
                          onClick={() =>
                            setReplyModal({ person, scenario: 'remind_interview' })
                          }
                          className="flex-1 rounded bg-[#16A34A] px-2.5 py-1.5 text-[11px] font-semibold text-white hover:bg-[#15803D] whitespace-nowrap cursor-pointer"
                        >
                          {t.remindScheduleBtn}
                        </button>
                        <button
                          type="button"
                          onClick={() => handleMarkContactedToday(person, isSchool)}
                          className="flex-1 rounded border border-[#E2E8F0] bg-[#F8FAFC] px-2.5 py-1.5 text-[11px] font-semibold text-[#475569] hover:border-[#CBD5E1] hover:text-[#0F172A] whitespace-nowrap cursor-pointer"
                        >
                          {t.contactedShort}
                        </button>
                      </div>

                      <div className="mt-2 flex items-center gap-1.5 border-t border-dashed border-[#E2E8F0] pt-2">
                        <button
                          type="button"
                          disabled={syncingCalendarId === person.id}
                          onClick={() => requestSyncCandidateCalendar(person)}
                          className={`inline-flex flex-1 items-center justify-center gap-1 rounded border px-2 py-1 text-[10px] font-semibold transition-colors whitespace-nowrap cursor-pointer disabled:opacity-50 ${
                            hasSyncedCal
                              ? 'border-emerald-200 bg-emerald-50 text-emerald-700'
                              : 'border-[#E2E8F0] bg-[#F8FAFC] text-[#0F172A] hover:border-emerald-300'
                          }`}
                        >
                          <Calendar className="h-3 w-3" aria-hidden="true" />
                          {syncingCalendarId === person.id
                            ? t.syncingCalendarBtn
                            : hasSyncedCal
                            ? t.updateCalendarBtn
                            : t.syncCalendarBtn}
                        </button>
                        {hasSyncedCal &&
                          isValidGoogleCalendarUrl(person.googleCalendarEventLink) && (
                            <a
                              href={person.googleCalendarEventLink}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="inline-flex items-center gap-1 rounded border border-[#E2E8F0] bg-[#F8FAFC] px-2 py-1 text-[10px] font-semibold text-[#64748B] hover:border-emerald-300 hover:text-emerald-700 whitespace-nowrap"
                              title={t.openCalendarBtn}
                            >
                              <ExternalLink className="h-3 w-3" aria-hidden="true" />
                              {t.openShort}
                            </a>
                          )}
                      </div>
                    </div>
                  );
                })
              ) : (
                <div className="rounded-lg border border-[#E2E8F0] bg-[#F8FAFC] p-3.5">
                  <div className="font-mono text-[10px] font-semibold uppercase tracking-wider text-[#64748B]">
                    {t.nextPvMendan}
                  </div>
                  <p className="mt-1.5 text-xs text-[#64748B]">{t.noUpcomingMendan}</p>
                </div>
              )}
            </div>

            {/* Alert Card 2: Critical Follow-Up */}
            {urgentFollowUpList.length > 0 && (
              <div className="mb-4 rounded-lg border border-rose-200 bg-rose-50/70 p-3.5">
                <div className="mb-2 font-mono text-[10px] font-semibold uppercase tracking-wider text-rose-700">
                  {t.criticalFollowup} ({urgentFollowUpList.length})
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
                        className="truncate text-left font-medium text-[#0F172A] hover:text-rose-700 cursor-pointer"
                      >
                        <strong>{getCandidateDisplayName(person, uiLang)}</strong>{' '}
                        <span className="text-[#64748B]">({reason})</span>
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
                  className="mt-3 text-[11px] font-semibold text-rose-700 underline cursor-pointer"
                >
                  {careFilter === 'urgent'
                    ? t.clearFilterArrow
                    : t.showAllUrgentArrow(urgentFollowUpList.length)}
                </button>
              </div>
            )}

            {/* Daily Focus Checklist */}
            <div className="mb-4 rounded-lg border border-[#E2E8F0] bg-[#F8FAFC] p-3.5">
              <div className="mb-2 flex items-center justify-between gap-2">
                <span className="font-mono text-[10px] font-semibold uppercase tracking-wider text-emerald-700">
                  {t.dailyFocusTitle}
                </span>
                <span className="font-mono text-[10px] text-[#64748B] tabular-nums">
                  {dailyFocusItems.filter((i) => dailyFocusCompleted[i.id]).length}/
                  {dailyFocusItems.length}
                </span>
              </div>
              {dailyFocusItems.length === 0 ? (
                <p className="text-xs text-[#64748B]">{t.dailyFocusEmpty}</p>
              ) : (
                <div className="max-h-60 space-y-1.5 overflow-y-auto pr-1">
                  {dailyFocusItems.map((item) => {
                    const checked = Boolean(dailyFocusCompleted[item.id]);
                    return (
                      <div
                        key={item.id}
                        className={`flex items-start gap-2 rounded border px-2.5 py-2 text-xs transition-colors ${
                          checked
                            ? 'border-[#E2E8F0] bg-white/60 opacity-60'
                            : item.urgency === 'critical'
                            ? 'border-rose-200 bg-rose-50/70'
                            : 'border-[#E2E8F0] bg-white'
                        }`}
                      >
                        <button
                          type="button"
                          onClick={() => toggleDailyFocusItem(item.id)}
                          aria-label={item.title}
                          className="mt-0.5 shrink-0 text-[#16A34A] cursor-pointer"
                        >
                          {checked ? (
                            <CheckSquare className="h-4 w-4" aria-hidden="true" />
                          ) : (
                            <Square className="h-4 w-4 text-[#94A3B8]" aria-hidden="true" />
                          )}
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            setActiveTab(item.targetTab);
                            if (item.category === 'visa_submission') {
                              setProposedStageFilter('naitei');
                            }
                            if (
                              item.entityId &&
                              item.category !== 'resignation' &&
                              item.category !== 'resignation_unconfirmed'
                            ) {
                              setDrawerCandidateId(item.entityId);
                            }
                          }}
                          className="min-w-0 flex-1 text-left cursor-pointer"
                        >
                          <div
                            className={`font-semibold ${
                              checked
                                ? 'line-through text-[#94A3B8]'
                                : 'text-[#0F172A] hover:text-emerald-700'
                            }`}
                          >
                            {item.title}
                          </div>
                          <div className="truncate text-[10px] text-[#64748B]">
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
                  <span className="font-mono text-[10px] font-semibold uppercase tracking-wider text-amber-800">
                    {t.tabs.resignation} ({resignationAlerts.length})
                  </span>
                  <button
                    type="button"
                    onClick={() => setActiveTab('resignation')}
                    className="text-[10px] font-semibold text-sky-700 hover:underline cursor-pointer"
                  >
                    {uiLang === 'jp' ? 'すべて表示 →' : 'Xem tất cả →'}
                  </button>
                </div>
                {resignationAlerts.slice(0, 3).map(({ record, assessment }) => {
                  const isCrit =
                    assessment.urgency === 'overdue' || assessment.urgency === 'urgent3';
                  return (
                    <div
                      key={`resig_alert_${record.id}`}
                      className={`rounded-lg border p-3 text-xs ${
                        isCrit
                          ? 'border-rose-200 bg-rose-50/60'
                          : 'border-amber-200 bg-amber-50/60'
                      }`}
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span className="font-semibold text-[#0F172A] truncate">
                          {getResignationDisplayName(record, uiLang)}
                        </span>
                        <span
                          className={`rounded px-1.5 py-0.5 font-mono text-[9px] font-bold ${
                            isCrit
                              ? 'bg-rose-100 text-rose-700'
                              : 'bg-amber-100 text-amber-800'
                          }`}
                        >
                          {assessment.label}
                        </span>
                      </div>
                      <div className="mt-1 text-[11px] text-[#64748B] truncate">
                        {getResignationCompanyLabel(record.previousCompany, uiLang)}
                      </div>
                      <div className="mt-2.5 flex flex-wrap gap-1.5">
                        <button
                          type="button"
                          onClick={() => setResignationReminderModal(record)}
                          className="flex-1 rounded border border-amber-300 bg-amber-50 px-2 py-1 text-[10px] font-semibold text-amber-800 hover:bg-amber-100 whitespace-nowrap cursor-pointer"
                        >
                          {uiLang === 'jp' ? 'リマインド文' : 'Tạo tin nhắn nhắc'}
                        </button>
                        <button
                          type="button"
                          onClick={() => void handleSyncResignationCalendarSingle(record)}
                          className="rounded border border-emerald-200 bg-emerald-50 px-2 py-1 text-[10px] font-semibold text-emerald-700 hover:bg-emerald-100 whitespace-nowrap cursor-pointer"
                          title={t.syncCalendarBtn}
                        >
                          <Calendar className="h-3 w-3" aria-hidden="true" />
                        </button>
                        <button
                          type="button"
                          onClick={() => handleConvertResignationToCandidate(record)}
                          className="rounded border border-sky-200 bg-sky-50 px-2 py-1 text-[10px] font-semibold text-sky-700 hover:bg-sky-100 whitespace-nowrap cursor-pointer"
                          title={
                            uiLang === 'jp'
                              ? '求職者一覧へ転記'
                              : 'Chuyển sang Ứng viên tìm việc mới'
                          }
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
                    className="rounded-lg border border-dashed border-amber-300 bg-amber-50/50 p-3.5"
                  >
                    <div className="mb-1.5 font-mono text-[10px] font-semibold uppercase tracking-wider text-amber-800">
                      {uiLang === 'jp' ? '情報不足プロフィール' : 'Hồ sơ thiếu thông tin'}
                    </div>
                    <div className="text-xs font-semibold text-[#0F172A]">
                      {getCandidateDisplayName(person, uiLang)}
                    </div>
                    <div className="mt-1 font-mono text-[10px] text-[#64748B]">
                      {uiLang === 'jp' ? '不足: ' : 'Thiếu: '}
                      {missing.map((m) => (uiLang === 'jp' ? m.ja : m.vi)).join(', ')}
                    </div>
                    <button
                      type="button"
                      onClick={() => setReplyModal({ person, scenario: 'missing_info_request' })}
                      className="mt-3 w-full rounded border border-amber-300 bg-white px-2.5 py-1.5 text-[11px] font-semibold text-amber-800 hover:bg-amber-100 cursor-pointer"
                    >
                      {t.cardLabels.askMissingBtn}
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
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-slate-900/50 backdrop-blur-[1px] p-4">
          <div
            className="max-h-[90vh] w-full max-w-3xl overflow-y-auto rounded-xl border border-[#E2E8F0] bg-white p-5 text-xs text-[#0F172A] shadow-2xl"
            onPaste={(e) => {
              const img = extractFirstImageFromClipboard(e.clipboardData);
              if (img) {
                e.preventDefault();
                void handleExtractCandidateScreenshot(img, 'edit');
              }
            }}
          >
            <div className="mb-3 flex items-center justify-between border-b border-[#E2E8F0] pb-2.5">
              <h3 className="text-base font-bold text-[#0F172A]">
                {editingPerson.isNew
                  ? uiLang === 'jp'
                    ? '＋ 新規相談者登録'
                    : '+ Thêm ứng viên mới'
                  : `${uiLang === 'jp' ? 'プロフィール編集' : 'Chỉnh sửa hồ sơ'}：${getCandidateDisplayName(
                      editingPerson.data,
                      uiLang
                    )}`}
              </h3>
              <button
                type="button"
                onClick={() => setEditingPerson(null)}
                className="rounded-lg bg-slate-100 px-3 py-1 font-semibold text-slate-700 hover:bg-slate-200 cursor-pointer"
              >
                {t.close}
              </button>
            </div>

            {editingPerson.isNew &&
              (() => {
                const dup = findDuplicateCandidate(allPeopleCombined, {
                  id: editingPerson.data.id,
                  name: editingPerson.data.name,
                  fbName: editingPerson.data.fbName,
                  fbProfileUrl: editingPerson.data.fbProfileUrl,
                  phone: editingPerson.data.phone,
                });
                if (!dup) return null;
                const dupName = getCandidateDisplayName(dup, uiLang);
                const matchedBy =
                  editingPerson.data.fbProfileUrl && dup.fbProfileUrl
                    ? 'FB URL'
                    : editingPerson.data.phone && dup.phone
                      ? uiLang === 'jp'
                        ? '電話番号'
                        : 'SĐT'
                      : uiLang === 'jp'
                        ? '氏名 / FB名'
                        : 'Tên / Tên FB';
                return (
                  <div className="mb-3 flex flex-wrap items-center justify-between gap-2 rounded-lg border border-amber-300 bg-amber-50 px-3 py-2 text-[11px] text-amber-900">
                    <span className="font-semibold">
                      {uiLang === 'jp'
                        ? `⚠️ 重複候補を検出しました（一致項目: ${matchedBy}）:「${dupName}」(${dup.progressStatus || '新規相談'})`
                        : `⚠️ Phát hiện hồ sơ có thể bị trùng (khớp ${matchedBy}): "${dupName}" (${dup.progressStatus || '新規相談'})`}
                    </span>
                    <button
                      type="button"
                      onClick={() => {
                        setEditingPerson(null);
                        setDrawerCandidateId(dup.id);
                      }}
                      className="rounded border border-amber-400 bg-white px-2.5 py-0.5 font-semibold text-amber-900 hover:bg-amber-100 cursor-pointer"
                    >
                      {uiLang === 'jp' ? '既存プロフィールを開く' : 'Mở hồ sơ sẵn có'}
                    </button>
                  </div>
                );
              })()}

            <div className="mb-4 rounded-lg border border-sky-200 bg-sky-50/70 p-3">
              <div className="mb-1.5 flex flex-wrap items-center justify-between gap-2">
                <span className="font-bold text-sky-950">
                  {uiLang === 'jp'
                    ? 'チャット・メモ貼り付け または スクリーンショット（Ctrl+V）で自動解析'
                    : 'Dán đoạn chat / ghi chú dài hoặc Ctrl+V ảnh chụp Messenger / Thẻ ngoại kiều'}
                </span>
                <div className="flex flex-wrap items-center gap-1.5">
                  <label className="inline-flex cursor-pointer items-center gap-1 rounded-md border border-sky-300 bg-white px-2.5 py-1 text-[11px] font-semibold text-sky-800 hover:bg-sky-50">
                    <ImageIcon className="h-3.5 w-3.5" aria-hidden="true" />
                    {uiLang === 'jp' ? '画像をアップロード' : 'Tải ảnh chụp'}
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
                    className="inline-flex items-center gap-1 rounded-md bg-sky-600 px-3 py-1 text-[11px] font-semibold text-white hover:bg-sky-700 disabled:opacity-60 cursor-pointer"
                  >
                    <Sparkles className="h-3 w-3" aria-hidden="true" />
                    {isExtractingCandidateAi
                      ? uiLang === 'jp'
                        ? '解析中...'
                        : 'Đang bóc tách...'
                      : uiLang === 'jp'
                      ? 'AI自動抽出（原文保持）'
                      : 'AI bóc tách nguyên văn'}
                  </button>
                </div>
              </div>
              {currentImageBase64 && (
                <div className="mb-2 flex items-center gap-3 rounded-lg border border-sky-200 bg-white p-2">
                  <img
                    src={currentImageBase64.dataUrl}
                    alt={currentImageBase64.fileName || (uiLang === 'jp' ? 'スクリーンショット' : 'Ảnh chụp màn hình')}
                    className="h-20 w-28 shrink-0 rounded border border-slate-200 object-cover"
                  />
                  <div className="min-w-0 flex-1">
                    <div className="truncate font-semibold text-slate-800">
                      📷 {currentImageBase64.fileName || (uiLang === 'jp' ? 'クリップボード画像' : 'Ảnh chụp từ Clipboard')}
                    </div>
                    <div className="mt-0.5 text-[11px] text-slate-500">
                      {uiLang === 'jp'
                        ? `最適化済み（${Math.max(1, Math.round((currentImageBase64.base64Data.length * 0.75) / 1024))} KB）• 必要に応じて「AI自動抽出」で再解析できます`
                        : `Đã tối ưu kích thước (${Math.max(1, Math.round((currentImageBase64.base64Data.length * 0.75) / 1024))} KB) • Bấm "AI bóc tách nguyên văn" để phân tích lại nếu cần`}
                    </div>
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      setCurrentImageBase64(null);
                      setCandidateExtractBanner(null);
                    }}
                    className="inline-flex shrink-0 items-center gap-1 rounded border border-slate-200 bg-slate-50 px-2 py-1 text-[11px] font-semibold text-slate-600 hover:bg-slate-100 cursor-pointer"
                    title={uiLang === 'jp' ? '画像を削除' : 'Xóa ảnh đính kèm'}
                  >
                    <X className="h-3 w-3" aria-hidden="true" />
                    {uiLang === 'jp' ? '画像削除' : 'Xóa ảnh'}
                  </button>
                </div>
              )}
              <textarea
                rows={3}
                value={editChatSnippet}
                onChange={(e) => setEditChatSnippet(e.target.value)}
                placeholder={
                  uiLang === 'jp'
                    ? 'Messengerの会話内容、給与・条件交渉メモを貼り付けるか、Ctrl+Vでスクリーンショットを貼り付けてください...'
                    : 'Dán đoạn chat Messenger, ghi chú đàm phán lương/điều kiện hoặc nhấn Ctrl+V ảnh chụp màn hình vào đây...'
                }
                className="w-full rounded-lg border border-sky-200 bg-white p-2 text-xs"
              />
              {candidateExtractBanner && (
                <div
                  className={`mt-1.5 rounded border px-2.5 py-1.5 text-[11px] font-semibold ${
                    candidateExtractBanner.type === 'success'
                      ? 'border-emerald-300 bg-emerald-50 text-emerald-800'
                      : candidateExtractBanner.type === 'warning'
                      ? 'border-amber-300 bg-amber-50 text-amber-900'
                      : 'border-sky-300 bg-white text-sky-900'
                  }`}
                >
                  {candidateExtractBanner.text}
                </div>
              )}
              {!editingPerson.data.name && editingPerson.data.fbName && (
                <div className="mt-1.5 rounded border border-amber-200 bg-amber-50 px-2.5 py-1 text-[11px] font-medium text-amber-900">
                  {uiLang === 'jp'
                    ? `⚠️ 会話からはFacebook名（${editingPerson.data.fbName}）のみ取得され、本名が未確認です。そのまま保存するか、下欄に本名を入力できます。`
                    : `⚠️ Đoạn chat mới ghi nhận Tên Facebook (${editingPerson.data.fbName}), chưa thấy họ tên khai sinh. Có thể bấm Lưu để dùng tạm tên FB hoặc nhập họ tên thật bên dưới.`}
                </div>
              )}
            </div>

            <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
              <div>
                <label className="mb-1 block font-semibold">
                  {uiLang === 'jp' ? '氏名（本名）' : 'Họ tên khai sinh'}
                </label>
                <input
                  type="text"
                  placeholder={uiLang === 'jp' ? 'FB名のみの場合は空欄可' : 'Để trống nếu mới có tên FB'}
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
                <label className="mb-1 block font-semibold">
                  {uiLang === 'jp' ? 'FBアカウント名' : 'Tên nick Facebook'}
                </label>
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
                <label className="mb-1 block font-semibold">
                  {uiLang === 'jp' ? 'Facebook / Messenger リンク' : 'Link Facebook / Messenger'}
                </label>
                <input
                  type="url"
                  placeholder="https://m.me/... / https://facebook.com/..."
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
                <label className="mb-1 block font-semibold">
                  {uiLang === 'jp' ? '流入チャネル（CRMルート切替）' : 'Nguồn kênh CRM (Chuyển đổi mượt)'}
                </label>
                <div className="flex flex-wrap gap-2">
                  {(
                    [
                      { id: 'personal', label: t.tabs.personal },
                      { id: 'company', label: t.tabs.company },
                      { id: 'line', label: t.tabs.line },
                      { id: 'other', label: t.tabs.other },
                      { id: 'school', label: t.tabs.school },
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
                      className={`rounded-lg border px-3 py-1.5 text-[11px] font-semibold transition-colors cursor-pointer ${
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
                <label className="mb-1 block font-semibold">
                  {uiLang === 'jp' ? '国籍' : 'Quốc tịch'}
                </label>
                {(() => {
                  const currentNat = String(editingPerson.data.nationality || '').trim();
                  const isStandardNat =
                    !currentNat ||
                    (CANDIDATE_NATIONALITIES as readonly string[]).includes(currentNat);
                  const selectVal = isStandardNat ? currentNat : 'その他';
                  return (
                    <div className="space-y-1.5">
                      <select
                        value={selectVal}
                        onChange={(e) => {
                          const val = e.target.value;
                          setEditingPerson({
                            ...editingPerson,
                            data: {
                              ...editingPerson.data,
                              nationality:
                                val === 'その他'
                                  ? isStandardNat
                                    ? 'その他'
                                    : currentNat
                                  : val,
                              nationalityManuallyEdited: true,
                            },
                          });
                        }}
                        className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5"
                      >
                        <option value="">
                          {uiLang === 'jp' ? '-- 未設定 --' : '-- Chưa chọn --'}
                        </option>
                        {CANDIDATE_NATIONALITIES.map((nat) => (
                          <option key={nat} value={nat}>
                            {nat}
                          </option>
                        ))}
                      </select>
                      {selectVal === 'その他' && (
                        <input
                          type="text"
                          placeholder={
                            uiLang === 'jp'
                              ? '国名を入力 (例: バングラデシュ)'
                              : 'Nhập quốc tịch khác...'
                          }
                          value={currentNat === 'その他' ? '' : currentNat}
                          onChange={(e) =>
                            setEditingPerson({
                              ...editingPerson,
                              data: {
                                ...editingPerson.data,
                                nationality: e.target.value || 'その他',
                                nationalityManuallyEdited: true,
                              },
                            })
                          }
                          className="w-full rounded-lg border border-slate-300 px-2.5 py-1.5"
                        />
                      )}
                    </div>
                  );
                })()}
              </div>
              <div>
                <label className="mb-1 block font-semibold">
                  {uiLang === 'jp' ? '年齢 / 生年月日' : 'Tuổi / Ngày sinh'}
                </label>
                <div className="flex gap-1.5">
                  <input
                    type="text"
                    placeholder={uiLang === 'jp' ? '年齢' : 'Tuổi'}
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
                <label className="mb-1 block font-semibold">
                  {uiLang === 'jp' ? '日本語レベル' : 'Trình độ tiếng Nhật'}
                </label>
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
                <label className="mb-1 block font-semibold">
                  {uiLang === 'jp' ? '希望分野' : 'Ngành nghề mong muốn'}
                </label>
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
                <label className="mb-1 block font-semibold">
                  {uiLang === 'jp' ? '希望勤務エリア' : 'Khu vực mong muốn'}
                </label>
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
                <label className="mb-1 block font-semibold">
                  {uiLang === 'jp' ? '進捗ステータス' : 'Trạng thái tiến độ'}
                </label>
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
                <label className="mb-1 block font-semibold">
                  {uiLang === 'jp' ? '在留資格 / 在留期限' : 'Tư cách lưu trú / Hạn visa'}
                </label>
                <div className="flex gap-1.5">
                  <input
                    type="text"
                    placeholder={uiLang === 'jp' ? '在留資格' : 'Loại visa'}
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
                    placeholder={uiLang === 'jp' ? '期限 YYYY-MM-DD' : 'Hạn YYYY-MM-DD'}
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
                <label className="mb-1 block font-semibold">
                  {uiLang === 'jp' ? '合格済み試験・資格' : 'Kỳ thi / Chứng chỉ đã đỗ'}
                </label>
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
                <label className="mb-1 block font-semibold">
                  {uiLang === 'jp' ? '提案先企業' : 'Công ty đang đề xuất'}
                </label>
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
                  {uiLang === 'jp' ? '面談日時 (YYYY-MM-DD HH:mm)' : 'Lịch Mendan (YYYY-MM-DD HH:mm)'}
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
                  {uiLang === 'jp'
                    ? '企業面接日時 (YYYY-MM-DD HH:mm)'
                    : 'Lịch phỏng vấn công ty (YYYY-MM-DD HH:mm)'}
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
                <label className="mb-1 block font-semibold">
                  {uiLang === 'jp' ? '電話番号 / LINE' : 'Số điện thoại / LINE'}
                </label>
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
                <label className="mb-1 block font-semibold">
                  {uiLang === 'jp' ? '入社可能時期' : 'Thời điểm nhận việc'}
                </label>
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
                  {uiLang === 'jp'
                    ? '日本での職歴・経験・資格詳細'
                    : 'Kinh nghiệm & Bằng cấp chi tiết tại Nhật'}
                </label>
                <input
                  type="text"
                  placeholder={
                    uiLang === 'jp'
                      ? '例：有料老人ホーム介護職3年経験、特定技能介護試験合格...'
                      : 'VD: Đang là Kaigo năm cuối, 3 năm viện có phí...'
                  }
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
                  {uiLang === 'jp' ? '給与希望・交渉ライン' : 'Lương kỳ vọng / Mốc đàm phán'}
                </label>
                <input
                  type="text"
                  placeholder={
                    uiLang === 'jp'
                      ? '例：家賃3万円控除後 手取り16万円以上希望'
                      : 'VD: Kỳ vọng thực nhận ~16 man sau trừ nhà 3 man'
                  }
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
                {uiLang === 'jp'
                  ? '転職・在留資格変更の目的・希望条件'
                  : 'Nhu cầu & Mục tiêu chuyển đổi việc / visa'}
              </label>
              <input
                type="text"
                placeholder={
                  uiLang === 'jp'
                    ? '例：介護分野で特定技能への在留資格変更希望'
                    : 'VD: Tìm công ty mới phỏng vấn chuyển đổi visa Tokutei ngành Kaigo'
                }
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

            <div className="mt-3 grid grid-cols-1 gap-3 rounded-lg border border-[#E2E8F0] bg-[#F8FAFC] p-3 sm:grid-cols-2">
              <div>
                <label className="mb-1 block font-semibold">
                  {t.naiteiModal.companyLabel}
                </label>
                <input
                  type="text"
                  placeholder={t.naiteiModal.companyPlaceholder}
                  value={editingPerson.data.naiteiCompany || ''}
                  onChange={(e) =>
                    setEditingPerson({
                      ...editingPerson,
                      data: { ...editingPerson.data, naiteiCompany: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="mb-1 block font-semibold">
                  {t.naiteiModal.joinDateLabel}
                </label>
                <input
                  type="date"
                  value={editingPerson.data.expectedJoinDate || ''}
                  onChange={(e) =>
                    setEditingPerson({
                      ...editingPerson,
                      data: { ...editingPerson.data, expectedJoinDate: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 tabular-nums"
                />
              </div>
              <div>
                <label className="mb-1 block font-semibold">
                  {t.naiteiModal.visaDeadlineLabel}
                </label>
                <input
                  type="date"
                  value={editingPerson.data.visaSubmissionDeadline || ''}
                  onChange={(e) =>
                    setEditingPerson({
                      ...editingPerson,
                      data: {
                        ...editingPerson.data,
                        visaSubmissionDeadline: e.target.value,
                      },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5 tabular-nums"
                />
              </div>
              <div>
                <label className="mb-1 block font-semibold">
                  {t.jitaiModal.reasonLabel}
                </label>
                <input
                  type="text"
                  placeholder={t.jitaiModal.placeholder}
                  value={editingPerson.data.declineReason || ''}
                  onChange={(e) =>
                    setEditingPerson({
                      ...editingPerson,
                      data: { ...editingPerson.data, declineReason: e.target.value },
                    })
                  }
                  className="w-full rounded-lg border border-slate-300 bg-white px-2.5 py-1.5"
                />
              </div>
            </div>

            <div className="mt-3">
              <label className="mb-1 block font-semibold">
                {uiLang === 'jp'
                  ? '相談・条件交渉ログ（原文のまま保持）'
                  : 'Nhật ký & Tình trạng trao đổi nguyên văn (Không cắt bớt)'}
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
                placeholder={
                  uiLang === 'jp'
                    ? '給与や条件の交渉経緯、応募中の求人、本人の返信内容などをそのまま記録...'
                    : 'Lưu nguyên văn các mốc đàm phán lương/điều kiện, đơn đang hỏi, phản hồi của ứng viên...'
                }
                className="w-full rounded-lg border border-slate-300 p-2.5"
              />
            </div>

            <div className="mt-3">
              <label className="mb-1 block font-semibold">
                {uiLang === 'jp' ? '補足メモ（面談記録・対応履歴）' : 'Ghi chú bổ sung (Lịch sử tư vấn)'}
              </label>
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
                className="rounded-lg border border-slate-300 px-4 py-2 font-semibold text-slate-700 hover:bg-slate-50 cursor-pointer"
              >
                {t.cancel}
              </button>
              <button
                type="button"
                onClick={handleSaveEditingPerson}
                className="rounded-lg bg-[#16A34A] px-5 py-2 font-semibold text-white hover:bg-[#15803D] cursor-pointer"
              >
                {t.save}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 2: FB CHAT AUTO-REGISTER / UPDATE */}
      {fbChatModalOpen && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-slate-900/50 backdrop-blur-[1px] p-4">
          <div
            className="max-h-[90vh] w-full max-w-2xl overflow-y-auto rounded-xl border border-[#E2E8F0] bg-white p-5 text-xs text-[#0F172A] shadow-2xl"
            onPaste={(e) => {
              const img = extractFirstImageFromClipboard(e.clipboardData);
              if (img) {
                e.preventDefault();
                void handleExtractCandidateScreenshot(img, 'fbChat');
              }
            }}
          >
            <div className="mb-3 flex items-center justify-between border-b border-[#E2E8F0] pb-2.5">
              <h3 className="text-base font-bold text-[#0F172A]">
                {uiLang === 'jp'
                  ? 'FBチャットから自動登録・既存相談者の状況更新'
                  : 'Đăng ký tự động từ chat FB / Cập nhật hồ sơ có sẵn'}
              </h3>
              <button
                type="button"
                onClick={() => setFbChatModalOpen(false)}
                className="rounded-lg bg-slate-100 px-3 py-1 font-semibold text-slate-700 hover:bg-slate-200 cursor-pointer"
              >
                {t.close}
              </button>
            </div>

            <div className="mb-3 flex flex-wrap gap-2">
              <button
                type="button"
                onClick={() => setFbChatMode('new')}
                className={`rounded-lg px-3 py-2 font-semibold cursor-pointer ${
                  fbChatMode === 'new'
                    ? 'bg-[#16A34A] text-white'
                    : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                }`}
              >
                {uiLang === 'jp' ? '＋ 新規相談者として登録' : '+ Đăng ký ứng viên mới'}
              </button>
              <button
                type="button"
                onClick={() => setFbChatMode('update')}
                className={`rounded-lg px-3 py-2 font-semibold cursor-pointer ${
                  fbChatMode === 'update'
                    ? 'bg-sky-600 text-white'
                    : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                }`}
              >
                {uiLang === 'jp' ? '既存相談者の状況を更新' : 'Cập nhật ứng viên hiện có'}
              </button>
            </div>

            {fbChatMode === 'update' && (
              <div className="mb-3">
                <label className="mb-1 block font-semibold">
                  {uiLang === 'jp' ? '更新対象の相談者を選択' : 'Chọn ứng viên cần cập nhật'}
                </label>
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
                          : found.channel === 'line'
                          ? 'line'
                          : found.channel === 'other'
                          ? 'other'
                          : 'personal'
                      );
                    }
                  }}
                  className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 font-semibold"
                >
                  <option value="">
                    {uiLang === 'jp' ? '-- 相談者を選択してください --' : '-- Chọn ứng viên --'}
                  </option>
                  {allPeopleCombined.map((p) => {
                    const disp = getCandidateDisplayName(p, uiLang);
                    return (
                      <option key={p.id} value={p.id}>
                        {disp} {p.fbName && p.fbName !== disp ? `(${p.fbName})` : ''} -{' '}
                        {p.progressStatus}
                      </option>
                    );
                  })}
                </select>
              </div>
            )}

            {fbChatMode === 'new' &&
              fbChatRawText.trim().length > 3 &&
              (() => {
                const previewParsed = parseChatIntoUpdates(fbChatRawText);
                const dup = findDuplicateCandidate(allPeopleCombined, {
                  name: previewParsed.name,
                  fbName: previewParsed.fbName,
                  fbProfileUrl: previewParsed.fbProfileUrl,
                  phone: previewParsed.phone,
                });
                if (!dup) return null;
                const dupName = getCandidateDisplayName(dup, uiLang);
                const matchedBy =
                  previewParsed.fbProfileUrl && dup.fbProfileUrl
                    ? 'FB URL'
                    : previewParsed.phone && dup.phone
                      ? uiLang === 'jp'
                        ? '電話番号'
                        : 'SĐT'
                      : uiLang === 'jp'
                        ? '氏名 / FB名'
                        : 'Tên / Tên FB';
                return (
                  <div className="mb-3 flex flex-wrap items-center justify-between gap-2 rounded-lg border border-amber-300 bg-amber-50 px-3 py-2 text-[11px] text-amber-900">
                    <span className="font-semibold">
                      {uiLang === 'jp'
                        ? `⚠️ 既存データに一致する相談者があります（${matchedBy}）:「${dupName}」`
                        : `⚠️ Có hồ sơ trùng (${matchedBy}): "${dupName}"`}
                    </span>
                    <button
                      type="button"
                      onClick={() => {
                        setFbChatMode('update');
                        setFbChatTargetId(dup.id);
                      }}
                      className="rounded border border-amber-400 bg-white px-2.5 py-0.5 font-semibold text-amber-900 hover:bg-amber-100 cursor-pointer"
                    >
                      {uiLang === 'jp' ? 'この相談者を更新モードへ切替' : 'Chuyển sang Cập nhật hồ sơ này'}
                    </button>
                  </div>
                );
              })()}

            <div className="mb-3">
              <label className="mb-1 block font-semibold">
                {fbChatMode === 'new'
                  ? uiLang === 'jp'
                    ? '登録先チャネル'
                    : 'Chọn kênh nguồn lưu hồ sơ'
                  : uiLang === 'jp'
                  ? 'チャネル設定'
                  : 'Kênh lưu trữ'}
              </label>
              <div className="flex flex-wrap gap-2">
                {(
                  [
                    { id: 'personal', label: t.channelPersonal },
                    { id: 'company', label: t.channelCompany },
                    { id: 'line', label: t.channelLine },
                    { id: 'other', label: t.channelOther },
                    { id: 'school', label: t.tabs.school },
                  ] as Array<{ id: ChannelType; label: string }>
                ).map((ch) => (
                  <button
                    key={ch.id}
                    type="button"
                    onClick={() => setFbChatChannel(ch.id)}
                    className={`rounded-lg border px-3 py-1.5 font-semibold cursor-pointer ${
                      fbChatChannel === ch.id
                        ? 'border-emerald-300 bg-emerald-50 text-emerald-800'
                        : 'border-slate-200 bg-white text-slate-700 hover:bg-slate-50'
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
                  {uiLang === 'jp'
                    ? 'Messenger履歴・メモ貼り付け または スクリーンショット（Ctrl+V）'
                    : 'Dán đoạn chat Messenger, ghi chú hoặc Ctrl+V ảnh chụp màn hình'}
                </label>
                <label className="inline-flex cursor-pointer items-center gap-1 rounded-md border border-sky-300 bg-sky-50 px-2.5 py-1 text-[11px] font-semibold text-sky-800 hover:bg-sky-100">
                  <ImageIcon className="h-3.5 w-3.5" aria-hidden="true" />
                  {uiLang === 'jp' ? '画像を選択' : 'Chọn ảnh chụp màn hình'}
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
                  uiLang === 'jp'
                    ? '例：\nLê Thị Bích Phương (FB: Thị Bích Phương Lê)\n栃木で介護3年目、埼玉または群馬へ転職希望。在留期限 2027/04/09、JLPT N3...'
                    : 'VD:\nLê Thị Bích Phương (FB: Thị Bích Phương Lê)\nĐang là Kaigo năm cuối ở Tochigi, muốn chuyển sang Saitama hoặc Gunma. Hạn thẻ ngoại kiều 09/04/2027, JLPT N3...'
                }
                className="w-full rounded-lg border border-slate-300 p-3 text-xs"
              />
            </div>

            <div className="mt-4 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setFbChatModalOpen(false)}
                className="rounded-lg border border-slate-300 px-4 py-2 font-semibold text-slate-700 hover:bg-slate-50 cursor-pointer"
              >
                {t.close}
              </button>
              <button
                type="button"
                disabled={isExtractingFbChatAi}
                onClick={() => void handleSubmitFbChatModal()}
                className="rounded-lg bg-[#16A34A] px-5 py-2 font-semibold text-white hover:bg-[#15803D] disabled:opacity-60 cursor-pointer"
              >
                {isExtractingFbChatAi
                  ? uiLang === 'jp'
                    ? '解析・保存中...'
                    : 'Đang phân tích & lưu dữ liệu...'
                  : uiLang === 'jp'
                  ? 'AI抽出＆原文保存'
                  : 'AI bóc tách & Lưu nguyên văn'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 3: 2-COLUMN COMPANY STANDARD RIREKISHO */}
      {rirekishoModal && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-slate-900/50 backdrop-blur-[1px] p-4">
          <div className="max-h-[92vh] w-full max-w-6xl overflow-y-auto rounded-xl border border-[#E2E8F0] bg-white p-5 text-xs text-slate-900 shadow-2xl">
            <div className="mb-3 flex flex-wrap items-center justify-between gap-2 border-b border-slate-200 pb-3">
              <div>
                <div className="flex flex-wrap items-center gap-2">
                  <h3 className="text-base font-bold text-slate-900">
                    {uiLang === 'jp'
                      ? '履歴書作成（企業提出用2カラム標準様式）'
                      : 'Tạo Rirekisho (Mẫu chuẩn 2 cột nộp công ty)'}
                    ：{getCandidateDisplayName(rirekishoModal.person, uiLang)}
                  </h3>
                  {rirekishoModal.lastSavedAt && (
                    <span className="rounded border border-emerald-200 bg-emerald-50 px-2 py-0.5 font-mono text-[10px] font-semibold text-emerald-700 tabular-nums">
                      ✅ {t.rirekisho.autoSavedLabel}{' '}
                      {new Date(rirekishoModal.lastSavedAt).toLocaleTimeString([], {
                        hour: '2-digit',
                        minute: '2-digit',
                      })}
                    </span>
                  )}
                </div>
                <p className="mt-0.5 text-[11px] text-slate-500">
                  {uiLang === 'jp'
                    ? '編集内容は自動保存され、次回はそのままPDF・Excel出力できます'
                    : 'Mọi chỉnh sửa được tự động lưu trực tiếp vào hồ sơ ứng viên để xuất PDF / Excel bất cứ lúc nào'}
                </p>
              </div>
              <div className="flex flex-wrap items-center gap-1.5">
                <button
                  type="button"
                  onClick={handleSaveRirekishoManual}
                  className="rounded-lg bg-[#16A34A] px-3.5 py-1.5 text-xs font-semibold text-white hover:bg-[#15803D] cursor-pointer whitespace-nowrap"
                >
                  💾 {t.rirekisho.saveBtn}
                </button>
                <button
                  type="button"
                  disabled={isParsingRirekisho}
                  onClick={() => rirekishoFileInputRef.current?.click()}
                  className="rounded-lg border border-amber-300 bg-amber-50 px-3 py-1.5 text-xs font-semibold text-amber-800 hover:bg-amber-100 disabled:opacity-60 cursor-pointer whitespace-nowrap"
                >
                  {isParsingRirekisho
                    ? uiLang === 'jp'
                      ? '読込・解析中...'
                      : 'Đang đọc...'
                    : rirekishoModal.hasSavedVersion
                    ? t.rirekisho.overwriteUploadBtn
                    : uiLang === 'jp'
                    ? '既存履歴書ファイル読込'
                    : 'Tải file Rirekisho có sẵn'}
                </button>
                <button
                  type="button"
                  onClick={handleExportRirekishoExcel}
                  className="rounded-lg border border-emerald-300 bg-emerald-50 px-3 py-1.5 text-xs font-semibold text-emerald-800 hover:bg-emerald-100 cursor-pointer whitespace-nowrap"
                >
                  📊 {t.rirekisho.exportExcelBtn}
                </button>
                <button
                  type="button"
                  onClick={handlePrintRirekishoSafely}
                  className="rounded-lg bg-sky-600 px-3 py-1.5 text-xs font-semibold text-white hover:bg-sky-700 cursor-pointer whitespace-nowrap"
                >
                  🖨️ {t.rirekisho.printPdfBtn}
                </button>
                <button
                  type="button"
                  onClick={handleDownloadRirekishoFile}
                  className="rounded-lg border border-[#E2E8F0] bg-white px-3 py-1.5 text-xs font-semibold text-[#0F172A] hover:bg-[#F8FAFC] cursor-pointer whitespace-nowrap"
                >
                  📥 {t.rirekisho.exportHtmlBtn}
                </button>
                <button
                  type="button"
                  onClick={handleCloseRirekishoModal}
                  className="rounded-lg bg-slate-100 px-3 py-1.5 font-semibold text-slate-700 hover:bg-slate-200 cursor-pointer"
                >
                  {t.close}
                </button>
              </div>
            </div>

            {rirekishoModal.hasSavedVersion && (
              <div className="mb-3 flex flex-wrap items-center justify-between gap-2 rounded-lg border border-emerald-200 bg-emerald-50/70 px-3.5 py-2.5">
                <div className="space-y-0.5">
                  <div className="font-bold text-emerald-900">
                    ✅ {t.rirekisho.savedBannerTitle}
                  </div>
                  <div className="text-[11px] text-emerald-800">
                    {t.rirekisho.savedBannerSub}
                    {rirekishoModal.sourceFileName && (
                      <span className="ml-2 font-mono">
                        · {t.rirekisho.sourceFileTag} {rirekishoModal.sourceFileName}
                      </span>
                    )}
                    {rirekishoModal.exportedFileName && (
                      <span className="ml-2 font-mono">
                        · {t.rirekisho.lastExportTag} {rirekishoModal.exportedFileName}
                      </span>
                    )}
                  </div>
                </div>
                <div className="flex flex-wrap items-center gap-1.5">
                  <button
                    type="button"
                    disabled={isParsingRirekisho}
                    onClick={() => rirekishoFileInputRef.current?.click()}
                    className="rounded-md border border-amber-300 bg-white px-2.5 py-1 text-[11px] font-semibold text-amber-800 hover:bg-amber-50 cursor-pointer"
                  >
                    {t.rirekisho.overwriteUploadBtn}
                  </button>
                  <button
                    type="button"
                    onClick={() =>
                      setRirekishoModal((prev) =>
                        prev ? { ...prev, showUploadTools: !prev.showUploadTools } : prev
                      )
                    }
                    className="rounded-md border border-[#E2E8F0] bg-white px-2.5 py-1 text-[11px] font-semibold text-[#475569] hover:bg-[#F8FAFC] cursor-pointer"
                  >
                    {t.rirekisho.toggleChatFillBtn}
                  </button>
                </div>
              </div>
            )}

            {(!rirekishoModal.hasSavedVersion || rirekishoModal.showUploadTools) && (
              <div className="mb-3 flex flex-wrap items-center justify-between gap-3 rounded-lg border border-amber-300 bg-amber-50/80 p-3">
                <div>
                  <div className="font-bold text-amber-950">
                    {uiLang === 'jp'
                      ? '応募者が作成した既存の履歴書ファイルから自動入力'
                      : 'Tự động điền từ file Rirekisho có sẵn của ứng viên'}
                  </div>
                  <div className="mt-0.5 text-[11px] text-amber-900">
                    {uiLang === 'jp'
                      ? '対応形式：PDF・画像（JPG/PNG）・Excel（.xlsx / .xls）・Word（.docx）・CSV・Text'
                      : 'Định dạng hỗ trợ: PDF · Ảnh (JPG/PNG) · Excel (.xlsx / .xls) · Word (.docx) · CSV · Text'}
                  </div>
                </div>
                <button
                  type="button"
                  disabled={isParsingRirekisho}
                  onClick={() => rirekishoFileInputRef.current?.click()}
                  className="rounded-lg bg-amber-600 px-4 py-2 text-xs font-semibold text-white hover:bg-amber-700 disabled:opacity-60 cursor-pointer"
                >
                  {isParsingRirekisho
                    ? uiLang === 'jp'
                      ? 'ファイル解析中...'
                      : 'Đang phân tích file...'
                    : uiLang === 'jp'
                    ? 'ファイルを選択して自動反映（PDF / Excel / Word / 画像）'
                    : 'Chọn file để điền tự động (PDF / Excel / Word / Ảnh)'}
                </button>
              </div>
            )}

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

            {(!rirekishoModal.hasSavedVersion || rirekishoModal.showUploadTools) && (
              <div className="mb-4 rounded-lg border border-indigo-200 bg-indigo-50/60 p-3">
                <div className="mb-1.5 flex flex-wrap items-center justify-between gap-2">
                  <span className="font-bold text-indigo-950">
                    {uiLang === 'jp'
                      ? '応募者のチャット・ヒアリングメモから自動入力'
                      : 'Tự động điền từ đoạn chat / ghi chú phỏng vấn ứng viên'}
                  </span>
                  <button
                    type="button"
                    disabled={isParsingRirekisho}
                    onClick={handleGenerateRirekishoFromChat}
                    className="rounded-md bg-indigo-600 px-3 py-1.5 text-[11px] font-semibold text-white hover:bg-indigo-700 disabled:opacity-60 cursor-pointer"
                  >
                    {isParsingRirekisho
                      ? uiLang === 'jp'
                        ? 'AI解析中...'
                        : 'AI đang phân tích...'
                      : uiLang === 'jp'
                      ? 'チャットから履歴書を自動作成'
                      : 'Tạo tự động từ chat'}
                  </button>
                </div>
                <textarea
                  rows={2}
                  value={rirekishoChatInput}
                  onChange={(e) => setRirekishoChatInput(e.target.value)}
                  placeholder={
                    uiLang === 'jp'
                      ? '応募者とのFBメッセンジャーやり取りや面談メモを貼り付けてボタンを押すと、生年月日・国籍・在留期限・学歴職歴・資格・志望動機などを2カラム様式へ自動入力します...'
                      : 'Dán đoạn chat Messenger hoặc ghi chú phỏng vấn với ứng viên để AI tự điền ngày sinh, quốc tịch, hạn visa, học vấn, kinh nghiệm, chứng chỉ vào mẫu 2 cột...'
                  }
                  className="w-full rounded-lg border border-indigo-200 bg-white p-2 text-xs"
                />
              </div>
            )}

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

                <div className="flex flex-col gap-2 sm:flex-row sm:items-stretch">
                  <div className="grid flex-1 grid-cols-12 gap-2 border border-slate-400 p-2">
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
                    <div className="col-span-9 flex flex-wrap items-center gap-1.5">
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
                        className="w-28 border-b border-slate-300 px-1"
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
                        className="w-8 border-b border-slate-300 text-center font-bold"
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
                        className="w-12 border-b border-slate-300 text-center"
                      />
                    </div>
                  </div>

                  {/* Portrait Photo Box (証明写真 3x4) */}
                  <div className="flex w-28 shrink-0 flex-col items-center justify-between gap-1.5 rounded border border-dashed border-slate-400 bg-slate-50 p-1.5">
                    {rirekishoModal.draft.photoDataUrl ? (
                      <img
                        src={rirekishoModal.draft.photoDataUrl}
                        alt={t.rirekisho.photoBoxTitle}
                        className="h-28 w-22 rounded border border-slate-300 object-cover"
                      />
                    ) : (
                      <button
                        type="button"
                        onClick={() => rirekishoPhotoInputRef.current?.click()}
                        className="flex h-28 w-22 flex-col items-center justify-center rounded border border-slate-200 bg-white text-[10px] font-semibold text-slate-500 hover:border-emerald-400 hover:text-emerald-700 cursor-pointer"
                      >
                        <ImageIcon className="mb-1 h-4 w-4" aria-hidden="true" />
                        <span>{t.rirekisho.photoBoxTitle}</span>
                      </button>
                    )}
                    <div className="flex items-center gap-1">
                      <button
                        type="button"
                        onClick={() => rirekishoPhotoInputRef.current?.click()}
                        className="rounded border border-slate-300 bg-white px-1.5 py-0.5 text-[10px] font-semibold text-slate-700 hover:bg-slate-100 cursor-pointer"
                      >
                        {rirekishoModal.draft.photoDataUrl
                          ? t.rirekisho.changePhotoBtn
                          : t.rirekisho.uploadPhotoBtn}
                      </button>
                      {rirekishoModal.draft.photoDataUrl && (
                        <button
                          type="button"
                          onClick={() =>
                            setRirekishoModal({
                              ...rirekishoModal,
                              draft: { ...rirekishoModal.draft, photoDataUrl: '' },
                            })
                          }
                          className="rounded border border-rose-200 bg-rose-50 px-1.5 py-0.5 text-[10px] font-semibold text-rose-700 hover:bg-rose-100 cursor-pointer"
                        >
                          {t.rirekisho.removePhotoBtn}
                        </button>
                      )}
                    </div>
                  </div>
                </div>

                <div className="grid grid-cols-12 gap-2 border border-slate-400 p-2">
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
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-slate-900/50 backdrop-blur-[1px] p-4">
          <div className="max-h-[90vh] w-full max-w-3xl overflow-y-auto rounded-xl border border-[#E2E8F0] bg-white p-5 text-xs text-[#0F172A] shadow-2xl">
            <div className="mb-3 flex items-center justify-between border-b border-[#E2E8F0] pb-2.5">
              <h3 className="text-base font-bold text-[#0F172A]">
                {uiLang === 'jp' ? '返信文提案（VN / JP / EN）' : 'Gợi ý trả lời (VN / JP / EN)'}：
                {getCandidateDisplayName(replyModal.person, uiLang)}
              </h3>
              <button
                type="button"
                onClick={() => setReplyModal(null)}
                className="rounded-lg bg-slate-100 px-3 py-1 font-semibold text-slate-700 hover:bg-slate-200 cursor-pointer"
              >
                {t.close}
              </button>
            </div>

            <div className="mb-3 flex flex-wrap gap-1.5">
              {(
                [
                  {
                    id: 'first_mendan',
                    label:
                      uiLang === 'jp'
                        ? '1. 初回挨拶・面談調整'
                        : '1. Chào & Hẹn Mendan đầu',
                  },
                  {
                    id: 'missing_info_request',
                    label: uiLang === 'jp' ? '2. 不足情報の確認' : '2. Xin thông tin thiếu',
                  },
                  {
                    id: 'recommend_job',
                    label:
                      uiLang === 'jp'
                        ? '3. 最適求人3件を提案'
                        : '3. Gửi 3 Job phù hợp nhất',
                  },
                  {
                    id: 'remind_interview',
                    label:
                      uiLang === 'jp'
                        ? '4. 面談・面接リマインド'
                        : '4. Nhắc lịch Mendan / PV',
                  },
                  {
                    id: 'post_interview_followup',
                    label:
                      uiLang === 'jp'
                        ? '5. 面接後フォロー'
                        : '5. Hỏi thăm sau phỏng vấn',
                  },
                  {
                    id: 'followup_silent',
                    label:
                      uiLang === 'jp'
                        ? '6. 未返信フォロー'
                        : '6. Chăm sóc lại / Chưa phản hồi',
                  },
                ] as Array<{ id: ReplyScenario; label: string }>
              ).map((sc) => (
                <button
                  key={sc.id}
                  type="button"
                  onClick={() => setReplyModal({ ...replyModal, scenario: sc.id })}
                  className={`rounded-lg px-3 py-1.5 font-semibold transition-colors cursor-pointer ${
                    replyModal.scenario === sc.id
                      ? 'bg-[#16A34A] text-white'
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
                    {
                      lang: uiLang === 'jp' ? 'ベトナム語 (Tiếng Việt)' : 'Tiếng Việt',
                      text: msgs.vi,
                    },
                    {
                      lang: uiLang === 'jp' ? '日本語 (Japanese)' : 'Tiếng Nhật (日本語)',
                      text: msgs.jp,
                    },
                    {
                      lang: uiLang === 'jp' ? '英語 (English)' : 'Tiếng Anh (English)',
                      text: msgs.en,
                    },
                  ].map((item) => (
                    <div
                      key={item.lang}
                      className="rounded-lg border border-[#E2E8F0] bg-[#F8FAFC] p-3"
                    >
                      <div className="mb-1.5 flex items-center justify-between">
                        <span className="font-bold text-[#0F172A]">{item.lang}</span>
                        <button
                          type="button"
                          onClick={async () => {
                            const ok = await copyTextSafely(item.text);
                            showToast(
                              ok
                                ? uiLang === 'jp'
                                  ? `📋 ${item.lang} をコピーしました`
                                  : `📋 Đã sao chép ${item.lang}`
                                : t.toasts.copyFailed
                            );
                          }}
                          className="rounded-md bg-[#0F172A] px-3 py-1 text-[11px] font-semibold text-white hover:bg-slate-700 cursor-pointer"
                        >
                          {t.copy}
                        </button>
                      </div>
                      <pre className="whitespace-pre-wrap font-sans leading-relaxed text-[#334155]">
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
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-slate-900/50 backdrop-blur-[1px] p-4">
          <div className="max-h-[90vh] w-full max-w-2xl overflow-y-auto rounded-xl border border-[#E2E8F0] bg-white p-5 text-[#0F172A] shadow-2xl">
            <div className="mb-3 flex items-center justify-between border-b border-[#E2E8F0] pb-2.5">
              <h3 className="text-base font-bold text-[#0F172A]">
                {uiLang === 'jp'
                  ? '求職者用 紹介文プレビュー（企業名・詳細住所非公開）'
                  : 'Xem trước tin tuyển dụng ẩn danh (Ẩn tên công ty & địa chỉ chi tiết)'}
              </h3>
              <button
                type="button"
                onClick={() => setPreviewModal(null)}
                className="rounded-lg bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-700 hover:bg-slate-200 cursor-pointer"
              >
                {t.close}
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
                      <span className="font-bold text-emerald-900">
                        {uiLang === 'jp'
                          ? 'ベトナム語紹介文（匿名）'
                          : 'Bản tóm tắt Tiếng Việt (Ẩn danh)'}
                      </span>
                      <button
                        type="button"
                        onClick={async () => {
                          const ok = await copyTextSafely(viIntroText);
                          showToast(
                            ok
                              ? uiLang === 'jp'
                                ? '📋 ベトナム語紹介文をコピーしました'
                                : '📋 Đã sao chép bản tóm tắt Tiếng Việt'
                              : t.toasts.copyFailed
                          );
                        }}
                        className="rounded-md bg-[#16A34A] px-3 py-1 text-[11px] font-semibold text-white hover:bg-[#15803D] cursor-pointer"
                      >
                        {uiLang === 'jp' ? 'ベトナム語コピー' : 'Sao chép Tiếng Việt'}
                      </button>
                    </div>
                    <pre className="whitespace-pre-wrap font-sans leading-relaxed text-[#334155]">
                      {viIntroText}
                    </pre>
                  </div>

                  <div className="rounded-lg border border-sky-200 bg-sky-50/50 p-3">
                    <div className="mb-1.5 flex items-center justify-between">
                      <span className="font-bold text-sky-900">
                        {uiLang === 'jp'
                          ? '日本語紹介文（匿名）'
                          : 'Bản tóm tắt Tiếng Nhật (Ẩn danh)'}
                      </span>
                      <button
                        type="button"
                        onClick={async () => {
                          const ok = await copyTextSafely(jpIntroText);
                          showToast(
                            ok
                              ? uiLang === 'jp'
                                ? '📋 日本語紹介文をコピーしました'
                                : '📋 Đã sao chép bản tóm tắt Tiếng Nhật'
                              : t.toasts.copyFailed
                          );
                        }}
                        className="rounded-md bg-sky-600 px-3 py-1 text-[11px] font-semibold text-white hover:bg-sky-700 cursor-pointer"
                      >
                        {uiLang === 'jp' ? '日本語コピー' : 'Sao chép Tiếng Nhật'}
                      </button>
                    </div>
                    <pre className="whitespace-pre-wrap font-sans leading-relaxed text-[#334155]">
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
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-slate-900/50 backdrop-blur-[1px] p-4">
          <div className="max-h-[90vh] w-full max-w-lg overflow-y-auto rounded-xl border border-[#E2E8F0] bg-white p-5 text-xs text-[#0F172A] shadow-2xl">
            <div className="mb-3 flex items-center justify-between border-b border-[#E2E8F0] pb-2">
              <h3 className="text-base font-bold text-[#0F172A]">
                {ptModal.isNew
                  ? uiLang === 'jp'
                    ? '＋ 新規PT提携先登録'
                    : '+ Thêm đối tác PT mới'
                  : uiLang === 'jp'
                  ? 'PT提携先編集'
                  : 'Chỉnh sửa đối tác PT'}
              </h3>
              <button
                type="button"
                onClick={() => setPtModal(null)}
                className="rounded-lg bg-slate-100 px-3 py-1 font-semibold text-slate-700 hover:bg-slate-200 cursor-pointer"
              >
                {t.close}
              </button>
            </div>
            <div className="space-y-2.5">
              <div>
                <label className="mb-1 block font-semibold">
                  {uiLang === 'jp' ? 'PT団体・企業名 *' : 'Tên tổ chức / Công ty PT *'}
                </label>
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
                  <label className="mb-1 block font-semibold">
                    {uiLang === 'jp' ? '担当者名' : 'Người phụ trách'}
                  </label>
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
                  <label className="mb-1 block font-semibold">
                    {uiLang === 'jp' ? '提携ステータス' : 'Trạng thái hợp tác'}
                  </label>
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
                  <label className="mb-1 block font-semibold">
                    {uiLang === 'jp' ? '電話番号' : 'Số điện thoại'}
                  </label>
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
                  <label className="mb-1 block font-semibold">Email</label>
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
                <label className="mb-1 block font-semibold">
                  {uiLang === 'jp' ? '得意分野' : 'Ngành thế mạnh'}
                </label>
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
                <label className="mb-1 block font-semibold">
                  {uiLang === 'jp' ? '契約・紹介シェア条件' : 'Điều kiện hợp tác / Hoa hồng'}
                </label>
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
                <label className="mb-1 block font-semibold">
                  {uiLang === 'jp' ? '備考メモ' : 'Ghi chú'}
                </label>
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
                className="rounded-lg border border-slate-300 px-4 py-2 font-semibold text-slate-700 hover:bg-slate-50 cursor-pointer"
              >
                {t.cancel}
              </button>
              <button
                type="button"
                onClick={() => {
                  if (!ptModal.data.partnerName.trim()) {
                    showToast(
                      uiLang === 'jp'
                        ? '⚠️ PT団体名を入力してください'
                        : '⚠️ Vui lòng nhập tên đối tác PT'
                    );
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
                  showToast(
                    uiLang === 'jp'
                      ? '✅ PT提携先を保存しました'
                      : '✅ Đã lưu thông tin đối tác PT'
                  );
                }}
                className="rounded-lg bg-[#16A34A] px-4 py-2 font-semibold text-white hover:bg-[#15803D] cursor-pointer"
              >
                {t.save}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 7: CUSTOM CONFIRMATION DIALOG */}
      {confirmDialog && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-slate-900/50 backdrop-blur-[1px] p-4">
          <div className="w-full max-w-md rounded-xl border border-[#E2E8F0] bg-white p-5 text-[#0F172A] shadow-2xl">
            <h4 className="text-base font-bold text-[#0F172A]">{confirmDialog.title}</h4>
            <p className="mt-2 whitespace-pre-line text-xs leading-relaxed text-[#64748B]">
              {confirmDialog.message}
            </p>
            <div className="mt-4 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => {
                  confirmDialog.onCancel?.();
                  setConfirmDialog(null);
                }}
                className="rounded-md border border-[#E2E8F0] bg-white px-4 py-2 text-xs font-semibold text-[#475569] hover:bg-[#F8FAFC] hover:text-[#0F172A] cursor-pointer"
              >
                {confirmDialog.cancelLabel || t.cancel}
              </button>
              <button
                type="button"
                onClick={confirmDialog.onConfirm}
                className={`rounded-md px-4 py-2 text-xs font-semibold cursor-pointer ${
                  confirmDialog.tone === 'accent'
                    ? 'bg-[#16A34A] text-white hover:bg-[#15803D]'
                    : 'bg-rose-600 text-white hover:bg-rose-700'
                }`}
              >
                {confirmDialog.confirmLabel ||
                  (uiLang === 'jp' ? '削除を実行する' : 'Xác nhận xóa')}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 8: FACEBOOK MESSENGER CONNECTION MODAL */}
      {messengerLinkModalOpen && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-slate-900/50 backdrop-blur-[1px] p-4">
          <div className="max-h-[90vh] w-full max-w-xl overflow-y-auto rounded-xl border border-[#E2E8F0] bg-white p-6 text-xs text-[#0F172A] shadow-2xl">
            <div className="mb-4 flex items-center justify-between border-b border-[#E2E8F0] pb-3">
              <div className="flex items-center gap-2.5">
                <div className="flex h-8 w-8 items-center justify-center rounded-lg border border-sky-200 bg-sky-50 text-sky-700">
                  <Link2 className="h-4 w-4" aria-hidden="true" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-[#0F172A]">{t.linkMessenger}</h3>
                  <p className="text-[11px] text-[#64748B]">
                    {uiLang === 'jp'
                      ? 'FanpageとMeta Messenger Webhookを連携して自動登録します'
                      : 'Kết nối Fanpage & Meta Messenger Webhook tự động nhận hồ sơ ứng viên'}
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setMessengerLinkModalOpen(false)}
                className="rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-1.5 text-xs font-semibold text-[#475569] hover:text-[#0F172A] cursor-pointer"
              >
                {t.close}
              </button>
            </div>

            {/* Connection Status Grid */}
            <div className="mb-4 grid grid-cols-3 gap-2">
              <div className="rounded-lg border border-[#E2E8F0] bg-[#F8FAFC] p-3">
                <div className="font-mono text-[10px] uppercase text-[#64748B]">Webhook</div>
                <div className="mt-1 font-mono text-xs font-bold text-emerald-700">
                  {t.readyBadge}
                </div>
              </div>
              <div className="rounded-lg border border-[#E2E8F0] bg-[#F8FAFC] p-3">
                <div className="font-mono text-[10px] uppercase text-[#64748B]">Verify Token</div>
                <div
                  className={`mt-1 font-mono text-xs font-bold ${
                    webhookStatus?.hasVerifyToken ? 'text-emerald-700' : 'text-sky-700'
                  }`}
                >
                  {webhookStatus?.hasVerifyToken ? 'Configured' : 'Optional'}
                </div>
              </div>
              <div className="rounded-lg border border-[#E2E8F0] bg-[#F8FAFC] p-3">
                <div className="font-mono text-[10px] uppercase text-[#64748B]">App Secret</div>
                <div
                  className={`mt-1 font-mono text-xs font-bold ${
                    webhookStatus?.hasAppSecret ? 'text-emerald-700' : 'text-sky-700'
                  }`}
                >
                  {webhookStatus?.hasAppSecret ? 'HMAC Active' : 'Optional'}
                </div>
              </div>
            </div>

            {/* Webhook Callback URL */}
            <div className="mb-4 rounded-lg border border-[#E2E8F0] bg-[#F8FAFC] p-4">
              <label className="mb-1.5 block font-mono text-[10px] font-semibold uppercase tracking-wider text-[#64748B]">
                Callback URL (Meta Developers &rarr; Messenger &rarr; Webhooks)
              </label>
              <div className="flex flex-wrap gap-2">
                <input
                  type="text"
                  readOnly
                  value={sharedMessengerWebhookUrl || liveMessengerWebhookUrl}
                  aria-label="Callback URL Messenger Webhook"
                  className="min-w-[240px] flex-1 rounded-md border border-[#E2E8F0] bg-white px-3 py-2 font-mono text-[11px] text-[#0F172A]"
                />
                <button
                  type="button"
                  onClick={() => {
                    const target = sharedMessengerWebhookUrl || liveMessengerWebhookUrl;
                    copyTextSafely(target);
                    showToast(`📋 ${target}`);
                  }}
                  className="inline-flex items-center gap-1.5 rounded-md bg-[#16A34A] px-3.5 py-2 text-xs font-semibold text-white hover:bg-[#15803D] cursor-pointer"
                >
                  <Copy className="h-3.5 w-3.5" aria-hidden="true" />
                  {t.copyUrl}
                </button>
              </div>
              {sharedMessengerWebhookUrl && (
                <div className="mt-2 rounded border border-amber-300 bg-amber-50 px-3 py-2 text-[11px] leading-relaxed text-amber-900">
                  <strong>{uiLang === 'jp' ? 'ご注意:' : 'Lưu ý quan trọng:'}</strong>{' '}
                  {uiLang === 'jp'
                    ? 'Meta Developersから呼び出すにはShare / PublishしたURL（/api/webhooks/messenger）をご利用ください。'
                    : 'Để Meta Developers gọi Webhook từ bên ngoài, hãy dùng URL sau khi Share / Publish có đuôi /api/webhooks/messenger ở trên.'}
                </div>
              )}
              <p className="mt-2 text-[11px] leading-relaxed text-[#64748B]">
                {uiLang === 'jp'
                  ? 'Webhookイベント：messages、messaging_postbacks を有効にしてください。'
                  : 'Bật 2 trường đăng ký webhook: messages và messaging_postbacks.'}
              </p>
            </div>

            {/* Fanpage Messenger Quick Link */}
            <div className="mb-2 rounded-lg border border-[#E2E8F0] bg-[#F8FAFC] p-4">
              <label className="mb-1.5 block font-mono text-[10px] font-semibold uppercase tracking-wider text-[#64748B]">
                {uiLang === 'jp'
                  ? 'Fanpage Messengerリンク (m.me / Page ID)'
                  : 'Liên kết Fanpage Messenger (m.me / Page ID / Facebook URL)'}
              </label>
              <div className="flex flex-wrap gap-2">
                <input
                  type="text"
                  value={fbPageLinkDraft}
                  onChange={(e) => setFbPageLinkDraft(e.target.value)}
                  placeholder="VD: tokuteicareer hoặc https://m.me/tokuteicareer"
                  className="min-w-[220px] flex-1 rounded-md border border-[#E2E8F0] bg-white px-3 py-2 text-xs text-[#0F172A] placeholder:text-[#94A3B8] focus:border-sky-500 focus:outline-none"
                />
                <button
                  type="button"
                  onClick={handleSaveFbPageLink}
                  className="rounded-md bg-sky-600 px-3.5 py-2 text-xs font-semibold text-white hover:bg-sky-700 cursor-pointer"
                >
                  {t.save}
                </button>
              </div>

              <div className="mt-3 flex flex-wrap gap-2">
                {fbPageLink && (
                  <a
                    href={fbPageLink}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex items-center gap-1.5 rounded-md border border-sky-200 bg-sky-50 px-3 py-1.5 text-[11px] font-semibold text-sky-700 hover:bg-sky-100"
                  >
                    <MessageCircle className="h-3.5 w-3.5" aria-hidden="true" />
                    {t.openPage}
                  </a>
                )}
                <a
                  href="https://business.facebook.com/latest/inbox/all"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="inline-flex items-center gap-1.5 rounded-md border border-[#E2E8F0] bg-white px-3 py-1.5 text-[11px] font-semibold text-[#0F172A] hover:border-sky-300"
                >
                  <ExternalLink className="h-3.5 w-3.5" aria-hidden="true" />
                  Meta Business Inbox
                </a>
                <a
                  href="https://developers.facebook.com/apps/"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="inline-flex items-center gap-1.5 rounded-md border border-[#E2E8F0] bg-white px-3 py-1.5 text-[11px] font-semibold text-[#0F172A] hover:border-emerald-300"
                >
                  <ExternalLink className="h-3.5 w-3.5" aria-hidden="true" />
                  Meta Developers
                </a>
                <button
                  type="button"
                  disabled={isTestingWebhook}
                  onClick={() => void handleTestMessengerWebhook()}
                  className="inline-flex items-center gap-1.5 rounded-md border border-sky-200 bg-sky-50 px-3 py-1.5 text-[11px] font-semibold text-sky-700 hover:bg-sky-100 disabled:opacity-50 cursor-pointer"
                >
                  <MessageCircle className="h-3.5 w-3.5" aria-hidden="true" />
                  {isTestingWebhook
                    ? uiLang === 'jp'
                      ? 'テスト中...'
                      : 'Đang thử nghiệm...'
                    : uiLang === 'jp'
                    ? '受信テスト'
                    : 'Thử nhận tin nhắn Messenger'}
                </button>
                <button
                  type="button"
                  disabled={isRefreshingSql}
                  onClick={() => void refreshCrmFromSql(false)}
                  className="inline-flex items-center gap-1.5 rounded-md border border-[#E2E8F0] bg-white px-3 py-1.5 text-[11px] font-semibold text-emerald-700 hover:border-emerald-300 disabled:opacity-50 cursor-pointer"
                >
                  <RefreshCw
                    className={`h-3.5 w-3.5 ${isRefreshingSql ? 'animate-spin' : ''}`}
                    aria-hidden="true"
                  />
                  {uiLang === 'jp' ? '新着メッセージを同期' : 'Đồng bộ tin nhắn mới'}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 9: KYUJIN RAW TEXT PASTE MODAL */}
      {kyujinRawModalOpen && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-slate-900/50 backdrop-blur-[1px] p-4">
          <div className="max-h-[90vh] w-full max-w-2xl overflow-y-auto rounded-xl border border-[#E2E8F0] bg-white p-5 text-xs text-[#0F172A] shadow-2xl">
            <div className="mb-3 flex items-center justify-between border-b border-[#E2E8F0] pb-2.5">
              <div className="flex items-center gap-2">
                <FileText className="h-4 w-4 text-[#16A34A]" aria-hidden="true" />
                <h3 className="text-base font-bold text-[#0F172A]">
                  {t.kyujin.pasteTextBtn}
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setKyujinRawModalOpen(false)}
                className="rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-1 text-xs font-semibold text-[#475569] hover:text-[#0F172A] cursor-pointer"
              >
                {t.close}
              </button>
            </div>
            <p className="mb-2 text-[11px] text-[#64748B]">
              {uiLang === 'jp'
                ? '求人情報テキストを貼り付けてください（複数件は空行または---で区切り）。業種・雇用形態・給与・経験要件を自動抽出します。'
                : 'Dán nội dung tuyển dụng (tiếng Nhật hoặc tiếng Việt, hỗ trợ nhiều đơn cách nhau bằng dòng trống hoặc ---). Hệ thống tự chia Lĩnh vực, Ngành nghề, Hình thức, lương và yêu cầu kinh nghiệm.'}
            </p>
            <textarea
              rows={9}
              value={kyujinRawText}
              onChange={(e) => setKyujinRawText(e.target.value)}
              placeholder={
                'VD:\nCông ty ABC - Tuyển Baito Konbini tại Tokyo (Shinjuku)\n時給: 1,250円 / 交通費支給 / 週3〜5日OK / N3 / 未経験OK...'
              }
              className="w-full rounded-lg border border-[#E2E8F0] bg-[#F8FAFC] p-3 text-xs text-[#0F172A] placeholder:text-[#94A3B8] focus:border-[#16A34A] focus:bg-white focus:outline-none"
            />
            <div className="mt-4 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setKyujinRawModalOpen(false)}
                className="rounded-md border border-[#E2E8F0] px-4 py-2 font-semibold text-[#64748B] hover:text-[#0F172A] cursor-pointer"
              >
                {t.cancel}
              </button>
              <button
                type="button"
                disabled={isImportingKyujin}
                onClick={() => void handleImportKyujinRawText()}
                className="rounded-md bg-[#16A34A] px-4 py-2 font-semibold text-white hover:bg-[#15803D] disabled:opacity-60 cursor-pointer"
              >
                {isImportingKyujin
                  ? t.kyujin.analyzing
                  : uiLang === 'jp'
                  ? '抽出して求人登録'
                  : 'Bóc tách & Nhập Kyujin'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 10: RESIGNATION PAPERWORK REMINDER MODAL (FROM RIGHT SIDEBAR) */}
      {resignationReminderModal && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-slate-900/50 backdrop-blur-[1px] p-4">
          <div className="max-h-[90vh] w-full max-w-2xl overflow-y-auto rounded-xl border border-[#E2E8F0] bg-white p-5 text-xs text-[#0F172A] shadow-2xl">
            <div className="mb-3 flex items-center justify-between border-b border-[#E2E8F0] pb-2.5">
              <h3 className="text-base font-bold text-[#0F172A]">
                {uiLang === 'jp'
                  ? '退職者書類提出リマインド文'
                  : 'Nhắn nhắc nộp giấy tờ nghỉ việc'}
                : {getResignationDisplayName(resignationReminderModal, uiLang)}
              </h3>
              <button
                type="button"
                onClick={() => setResignationReminderModal(null)}
                className="rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-1 text-xs font-semibold text-[#475569] hover:text-[#0F172A] cursor-pointer"
              >
                {t.close}
              </button>
            </div>
            {(() => {
              const msgs = buildResignationReminderMessages(resignationReminderModal);
              return (
                <div className="space-y-3">
                  {[
                    { label: 'Tiếng Việt', text: msgs.vi },
                    { label: '日本語 (Tiếng Nhật)', text: msgs.jp },
                  ].map((item) => (
                    <div
                      key={item.label}
                      className="rounded-lg border border-[#E2E8F0] bg-[#F8FAFC] p-3"
                    >
                      <div className="mb-1.5 flex items-center justify-between">
                        <span className="font-semibold text-emerald-700">{item.label}</span>
                        <button
                          type="button"
                          onClick={async () => {
                            const ok = await copyTextSafely(item.text);
                            showToast(ok ? `📋 ${item.label}` : t.toasts.copyFailed);
                          }}
                          className="rounded bg-[#16A34A] px-3 py-1 text-[11px] font-semibold text-white hover:bg-[#15803D] cursor-pointer"
                        >
                          {t.copy}
                        </button>
                      </div>
                      <pre className="whitespace-pre-wrap font-sans leading-relaxed text-[#334155]">
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

      {/* MODAL 11: 1-CLICK GOOGLE CALENDAR TEMPLATE LINK MODAL (ZERO OAUTH / 403) */}
      {resignationCalendarModal && (
        <div className="fixed inset-0 z-[9999] flex items-center justify-center bg-slate-900/50 backdrop-blur-[1px] p-4">
          <div className="max-h-[90vh] w-full max-w-2xl overflow-y-auto rounded-xl border border-[#E2E8F0] bg-white p-5 text-xs text-[#0F172A] shadow-2xl">
            <div className="mb-3 flex items-center justify-between border-b border-[#E2E8F0] pb-2.5">
              <div className="flex items-center gap-2">
                <Calendar className="h-4 w-4 text-sky-600" aria-hidden="true" />
                <h3 className="text-base font-bold text-[#0F172A]">
                  {resignationCalendarModal.mode === 'bulk'
                    ? uiLang === 'jp'
                      ? `${resignationCalendarModal.items.length}件の書類提出期限をGoogleカレンダーへ保存`
                      : `Lưu ${resignationCalendarModal.items.length} hạn giấy tờ lên Google Calendar`
                    : uiLang === 'jp'
                    ? '書類提出期限をGoogleカレンダーへ保存'
                    : 'Lưu hạn nộp giấy tờ lên Google Calendar'}
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setResignationCalendarModal(null)}
                className="rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-1 text-xs font-semibold text-[#475569] hover:text-[#0F172A] cursor-pointer"
              >
                {t.close}
              </button>
            </div>

            <p className="mb-3 text-[11px] leading-relaxed text-[#64748B]">
              {uiLang === 'jp'
                ? 'タイトルと退職日3日前の提出期限（Asia/Tokyo）が入力済みです。下のボタンから直接Googleカレンダーへ登録できます。'
                : 'Đã điền sẵn tiêu đề và ngày Hạn nộp giấy tờ (trước ngày nghỉ việc 3 ngày) theo múi giờ Asia/Tokyo. Bấm nút bên dưới để lưu trực tiếp vào Google Calendar.'}
            </p>

            <div className="space-y-2.5">
              {resignationCalendarModal.items.map(({ record, draft, templateUrl }) => (
                <div
                  key={record.id}
                  className="flex flex-col gap-2 rounded-lg border border-[#E2E8F0] bg-[#F8FAFC] p-3 sm:flex-row sm:items-center sm:justify-between"
                >
                  <div className="min-w-0 space-y-0.5">
                    <div className="font-semibold text-[#0F172A]">{draft.eventTitle}</div>
                    <div className="text-[11px] text-amber-800">
                      {draft.dateIso} (09:00–10:00 JST) ·{' '}
                      {getResignationCompanyLabel(record.previousCompany, uiLang)}
                    </div>
                  </div>
                  <a
                    href={templateUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="inline-flex shrink-0 items-center justify-center gap-1.5 rounded-md bg-[#16A34A] px-3.5 py-2 text-[11px] font-semibold text-white hover:bg-[#15803D] whitespace-nowrap"
                  >
                    <ExternalLink className="h-3.5 w-3.5" aria-hidden="true" />
                    {t.syncCalendarBtn}
                  </a>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* Prominent Corner Reminder Toast Banner for Urgent Resignation Deadlines (<= 3 days / overdue) */}
      {urgentResignationAlerts.length > 0 &&
        dismissedUrgentBannerSig !== urgentResignationSignature && (
          <div className="fixed bottom-20 right-5 z-40 w-[calc(100vw-2.5rem)] max-w-sm rounded-xl border border-rose-200 bg-white p-3.5 text-xs text-[#0F172A] shadow-xl print:hidden">
            <div className="mb-2 flex items-start justify-between gap-2">
              <div className="flex items-center gap-1.5 font-bold text-rose-700">
                <AlertTriangle className="h-4 w-4 shrink-0 text-rose-600" aria-hidden="true" />
                <span>{t.cornerResigTitle(urgentResignationAlerts.length)}</span>
              </div>
              <button
                type="button"
                aria-label={t.close}
                onClick={() => setDismissedUrgentBannerSig(urgentResignationSignature)}
                className="rounded p-0.5 text-[#64748B] hover:bg-slate-100 hover:text-[#0F172A] cursor-pointer"
              >
                <X className="h-3.5 w-3.5" aria-hidden="true" />
              </button>
            </div>
            <div className="space-y-1.5">
              {urgentResignationAlerts.slice(0, 3).map(({ record, assessment }) => (
                <div
                  key={`toast_resig_${record.id}`}
                  className="rounded border border-rose-200 bg-rose-50 px-2.5 py-1.5"
                >
                  <div className="font-semibold text-[#0F172A]">
                    {getResignationDisplayName(record, uiLang)}
                  </div>
                  <div className="text-[11px] text-rose-700">{assessment.label}</div>
                </div>
              ))}
            </div>
            <div className="mt-2.5 flex flex-wrap items-center gap-1.5">
              <button
                type="button"
                onClick={() => {
                  setViewMode('workspace');
                  setActiveTab('resignation');
                }}
                className="flex-1 rounded-md bg-rose-600 px-2.5 py-1.5 text-[11px] font-semibold text-white hover:bg-rose-700 cursor-pointer"
              >
                {t.viewAndHandleNow}
              </button>
              {notificationPermission === 'default' && (
                <button
                  type="button"
                  onClick={() => void handleRequestBrowserNotifications()}
                  className="inline-flex items-center gap-1 rounded-md border border-amber-300 bg-amber-50 px-2.5 py-1.5 text-[11px] font-semibold text-amber-800 hover:bg-amber-100 cursor-pointer"
                >
                  <Bell className="h-3 w-3" aria-hidden="true" />
                  {t.enableNotify}
                </button>
              )}
            </div>
          </div>
        )}

      {/* Floating Chatbot Launcher Button */}
      {!chatbotOpen && (
        <button
          type="button"
          onClick={() => setChatbotOpen(true)}
          aria-label={t.aiAssistant}
          className="fixed bottom-5 right-5 z-40 inline-flex items-center gap-2 rounded-full bg-[#16A34A] px-4 py-3 text-xs font-semibold text-white shadow-lg transition-transform hover:scale-[1.02] hover:bg-[#15803D] cursor-pointer print:hidden"
        >
          <MessageSquareText className="h-4 w-4" aria-hidden="true" />
          {t.aiAssistant}
        </button>
      )}

      {/* RIGHT SIDE DRAWER FOR CANDIDATE DETAILS */}
      {(() => {
        const drawerPerson = drawerCandidateId
          ? allPeopleCombined.find((p) => p.id === drawerCandidateId) || null
          : null;
        if (!drawerPerson) return null;
        const drawerDisplayName = getCandidateDisplayName(drawerPerson, uiLang);
        return (
          <div className="fixed inset-0 z-50 flex justify-end bg-slate-900/35 backdrop-blur-[1px]">
            <div
              className="flex-1"
              onClick={() => setDrawerCandidateId(null)}
              aria-hidden="true"
            />
            <aside className="relative flex h-full w-full max-w-3xl flex-col border-l border-[#E2E8F0] bg-[#F8FAFC] shadow-2xl">
              <div className="flex items-center justify-between gap-3 border-b border-[#E2E8F0] bg-white px-5 py-3.5">
                <div className="min-w-0">
                  <div className="font-mono text-[10px] font-semibold uppercase tracking-wider text-emerald-700">
                    {t.drawer.title}
                  </div>
                  <h3 className="truncate text-base font-bold text-[#0F172A]">
                    {drawerDisplayName}
                  </h3>
                </div>
                <button
                  type="button"
                  onClick={() => setDrawerCandidateId(null)}
                  className="inline-flex items-center gap-1 rounded-md border border-[#E2E8F0] bg-white px-3 py-1.5 text-xs font-semibold text-[#475569] hover:bg-[#F8FAFC] hover:text-[#0F172A] cursor-pointer"
                >
                  <X className="h-3.5 w-3.5" aria-hidden="true" />
                  {t.close}
                </button>
              </div>
              <div className="flex-1 overflow-y-auto p-4 sm:p-5">
                {renderCandidateCard(drawerPerson)}
              </div>
            </aside>
          </div>
        );
      })()}

      {/* MODAL: NAITEI / VISA SCHEDULE PROMPT */}
      {naiteiPromptModal && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center bg-slate-900/45 p-4 backdrop-blur-[1px]">
          <div className="w-full max-w-md rounded-xl border border-[#E2E8F0] bg-white p-5 shadow-xl">
            <h3 className="text-base font-bold text-[#0F172A]">{t.naiteiModal.title}</h3>
            <p className="mt-1 text-xs leading-relaxed text-[#64748B]">
              {naiteiPromptModal.candidateName} ({naiteiPromptModal.targetStatus}) —{' '}
              {t.naiteiModal.subtitle}
            </p>
            <div className="mt-4 space-y-3 text-xs">
              <div>
                <label className="mb-1 block font-semibold text-[#0F172A]">
                  {t.naiteiModal.companyLabel}
                </label>
                <input
                  type="text"
                  placeholder={t.naiteiModal.companyPlaceholder}
                  value={naiteiPromptModal.naiteiCompany}
                  onChange={(e) =>
                    setNaiteiPromptModal({
                      ...naiteiPromptModal,
                      naiteiCompany: e.target.value,
                    })
                  }
                  className="w-full rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-2 text-[#0F172A] focus:border-[#16A34A] focus:bg-white focus:outline-none"
                />
              </div>
              <div>
                <label className="mb-1 block font-semibold text-[#0F172A]">
                  {t.naiteiModal.joinDateLabel}
                </label>
                <input
                  type="date"
                  value={naiteiPromptModal.expectedJoinDate}
                  onChange={(e) =>
                    setNaiteiPromptModal({
                      ...naiteiPromptModal,
                      expectedJoinDate: e.target.value,
                    })
                  }
                  className="w-full rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-2 text-[#0F172A] tabular-nums focus:border-[#16A34A] focus:bg-white focus:outline-none"
                />
              </div>
              <div>
                <label className="mb-1 block font-semibold text-[#0F172A]">
                  {t.naiteiModal.visaDeadlineLabel}
                </label>
                <input
                  type="date"
                  value={naiteiPromptModal.visaSubmissionDeadline}
                  onChange={(e) =>
                    setNaiteiPromptModal({
                      ...naiteiPromptModal,
                      visaSubmissionDeadline: e.target.value,
                    })
                  }
                  className="w-full rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-2 text-[#0F172A] tabular-nums focus:border-[#16A34A] focus:bg-white focus:outline-none"
                />
              </div>
            </div>
            <div className="mt-5 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setNaiteiPromptModal(null)}
                className="rounded-md border border-[#E2E8F0] bg-white px-3.5 py-2 text-xs font-semibold text-[#475569] hover:bg-[#F8FAFC] cursor-pointer"
              >
                {t.cancel}
              </button>
              <button
                type="button"
                onClick={handleConfirmNaiteiPrompt}
                className="rounded-md bg-[#16A34A] px-4 py-2 text-xs font-semibold text-white hover:bg-[#15803D] cursor-pointer"
              >
                {t.naiteiModal.confirmBtn}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: JITAI / FAIL REASON PROMPT */}
      {jitaiPromptModal && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center bg-slate-900/45 p-4 backdrop-blur-[1px]">
          <div className="w-full max-w-md rounded-xl border border-[#E2E8F0] bg-white p-5 shadow-xl">
            <h3 className="text-base font-bold text-[#0F172A]">{t.jitaiModal.title}</h3>
            <p className="mt-1 text-xs leading-relaxed text-[#64748B]">
              {jitaiPromptModal.candidateName} ({jitaiPromptModal.targetStatus}) —{' '}
              {t.jitaiModal.subtitle}
            </p>
            <div className="mt-3 flex flex-wrap gap-1.5">
              {t.jitaiModal.quickReasons.map((quickText) => (
                <button
                  key={quickText}
                  type="button"
                  onClick={() =>
                    setJitaiPromptModal({
                      ...jitaiPromptModal,
                      reason: quickText,
                    })
                  }
                  className="rounded border border-[#E2E8F0] bg-[#F8FAFC] px-2 py-1 text-[11px] font-medium text-[#334155] hover:border-rose-300 hover:bg-rose-50 hover:text-rose-700 cursor-pointer"
                >
                  {quickText}
                </button>
              ))}
            </div>
            <div className="mt-3 text-xs">
              <label className="mb-1 block font-semibold text-[#0F172A]">
                {t.jitaiModal.reasonLabel}
              </label>
              <textarea
                rows={3}
                placeholder={t.jitaiModal.placeholder}
                value={jitaiPromptModal.reason}
                onChange={(e) =>
                  setJitaiPromptModal({
                    ...jitaiPromptModal,
                    reason: e.target.value,
                  })
                }
                className="w-full rounded-md border border-[#E2E8F0] bg-[#F8FAFC] p-2.5 text-[#0F172A] focus:border-rose-400 focus:bg-white focus:outline-none"
              />
            </div>
            <div className="mt-5 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setJitaiPromptModal(null)}
                className="rounded-md border border-[#E2E8F0] bg-white px-3.5 py-2 text-xs font-semibold text-[#475569] hover:bg-[#F8FAFC] cursor-pointer"
              >
                {t.cancel}
              </button>
              <button
                type="button"
                onClick={handleConfirmJitaiPrompt}
                className="rounded-md bg-rose-600 px-4 py-2 text-xs font-semibold text-white hover:bg-rose-700 cursor-pointer"
              >
                {t.jitaiModal.confirmBtn}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: KYUJIN CONDITION & NG EDITOR */}
      {kyujinConditionModal && (
        <div className="fixed inset-0 z-[60] flex items-center justify-center bg-slate-900/45 p-4 backdrop-blur-[1px]">
          <div className="w-full max-w-md rounded-xl border border-[#E2E8F0] bg-white p-5 shadow-xl">
            <h3 className="text-base font-bold text-[#0F172A]">
              {kyujinConditionModal.companyName} — {t.kyujinConds.editNgBtn}
            </h3>
            <div className="mt-4 space-y-3 text-xs">
              <div>
                <label className="mb-1 block font-semibold text-[#0F172A]">
                  {uiLang === 'jp'
                    ? 'NG国籍（カンマ区切り・例: ネパール, ミャンマー）'
                    : 'Quốc tịch NG (cách nhau bằng dấu phẩy, VD: ネパール, ミャンマー)'}
                </label>
                <input
                  type="text"
                  placeholder="ネパール, インドネシア"
                  value={kyujinConditionModal.excludedNationalitiesText}
                  onChange={(e) =>
                    setKyujinConditionModal({
                      ...kyujinConditionModal,
                      excludedNationalitiesText: e.target.value,
                    })
                  }
                  className="w-full rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-2 text-[#0F172A] focus:border-[#16A34A] focus:bg-white focus:outline-none"
                />
              </div>

              <div>
                <label className="mb-1 block font-semibold text-[#0F172A]">
                  {uiLang === 'jp' ? 'NG性別' : 'Giới tính NG'}
                </label>
                <div className="flex flex-wrap gap-3">
                  {(['男性', '女性'] as const).map((g) => {
                    const checked = kyujinConditionModal.excludedGenders.includes(g);
                    return (
                      <label key={g} className="inline-flex items-center gap-1.5 cursor-pointer">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={(e) => {
                            const next = e.target.checked
                              ? Array.from(new Set([...kyujinConditionModal.excludedGenders, g]))
                              : kyujinConditionModal.excludedGenders.filter((x) => x !== g);
                            setKyujinConditionModal({
                              ...kyujinConditionModal,
                              excludedGenders: next,
                            });
                          }}
                        />
                        <span className="font-medium text-[#0F172A]">{g}NG</span>
                      </label>
                    );
                  })}
                </div>
              </div>

              <div className="flex flex-wrap gap-4 pt-1">
                <label className="inline-flex items-center gap-1.5 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={kyujinConditionModal.requiresInJapan}
                    onChange={(e) =>
                      setKyujinConditionModal({
                        ...kyujinConditionModal,
                        requiresInJapan: e.target.checked,
                      })
                    }
                  />
                  <span className="font-medium text-[#0F172A]">
                    {uiLang === 'jp'
                      ? '日本国内在住限定（海外NG）'
                      : 'Chỉ tuyển trong nước Nhật (海外NG)'}
                  </span>
                </label>
                <label className="inline-flex items-center gap-1.5 cursor-pointer">
                  <input
                    type="checkbox"
                    checked={kyujinConditionModal.requiresTokutei}
                    onChange={(e) =>
                      setKyujinConditionModal({
                        ...kyujinConditionModal,
                        requiresTokutei: e.target.checked,
                      })
                    }
                  />
                  <span className="font-medium text-[#0F172A]">
                    {uiLang === 'jp' ? '特定技能必須' : 'Yêu cầu Tokutei (特定技能)'}
                  </span>
                </label>
              </div>

              <div>
                <label className="mb-1 block font-semibold text-[#0F172A]">
                  {t.kyujinConds.requiredLabel}
                </label>
                <input
                  type="text"
                  value={kyujinConditionModal.extraRequiredText}
                  onChange={(e) =>
                    setKyujinConditionModal({
                      ...kyujinConditionModal,
                      extraRequiredText: e.target.value,
                    })
                  }
                  className="w-full rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-2 text-[#0F172A] focus:border-[#16A34A] focus:bg-white focus:outline-none"
                />
              </div>

              <div>
                <label className="mb-1 block font-semibold text-[#0F172A]">
                  {t.kyujinConds.ngLabel}
                </label>
                <input
                  type="text"
                  value={kyujinConditionModal.extraNgText}
                  onChange={(e) =>
                    setKyujinConditionModal({
                      ...kyujinConditionModal,
                      extraNgText: e.target.value,
                    })
                  }
                  className="w-full rounded-md border border-[#E2E8F0] bg-[#F8FAFC] px-3 py-2 text-[#0F172A] focus:border-rose-400 focus:bg-white focus:outline-none"
                />
              </div>
            </div>
            <div className="mt-5 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setKyujinConditionModal(null)}
                className="rounded-md border border-[#E2E8F0] bg-white px-3.5 py-2 text-xs font-semibold text-[#475569] hover:bg-[#F8FAFC] cursor-pointer"
              >
                {t.cancel}
              </button>
              <button
                type="button"
                onClick={handleSaveKyujinConditions}
                className="rounded-md bg-[#16A34A] px-4 py-2 text-xs font-semibold text-white hover:bg-[#15803D] cursor-pointer"
              >
                {t.save}
              </button>
            </div>
          </div>
        </div>
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
        lang={uiLang}
      />
    </div>
  );
}
