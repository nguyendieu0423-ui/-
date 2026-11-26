import React, { useMemo, useRef, useState } from 'react';
import {
  AlertTriangle,
  Bell,
  Calendar,
  CalendarCheck,
  CheckCircle2,
  Copy,
  Edit3,
  ExternalLink,
  FileSpreadsheet,
  Loader2,
  Lock,
  MessageSquare,
  Plus,
  Search,
  Sparkles,
  Trash2,
  UserPlus,
  X,
} from 'lucide-react';
import type { ResignationRecord } from '../types/crm';
import {
  INDUSTRIES,
  JLPT_LEVELS,
  RESIGNATION_NATIONALITIES,
  RESIGNATION_STATUSES,
} from '../data/seedData';
import {
  applyManualProtectionOnResignationSave,
  assessResignationDeadline,
  buildResignationReminderMessages,
  extractResignationDateFromText,
  normalizeAndAutoComputeResignationDates,
  normalizeResignationNationality,
  parseResignationSpreadsheetFile,
  subtractDaysFromYmd,
  suggestNationalityFromName,
  type ResignationImportPreviewRow,
} from '../utils/resignationHelpers';
import { copyTextSafely } from '../utils/crmHelpers';
import { isValidGoogleCalendarUrl } from '../utils/calendarHelpers';
import { getResignationDisplayName, type UiLang } from '../i18n/crmTranslations';

interface ResignationTabProps {
  lang?: UiLang;
  records: ResignationRecord[];
  onSaveRecord: (record: ResignationRecord, isNew: boolean) => Promise<void> | void;
  onDeleteRecord: (record: ResignationRecord) => Promise<void> | void;
  onStatusChange: (record: ResignationRecord, nextStatus: string) => Promise<void> | void;
  onBulkImport: (rows: ResignationRecord[]) => Promise<void> | void;
  onSyncCalendarSingle: (record: ResignationRecord) => Promise<void> | void;
  onSyncCalendarBulk: () => Promise<void> | void;
  onConvertToCandidate: (record: ResignationRecord) => void;
  syncingCalendarIds: Record<string, boolean>;
  isBulkCalendarSyncing: boolean;
  onNotify: (msg: string, type?: 'success' | 'error' | 'info') => void;
  hasGasUrl?: boolean;
  notificationPermission?: NotificationPermission | 'unsupported';
  onRequestBrowserNotifications?: () => void;
}

const EMPTY_FORM: Omit<ResignationRecord, 'id'> = {
  fullName: '',
  nationality: '国籍不明',
  previousCompany: '',
  resignationDate: '',
  paperworkDeadline: '',
  paperworkStatus: 'Chưa làm',
  handoverNotes: '',
  newJobNeed: '',
  phone: '',
  fbName: '',
  industry: '',
  jlpt: 'N3',
};

const RESIG_COPY = {
  vi: {
    title: 'Quản lý Người nghỉ việc & Thủ tục Giấy tờ',
    subtitleGas:
      'Tự động tính hạn giấy tờ = Ngày nghỉ việc - 3 ngày, cảnh báo nội bộ khi ≤ 3 ngày & đồng bộ Google Calendar (ưu tiên GAS CalendarApp).',
    subtitleLink:
      'Tự động tính hạn giấy tờ = Ngày nghỉ việc - 3 ngày, cảnh báo nội bộ khi ≤ 3 ngày & đồng bộ Google Calendar (tạo lịch 1-click).',
    enableNotifyBtn: 'Bật nhắc lịch trình duyệt',
    importExcelBtn: 'Nhập Excel / CSV hàng loạt',
    syncBulkCalBtn: 'Đồng bộ toàn bộ Lịch Google',
    addBtn: 'Thêm người nghỉ việc',
    urgentBanner: (n: number) =>
      `Cảnh báo tự động: Có ${n} hồ sơ sắp đến hạn nộp giấy tờ (≤ 3 ngày) hoặc đã quá hạn — đã ghim lên đầu Việc cần làm hôm nay!`,
    ruleTag: 'Hạn = Nghỉ việc - 3 ngày',
    kpiTotal: 'Tổng hồ sơ',
    kpiUrgent: 'Gấp ≤ 3 ngày / Quá hạn',
    kpiWarn: 'Sắp đến hạn (≤ 7 ngày)',
    kpiDone: 'Đã hoàn tất giấy tờ',
    searchPlaceholder: 'Tìm theo Họ tên, Công ty cũ, Quốc tịch, Ghi chú bàn giao...',
    allNat: 'Tất cả quốc tịch',
    needsNatReview: (n: number) => `⚠️ Cần rà soát quốc tịch (${n})`,
    unknownNatShort: 'Chưa rõ quốc tịch',
    unknownNatOpt: 'Chưa rõ quốc tịch',
    allStatus: (n: number) => `Tất cả trạng thái (${n})`,
    emptyMsg:
      'Chưa có hồ sơ nghỉ việc nào khớp bộ lọc. Nhấn "Thêm người nghỉ việc" hoặc "Nhập Excel / CSV hàng loạt".',
    natAria: (name: string) => `Quốc tịch của ${name}`,
    suggestNatTitle: 'Nhận diện theo tên — bấm để sửa nhanh và khóa bảo vệ GAS',
    suggestNatBtn: (nat: string) => `Gợi ý: ${nat} (Sửa nhanh)`,
    gasProtectedTitle: (fields: string) =>
      `Trường đã chỉnh thủ công không bị GAS ghi đè: ${fields}`,
    gasProtectedBadge: 'Bảo vệ GAS',
    prevCompanyLabel: 'Công ty cũ:',
    resigDateLabel: 'Ngày nghỉ việc:',
    deadlineLabel: 'Hạn giấy tờ:',
    contactIndLabel: 'Liên hệ / Ngành:',
    unset: 'Chưa cập nhật',
    noDeadline: 'Chưa đặt hạn',
    handoverLabel: 'Ghi chú bàn giao:',
    newJobLabel: 'Nhu cầu chuyển việc mới:',
    statusPrefix: 'Trạng thái:',
    updateCalBtn: 'Cập nhật Lịch GG',
    syncCalBtn: 'Đồng bộ Lịch Google',
    openCalTitle: 'Mở sự kiện trên Google Calendar',
    createReminderBtn: 'Tạo tin nhắn nhắc nộp giấy tờ',
    convertedBtn: 'Đã chuyển (Tạo lại UV mới)',
    convertBtn: 'Chuyển sang Ứng viên tìm việc mới',
    editBtn: 'Sửa',
    deleteBtn: 'Xóa',
    modalEditTitle: 'Cập nhật hồ sơ nghỉ việc',
    modalAddTitle: 'Thêm hồ sơ Người nghỉ việc',
    nameLabel: 'Họ và tên *',
    namePlaceholder: 'Ví dụ: LE THI BICH PHUONG',
    natLabel: 'Quốc tịch',
    suggestByNameBtn: (nat: string) => `Gợi ý theo tên: chọn ${nat}`,
    prevCompanyInputLabel: 'Công ty cũ đang làm việc *',
    prevCompanyPlaceholder: 'Tên công ty / cơ sở sắp nghỉ việc',
    resigDateInputLabel: 'Ngày nghỉ việc chính thức',
    deadlineInputLabel: 'Hạn chót làm & nộp giấy tờ',
    calcMinus3Btn: 'Tính = Nghỉ việc - 3 ngày',
    autoDeadlineHint: (ymd: string) => `Mặc định tự động: ${ymd} (trước ngày nghỉ việc 3 ngày)`,
    paperworkStatusLabel: 'Trạng thái giấy tờ',
    indJlptLabel: 'Ngành / JLPT (phục vụ chuyển UV)',
    selectIndOpt: '-- Chọn ngành --',
    phoneLabel: 'Số điện thoại',
    fbNameLabel: 'Tên Facebook',
    fbNamePlaceholder: 'Tên hiển thị Messenger',
    handoverNotesLabel: 'Ghi chú bàn giao & giấy tờ',
    detectedDateBtn: (d: string, dl: string) =>
      `Tách được ngày nghỉ ${d} (Hạn: ${dl}) → Áp dụng`,
    handoverPlaceholder:
      'Ví dụ: Dự kiến nghỉ việc 11月30日 / 12月末. Đang chờ công ty cũ gửi 退職証明書, 源泉徴収票...',
    newJobNeedLabel: 'Nhu cầu chuyển việc mới',
    newJobPlaceholder: 'Ví dụ: Tìm đơn Kaigo lương thực nhận >= 17 man tại Saitama / Tochigi...',
    cancelBtn: 'Hủy',
    saveSyncBtn: 'Lưu & Đồng bộ nhắc lịch',
    importTitle: (sheet: string) => `Xem trước dữ liệu nhập từ Excel / CSV (${sheet})`,
    importDesc: (n: number) =>
      `Đã tự động ánh xạ ${n} dòng. Kiểm tra trước khi bấm Xác nhận nhập hàng loạt.`,
    thName: 'Họ tên',
    thNat: 'Quốc tịch',
    thPrevComp: 'Công ty cũ',
    thResigDate: 'Ngày nghỉ',
    thDeadline: 'Hạn giấy tờ',
    thStatus: 'Trạng thái',
    thNotes: 'Ghi chú / Nhu cầu',
    closeBtn: 'Đóng',
    confirmImportBtn: (n: number) => `Xác nhận nhập ${n} hồ sơ`,
    reminderTitle: (name: string) => `Tin nhắn nhắc nộp giấy tờ nghỉ việc — ${name}`,
    viLabel: 'Tiếng Việt',
    jpLabel: 'Tiếng Nhật (日本語)',
    copyBtn: 'Sao chép',
    notifyMissingFields: 'Vui lòng nhập đầy đủ Họ tên và Công ty cũ.',
    notifyNoValidRows: 'Không tìm thấy dòng dữ liệu hợp lệ trong file Excel/CSV.',
    notifyImportErr: (msg: string) => `Lỗi đọc file Excel/CSV: ${msg}`,
    notifyCopiedVi: 'Đã sao chép tin nhắn nhắc giấy tờ (Tiếng Việt)!',
    notifyCopiedJp: 'Đã sao chép tin nhắn nhắc giấy tờ (Tiếng Nhật)!',
  },
  jp: {
    title: '退職者管理・書類提出期限フォロー',
    subtitleGas:
      '提出期限＝退職日−3日を自動算出し、3日以内・超過を内部アラート表示＆Googleカレンダー同期（GAS CalendarApp優先）。',
    subtitleLink:
      '提出期限＝退職日−3日を自動算出し、3日以内・超過を内部アラート表示＆Googleカレンダー登録（1クリック）。',
    enableNotifyBtn: 'ブラウザ通知を有効化',
    importExcelBtn: 'Excel / CSV一括読込',
    syncBulkCalBtn: 'Googleカレンダー一括同期',
    addBtn: '＋ 退職者を追加',
    urgentBanner: (n: number) =>
      `自動アラート：提出期限まで3日以内または期限超過の案件が ${n} 件あります（本日のタスク最上部に固定表示中）`,
    ruleTag: '提出期限 = 退職日 - 3日',
    kpiTotal: '登録総数',
    kpiUrgent: '至急（3日以内・超過）',
    kpiWarn: '期限間近（7日以内）',
    kpiDone: '書類手続完了',
    searchPlaceholder: '氏名・前職企業名・国籍・引継ぎメモで検索...',
    allNat: 'すべての国籍',
    needsNatReview: (n: number) => `⚠️ 国籍要確認 (${n})`,
    unknownNatShort: '国籍不明',
    unknownNatOpt: '国籍不明',
    allStatus: (n: number) => `すべての進捗 (${n})`,
    emptyMsg:
      '条件に一致する退職者データがありません。「＋ 退職者を追加」または「Excel / CSV一括読込」をご利用ください。',
    natAria: (name: string) => `${name} の国籍`,
    suggestNatTitle: '氏名から国籍を推定 — クリックで修正＆GAS上書き保護',
    suggestNatBtn: (nat: string) => `候補: ${nat} (ワンクリック修正)`,
    gasProtectedTitle: (fields: string) => `手動修正済み（GAS上書き保護）: ${fields}`,
    gasProtectedBadge: 'GAS保護中',
    prevCompanyLabel: '前職企業:',
    resigDateLabel: '退職日:',
    deadlineLabel: '書類提出期限:',
    contactIndLabel: '連絡先・分野:',
    unset: '未設定',
    noDeadline: '期限未設定',
    handoverLabel: '引継ぎ・書類メモ:',
    newJobLabel: '転職希望条件:',
    statusPrefix: '進捗:',
    updateCalBtn: 'カレンダー更新',
    syncCalBtn: 'Googleカレンダー同期',
    openCalTitle: 'Googleカレンダーで予定を開く',
    createReminderBtn: '書類リマインド文を作成',
    convertedBtn: '移行済（新規求職者として再作成）',
    convertBtn: '新規求職者へ移行する',
    editBtn: '編集',
    deleteBtn: '削除',
    modalEditTitle: '退職者情報の編集',
    modalAddTitle: '新規退職者の登録',
    nameLabel: '氏名 *',
    namePlaceholder: '例: LE THI BICH PHUONG',
    natLabel: '国籍',
    suggestByNameBtn: (nat: string) => `氏名から推定: ${nat} を選択`,
    prevCompanyInputLabel: '退職予定・前職の企業名 *',
    prevCompanyPlaceholder: '退職元の企業名・施設名',
    resigDateInputLabel: '正式な退職日',
    deadlineInputLabel: '書類作成・提出期限',
    calcMinus3Btn: '退職日 − 3日で自動計算',
    autoDeadlineHint: (ymd: string) => `自動計算目安: ${ymd}（退職日の3日前）`,
    paperworkStatusLabel: '書類手続ステータス',
    indJlptLabel: '希望分野 / JLPT（求職者移行用）',
    selectIndOpt: '-- 分野を選択 --',
    phoneLabel: '電話番号',
    fbNameLabel: 'Facebook表示名',
    fbNamePlaceholder: 'Messengerの表示名',
    handoverNotesLabel: '引継ぎ・書類メモ',
    detectedDateBtn: (d: string, dl: string) =>
      `メモから退職日 ${d} を検出（期限: ${dl}）→ 適用`,
    handoverPlaceholder:
      '例: 11月30日退職予定。退職証明書・源泉徴収票の発行待ち...',
    newJobNeedLabel: '次の転職希望',
    newJobPlaceholder: '例: 埼玉・栃木で介護（手取り17万円以上）を希望...',
    cancelBtn: 'キャンセル',
    saveSyncBtn: '保存してリマインド同期',
    importTitle: (sheet: string) => `Excel / CSV 読込プレビュー（${sheet}）`,
    importDesc: (n: number) =>
      `${n} 行を自動マッピングしました。内容を確認して一括登録を実行してください。`,
    thName: '氏名',
    thNat: '国籍',
    thPrevComp: '前職企業',
    thResigDate: '退職日',
    thDeadline: '提出期限',
    thStatus: '進捗',
    thNotes: 'メモ・希望',
    closeBtn: '閉じる',
    confirmImportBtn: (n: number) => `${n} 件を一括登録する`,
    reminderTitle: (name: string) => `退職書類リマインド文面 — ${name}`,
    viLabel: 'ベトナム語 (Tiếng Việt)',
    jpLabel: '日本語',
    copyBtn: 'コピー',
    notifyMissingFields: '氏名と前職企業名を入力してください。',
    notifyNoValidRows: 'Excel/CSV内に有効なデータ行が見つかりませんでした。',
    notifyImportErr: (msg: string) => `Excel/CSV読込エラー: ${msg}`,
    notifyCopiedVi: 'リマインド文（ベトナム語）をコピーしました！',
    notifyCopiedJp: 'リマインド文（日本語）をコピーしました！',
  },
} as const;

export const ResignationTab: React.FC<ResignationTabProps> = ({
  lang = 'vi',
  records,
  onSaveRecord,
  onDeleteRecord,
  onStatusChange,
  onBulkImport,
  onSyncCalendarSingle,
  onSyncCalendarBulk,
  onConvertToCandidate,
  syncingCalendarIds,
  isBulkCalendarSyncing,
  onNotify,
  hasGasUrl = false,
  notificationPermission = 'unsupported',
  onRequestBrowserNotifications,
}) => {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState('all');
  const [nationalityFilter, setNationalityFilter] = useState('all');

  const [editingRecord, setEditingRecord] = useState<ResignationRecord | null>(null);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [formState, setFormState] = useState<Omit<ResignationRecord, 'id'>>(EMPTY_FORM);

  const [importPreviewRows, setImportPreviewRows] = useState<ResignationImportPreviewRow[]>([]);
  const [importSheetName, setImportSheetName] = useState('');
  const [isImportModalOpen, setIsImportModalOpen] = useState(false);
  const [isImporting, setIsImporting] = useState(false);

  const [reminderTarget, setReminderTarget] = useState<ResignationRecord | null>(null);

  const c = RESIG_COPY[lang];

  const needsNationalityReview = (r: ResignationRecord): boolean => {
    const nat = normalizeResignationNationality(r.nationality);
    if (nat === '国籍不明') return true;
    if (nat === 'ベトナム' && !r.manualProtectedFields?.includes('nationality')) {
      const suggested = suggestNationalityFromName(r.fullName);
      if (suggested && suggested !== 'ベトナム') return true;
    }
    return false;
  };

  const stats = useMemo(() => {
    let overdueOrUrgent = 0;
    let warn7 = 0;
    let done = 0;
    let reviewNatCount = 0;
    records.forEach((r) => {
      const a = assessResignationDeadline(r.paperworkDeadline, r.paperworkStatus);
      if (a.urgency === 'done') done += 1;
      else if (a.urgency === 'overdue' || a.urgency === 'urgent3') overdueOrUrgent += 1;
      else if (a.urgency === 'warn7') warn7 += 1;
      if (needsNationalityReview(r)) reviewNatCount += 1;
    });
    return { total: records.length, overdueOrUrgent, warn7, done, reviewNatCount };
  }, [records]);

  const filteredRecords = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    return [...records]
      .filter((r) => {
        if (statusFilter !== 'all' && r.paperworkStatus !== statusFilter) return false;
        const normalizedNat = normalizeResignationNationality(r.nationality);
        if (nationalityFilter === 'needs_review') {
          if (!needsNationalityReview(r)) return false;
        } else if (nationalityFilter !== 'all' && normalizedNat !== nationalityFilter) {
          return false;
        }
        if (!q) return true;
        return [
          r.fullName,
          r.previousCompany,
          normalizedNat,
          r.handoverNotes,
          r.newJobNeed,
          r.fbName,
          r.phone,
        ]
          .filter(Boolean)
          .some((v) => String(v).toLowerCase().includes(q));
      })
      .sort((a, b) => {
        const rank = (r: ResignationRecord) => {
          const u = assessResignationDeadline(r.paperworkDeadline, r.paperworkStatus).urgency;
          if (u === 'overdue') return 0;
          if (u === 'urgent3') return 1;
          if (u === 'warn7') return 2;
          if (u === 'normal') return 3;
          return 4;
        };
        return rank(a) - rank(b);
      });
  }, [records, searchQuery, statusFilter, nationalityFilter]);

  const openAddModal = () => {
    setEditingRecord(null);
    setFormState({ ...EMPTY_FORM });
    setIsModalOpen(true);
  };

  const openEditModal = (rec: ResignationRecord) => {
    const normalized = normalizeAndAutoComputeResignationDates(rec);
    setEditingRecord(rec);
    setFormState({
      fullName: normalized.fullName,
      nationality: normalizeResignationNationality(normalized.nationality),
      previousCompany: normalized.previousCompany,
      resignationDate: normalized.resignationDate || '',
      paperworkDeadline: normalized.paperworkDeadline || '',
      paperworkStatus: normalized.paperworkStatus || 'Chưa làm',
      handoverNotes: normalized.handoverNotes || '',
      newJobNeed: normalized.newJobNeed || '',
      phone: normalized.phone || '',
      fbName: normalized.fbName || '',
      fbProfileUrl: normalized.fbProfileUrl || '',
      industry: normalized.industry || '',
      jlpt: normalized.jlpt || 'N3',
      convertedCandidateId: normalized.convertedCandidateId,
      googleCalendarEventId: normalized.googleCalendarEventId,
      googleCalendarEventLink: normalized.googleCalendarEventLink,
      manualProtectedFields: normalized.manualProtectedFields,
      gasProtected: normalized.gasProtected,
    });
    setIsModalOpen(true);
  };

  const handleQuickNationalityChange = async (
    rec: ResignationRecord,
    nextNationality: string
  ) => {
    const normalized = normalizeResignationNationality(nextNationality);
    const updated = applyManualProtectionOnResignationSave(
      rec,
      {
        ...rec,
        nationality: normalized,
        updatedAt: new Date().toISOString(),
      },
      ['nationality']
    );
    await onSaveRecord(updated, false);
  };

  const handleSubmitForm = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formState.fullName.trim() || !formState.previousCompany.trim()) {
      onNotify(c.notifyMissingFields, 'error');
      return;
    }
    const isNew = !editingRecord;
    const rawSaved = normalizeAndAutoComputeResignationDates({
      id: editingRecord?.id || `resig_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
      ...formState,
      nationality: normalizeResignationNationality(formState.nationality),
      fullName: formState.fullName.trim(),
      previousCompany: formState.previousCompany.trim(),
      updatedAt: new Date().toISOString(),
    });
    const saved = applyManualProtectionOnResignationSave(editingRecord, rawSaved);
    setIsModalOpen(false);
    await onSaveRecord(saved, isNew);
  };

  const handleSpreadsheetUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const { rows, sheetName } = await parseResignationSpreadsheetFile(file);
      if (rows.length === 0) {
        onNotify(c.notifyNoValidRows, 'error');
        return;
      }
      setImportPreviewRows(rows);
      setImportSheetName(sheetName);
      setIsImportModalOpen(true);
    } catch (err) {
      onNotify(
        c.notifyImportErr(err instanceof Error ? err.message : 'Unknown error'),
        'error'
      );
    } finally {
      e.target.value = '';
    }
  };

  const handleConfirmImport = async () => {
    if (importPreviewRows.length === 0) return;
    setIsImporting(true);
    try {
      const cleanRecords: ResignationRecord[] = importPreviewRows.map(
        ({ _rowIndex: _ri, _isValid: _iv, _warning: _w, ...rest }) => ({
          ...rest,
          updatedAt: new Date().toISOString(),
        })
      );
      await onBulkImport(cleanRecords);
      setIsImportModalOpen(false);
      setImportPreviewRows([]);
    } finally {
      setIsImporting(false);
    }
  };

  const reminderMsgs = useMemo(
    () => (reminderTarget ? buildResignationReminderMessages(reminderTarget) : null),
    [reminderTarget]
  );

  return (
    <div className="space-y-4">
      {/* Header & Top Action Bar */}
      <div className="flex flex-col gap-3 rounded-xl border border-[#E2E8F0] bg-white p-4 shadow-[0_1px_3px_rgba(0,0,0,0.05)] sm:flex-row sm:items-center sm:justify-between">
        <div>
          <h2 className="text-base font-bold text-[#0F172A]">{c.title}</h2>
          <p className="text-xs text-[#64748B]">
            {hasGasUrl ? c.subtitleGas : c.subtitleLink}
          </p>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          {notificationPermission === 'default' && onRequestBrowserNotifications && (
            <button
              type="button"
              onClick={onRequestBrowserNotifications}
              className="inline-flex items-center gap-1.5 rounded-lg border border-amber-300 bg-amber-50 px-3 py-1.5 text-xs font-semibold text-amber-800 hover:bg-amber-100 whitespace-nowrap shrink-0 cursor-pointer"
            >
              <Bell className="h-3.5 w-3.5" aria-hidden="true" />
              {c.enableNotifyBtn}
            </button>
          )}
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={handleSpreadsheetUpload}
            className="hidden"
          />
          <button
            type="button"
            onClick={() => fileInputRef.current?.click()}
            className="inline-flex items-center gap-1.5 rounded-lg border border-[#CBD5E1] bg-white px-3 py-1.5 text-xs font-semibold text-[#0F172A] hover:bg-slate-50 whitespace-nowrap shrink-0 cursor-pointer"
          >
            <FileSpreadsheet className="h-3.5 w-3.5 text-[#16A34A]" aria-hidden="true" />
            {c.importExcelBtn}
          </button>
          <button
            type="button"
            onClick={onSyncCalendarBulk}
            disabled={isBulkCalendarSyncing}
            className="inline-flex items-center gap-1.5 rounded-lg border border-sky-200 bg-sky-50 px-3 py-1.5 text-xs font-semibold text-sky-700 hover:bg-sky-100 disabled:opacity-50 whitespace-nowrap shrink-0 cursor-pointer"
          >
            {isBulkCalendarSyncing ? (
              <Loader2 className="h-3.5 w-3.5 animate-spin" aria-hidden="true" />
            ) : (
              <Calendar className="h-3.5 w-3.5" aria-hidden="true" />
            )}
            {c.syncBulkCalBtn}
          </button>
          <button
            type="button"
            onClick={openAddModal}
            className="inline-flex items-center gap-1.5 rounded-lg bg-[#16A34A] px-3 py-1.5 text-xs font-semibold text-white hover:bg-[#15803D] whitespace-nowrap shrink-0 cursor-pointer"
          >
            <Plus className="h-3.5 w-3.5" aria-hidden="true" />
            {c.addBtn}
          </button>
        </div>
      </div>

      {stats.overdueOrUrgent > 0 && (
        <div className="flex flex-wrap items-center justify-between gap-2 rounded-xl border border-rose-200 bg-rose-50 px-4 py-2.5 text-xs text-rose-800 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
          <div className="flex items-center gap-2 font-semibold">
            <AlertTriangle className="h-4 w-4 shrink-0 text-rose-600" aria-hidden="true" />
            <span>{c.urgentBanner(stats.overdueOrUrgent)}</span>
          </div>
          <span className="rounded border border-rose-200 bg-white px-2 py-0.5 font-mono text-[10px] font-bold text-rose-700">
            {c.ruleTag}
          </span>
        </div>
      )}

      {/* KPI Strip */}
      <div className="grid grid-cols-2 gap-2.5 sm:grid-cols-4">
        <div className="rounded-lg border border-[#E2E8F0] bg-white px-3.5 py-2.5 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
          <div className="text-[11px] font-medium text-[#64748B]">{c.kpiTotal}</div>
          <div className="mt-0.5 text-lg font-bold text-[#0F172A] tabular-nums">{stats.total}</div>
        </div>
        <div className="rounded-lg border border-rose-200 bg-rose-50/70 px-3.5 py-2.5 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
          <div className="text-[11px] font-semibold text-rose-700">{c.kpiUrgent}</div>
          <div className="mt-0.5 text-lg font-bold text-rose-700 tabular-nums">
            {stats.overdueOrUrgent}
          </div>
        </div>
        <div className="rounded-lg border border-amber-200 bg-amber-50/70 px-3.5 py-2.5 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
          <div className="text-[11px] font-semibold text-amber-800">{c.kpiWarn}</div>
          <div className="mt-0.5 text-lg font-bold text-amber-700 tabular-nums">{stats.warn7}</div>
        </div>
        <div className="rounded-lg border border-emerald-200 bg-emerald-50/70 px-3.5 py-2.5 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
          <div className="text-[11px] font-semibold text-emerald-700">{c.kpiDone}</div>
          <div className="mt-0.5 text-lg font-bold text-emerald-700 tabular-nums">{stats.done}</div>
        </div>
      </div>

      {/* Search & Filter */}
      <div className="flex flex-col gap-2 rounded-xl border border-[#E2E8F0] bg-white p-3 shadow-[0_1px_3px_rgba(0,0,0,0.05)] sm:flex-row sm:items-center">
        <div className="relative flex-1">
          <Search
            className="pointer-events-none absolute left-2.5 top-2.5 h-3.5 w-3.5 text-[#64748B]"
            aria-hidden="true"
          />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder={c.searchPlaceholder}
            className="w-full rounded-lg border border-[#CBD5E1] bg-[#F8FAFC] py-1.5 pl-8 pr-3 text-xs text-[#0F172A] placeholder:text-[#64748B] focus:border-[#16A34A] focus:bg-white focus:outline-none"
          />
        </div>
        <select
          value={nationalityFilter}
          onChange={(e) => setNationalityFilter(e.target.value)}
          className="rounded-lg border border-[#CBD5E1] bg-white px-3 py-1.5 text-xs font-medium text-[#0F172A] focus:border-[#16A34A] focus:outline-none"
        >
          <option value="all">{c.allNat}</option>
          {stats.reviewNatCount > 0 && (
            <option value="needs_review">{c.needsNatReview(stats.reviewNatCount)}</option>
          )}
          {RESIGNATION_NATIONALITIES.map((nat) => (
            <option key={nat} value={nat}>
              {nat === '国籍不明' ? c.unknownNatOpt : nat}
            </option>
          ))}
        </select>
        <select
          value={statusFilter}
          onChange={(e) => setStatusFilter(e.target.value)}
          className="rounded-lg border border-[#CBD5E1] bg-white px-3 py-1.5 text-xs font-medium text-[#0F172A] focus:border-[#16A34A] focus:outline-none"
        >
          <option value="all">{c.allStatus(records.length)}</option>
          {RESIGNATION_STATUSES.map((st) => (
            <option key={st} value={st}>
              {st}
            </option>
          ))}
        </select>
      </div>

      {/* Cards List */}
      {filteredRecords.length === 0 ? (
        <div className="rounded-xl border border-dashed border-[#CBD5E1] bg-white p-8 text-center text-xs text-[#64748B] shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
          {c.emptyMsg}
        </div>
      ) : (
        <div className="space-y-3">
          {filteredRecords.map((rec) => {
            const displayName = getResignationDisplayName(rec, lang);
            const deadlineInfo = assessResignationDeadline(
              rec.paperworkDeadline,
              rec.paperworkStatus
            );
            const isSyncing = Boolean(syncingCalendarIds[rec.id]);
            const hasCalendarLink = Boolean(rec.googleCalendarEventId);
            const currentNat = normalizeResignationNationality(rec.nationality);
            const isCustomNat = !RESIGNATION_NATIONALITIES.includes(
              currentNat as (typeof RESIGNATION_NATIONALITIES)[number]
            );
            const suggestedNat = suggestNationalityFromName(rec.fullName);
            const showNatSuggestion =
              Boolean(suggestedNat) &&
              ((currentNat === '国籍不明' && suggestedNat !== null) ||
                (currentNat === 'ベトナム' &&
                  suggestedNat !== 'ベトナム' &&
                  !rec.manualProtectedFields?.includes('nationality')));
            const isProtected =
              Boolean(rec.gasProtected) ||
              Boolean(rec.manualProtectedFields && rec.manualProtectedFields.length > 0);

            const urgencyClass =
              deadlineInfo.urgency === 'overdue' || deadlineInfo.urgency === 'urgent3'
                ? 'border-rose-200 bg-rose-50 text-rose-700'
                : deadlineInfo.urgency === 'warn7'
                ? 'border-amber-300 bg-amber-50 text-amber-800'
                : deadlineInfo.urgency === 'done'
                ? 'border-emerald-200 bg-emerald-50 text-emerald-700'
                : 'border-[#E2E8F0] bg-slate-100 text-[#475569]';

            return (
              <div
                key={rec.id}
                className="rounded-xl border border-[#E2E8F0] bg-white p-4 shadow-[0_1px_3px_rgba(0,0,0,0.05)] transition hover:border-[#CBD5E1]"
              >
                <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
                  <div className="space-y-2">
                    <div className="flex flex-wrap items-center gap-2">
                      <h3 className="text-sm font-bold text-[#0F172A]">{displayName}</h3>
                      <select
                        aria-label={c.natAria(displayName)}
                        value={currentNat}
                        onChange={(e) =>
                          void handleQuickNationalityChange(rec, e.target.value)
                        }
                        className={`rounded border px-2 py-0.5 text-[11px] font-semibold focus:border-[#16A34A] focus:outline-none ${
                          currentNat === '国籍不明'
                            ? 'border-amber-300 bg-amber-50 text-amber-800'
                            : 'border-[#CBD5E1] bg-[#F8FAFC] text-[#0F172A]'
                        }`}
                      >
                        {isCustomNat && <option value={currentNat}>{currentNat}</option>}
                        {RESIGNATION_NATIONALITIES.map((nat) => (
                          <option key={nat} value={nat}>
                            {nat === '国籍不明' ? c.unknownNatShort : nat}
                          </option>
                        ))}
                      </select>
                      {showNatSuggestion && suggestedNat && (
                        <button
                          type="button"
                          onClick={() => void handleQuickNationalityChange(rec, suggestedNat)}
                          title={c.suggestNatTitle}
                          className="inline-flex items-center gap-1 rounded border border-sky-200 bg-sky-50 px-1.5 py-0.5 text-[10px] font-semibold text-sky-700 hover:bg-sky-100 cursor-pointer"
                        >
                          <Sparkles className="h-2.5 w-2.5" aria-hidden="true" />
                          {c.suggestNatBtn(suggestedNat)}
                        </button>
                      )}
                      {rec.jlpt && (
                        <span className="rounded border border-emerald-200 bg-emerald-50 px-1.5 py-0.5 text-[11px] font-semibold text-emerald-700">
                          {rec.jlpt}
                        </span>
                      )}
                      {isProtected && (
                        <span
                          title={c.gasProtectedTitle(
                            (rec.manualProtectedFields || ['nationality']).join(', ')
                          )}
                          className="inline-flex items-center gap-1 rounded border border-emerald-200 bg-emerald-50 px-1.5 py-0.5 text-[10px] font-semibold text-emerald-700"
                        >
                          <Lock className="h-2.5 w-2.5" aria-hidden="true" />
                          {c.gasProtectedBadge}
                        </span>
                      )}
                      <span
                        className={`inline-flex items-center gap-1 rounded border px-2 py-0.5 text-[11px] font-semibold ${urgencyClass}`}
                      >
                        {(deadlineInfo.urgency === 'overdue' ||
                          deadlineInfo.urgency === 'urgent3' ||
                          deadlineInfo.urgency === 'warn7') && (
                          <AlertTriangle className="h-3 w-3" aria-hidden="true" />
                        )}
                        {deadlineInfo.urgency === 'done' && (
                          <CheckCircle2 className="h-3 w-3" aria-hidden="true" />
                        )}
                        {deadlineInfo.label}
                      </span>
                    </div>

                    <div className="grid grid-cols-1 gap-x-4 gap-y-1 text-xs text-[#475569] sm:grid-cols-2">
                      <div>
                        <span className="text-[#64748B]">{c.prevCompanyLabel} </span>
                        <span className="font-semibold text-[#0F172A]">{rec.previousCompany}</span>
                      </div>
                      <div>
                        <span className="text-[#64748B]">{c.resigDateLabel} </span>
                        <span className="font-semibold text-[#0F172A] tabular-nums">
                          {rec.resignationDate || c.unset}
                        </span>
                      </div>
                      <div>
                        <span className="text-[#64748B]">{c.deadlineLabel} </span>
                        <span className="font-semibold text-amber-700 tabular-nums">
                          {rec.paperworkDeadline || c.noDeadline}
                        </span>
                      </div>
                      <div>
                        <span className="text-[#64748B]">{c.contactIndLabel} </span>
                        <span className="text-[#0F172A]">
                          {[rec.phone, rec.fbName, rec.industry].filter(Boolean).join(' • ') ||
                            c.unset}
                        </span>
                      </div>
                    </div>

                    {rec.handoverNotes && (
                      <div className="rounded-lg border border-[#E2E8F0] bg-[#F8FAFC] p-2.5 text-xs text-[#334155]">
                        <span className="font-semibold text-[#64748B]">{c.handoverLabel} </span>
                        <span className="whitespace-pre-wrap">{rec.handoverNotes}</span>
                      </div>
                    )}

                    {rec.newJobNeed && (
                      <div className="rounded-lg border border-emerald-200 bg-emerald-50/60 p-2.5 text-xs text-emerald-900">
                        <span className="font-semibold text-emerald-700">{c.newJobLabel} </span>
                        <span className="whitespace-pre-wrap">{rec.newJobNeed}</span>
                      </div>
                    )}
                  </div>

                  {/* Right Action Controls */}
                  <div className="flex shrink-0 flex-col items-stretch gap-2 sm:w-56">
                    <select
                      value={rec.paperworkStatus}
                      onChange={(e) => onStatusChange(rec, e.target.value)}
                      className="rounded-lg border border-[#CBD5E1] bg-white px-2.5 py-1.5 text-xs font-semibold text-[#0F172A] focus:border-[#16A34A] focus:outline-none"
                    >
                      {RESIGNATION_STATUSES.map((st) => (
                        <option key={st} value={st}>
                          {c.statusPrefix} {st}
                        </option>
                      ))}
                    </select>

                    <div className="flex items-center gap-1.5">
                      <button
                        type="button"
                        onClick={() => onSyncCalendarSingle(rec)}
                        disabled={isSyncing || rec.paperworkStatus === 'Hoàn tất'}
                        className={`flex flex-1 items-center justify-center gap-1 rounded-lg border px-2 py-1.5 text-[11px] font-semibold transition disabled:opacity-45 cursor-pointer ${
                          hasCalendarLink
                            ? 'border-emerald-200 bg-emerald-50 text-emerald-700 hover:bg-emerald-100'
                            : 'border-sky-200 bg-sky-50 text-sky-700 hover:bg-sky-100'
                        }`}
                      >
                        {isSyncing ? (
                          <Loader2 className="h-3 w-3 animate-spin" aria-hidden="true" />
                        ) : hasCalendarLink ? (
                          <CalendarCheck className="h-3 w-3" aria-hidden="true" />
                        ) : (
                          <Calendar className="h-3 w-3" aria-hidden="true" />
                        )}
                        {hasCalendarLink ? c.updateCalBtn : c.syncCalBtn}
                      </button>

                      {isValidGoogleCalendarUrl(rec.googleCalendarEventLink) && (
                        <a
                          href={rec.googleCalendarEventLink}
                          target="_blank"
                          rel="noopener noreferrer"
                          title={c.openCalTitle}
                          className="inline-flex items-center justify-center rounded-lg border border-[#CBD5E1] bg-white p-1.5 text-[#475569] hover:bg-slate-50 hover:text-[#0F172A]"
                        >
                          <ExternalLink className="h-3.5 w-3.5" aria-hidden="true" />
                        </a>
                      )}
                    </div>

                    <button
                      type="button"
                      onClick={() => setReminderTarget(rec)}
                      className="inline-flex items-center justify-center gap-1.5 rounded-lg border border-amber-300 bg-amber-50 px-2.5 py-1.5 text-[11px] font-semibold text-amber-800 hover:bg-amber-100 cursor-pointer"
                    >
                      <MessageSquare className="h-3 w-3" aria-hidden="true" />
                      {c.createReminderBtn}
                    </button>

                    <button
                      type="button"
                      onClick={() => onConvertToCandidate(rec)}
                      className="inline-flex items-center justify-center gap-1.5 rounded-lg border border-emerald-200 bg-emerald-50 px-2.5 py-1.5 text-[11px] font-semibold text-emerald-700 hover:bg-emerald-100 cursor-pointer"
                    >
                      <UserPlus className="h-3 w-3" aria-hidden="true" />
                      {rec.convertedCandidateId ? c.convertedBtn : c.convertBtn}
                    </button>

                    <div className="flex items-center justify-end gap-1.5 pt-1">
                      <button
                        type="button"
                        onClick={() => openEditModal(rec)}
                        className="inline-flex items-center gap-1 rounded border border-[#CBD5E1] bg-white px-2 py-1 text-[11px] font-semibold text-[#334155] hover:bg-slate-50 cursor-pointer"
                      >
                        <Edit3 className="h-3 w-3" aria-hidden="true" />
                        {c.editBtn}
                      </button>
                      <button
                        type="button"
                        onClick={() => onDeleteRecord(rec)}
                        className="inline-flex items-center gap-1 rounded border border-rose-200 bg-rose-50 px-2 py-1 text-[11px] font-semibold text-rose-700 hover:bg-rose-100 cursor-pointer"
                      >
                        <Trash2 className="h-3 w-3" aria-hidden="true" />
                        {c.deleteBtn}
                      </button>
                    </div>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* Add/Edit Resignation Modal */}
      {isModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4 backdrop-blur-[1px]">
          <div className="max-h-[90vh] w-full max-w-2xl overflow-y-auto rounded-xl border border-[#E2E8F0] bg-white p-5 text-[#0F172A] shadow-2xl">
            <div className="mb-4 flex items-center justify-between border-b border-[#E2E8F0] pb-3">
              <h3 className="text-sm font-bold text-[#0F172A]">
                {editingRecord ? c.modalEditTitle : c.modalAddTitle}
              </h3>
              <button
                type="button"
                onClick={() => setIsModalOpen(false)}
                className="rounded p-1 text-[#64748B] hover:bg-slate-100 hover:text-[#0F172A] cursor-pointer"
              >
                <X className="h-4 w-4" aria-hidden="true" />
              </button>
            </div>

            <form onSubmit={handleSubmitForm} className="space-y-3 text-xs">
              <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
                <div>
                  <label className="mb-1 block font-semibold text-[#475569]">{c.nameLabel}</label>
                  <input
                    required
                    type="text"
                    value={formState.fullName}
                    onChange={(e) => setFormState((s) => ({ ...s, fullName: e.target.value }))}
                    placeholder={c.namePlaceholder}
                    className="w-full rounded-lg border border-[#CBD5E1] bg-[#F8FAFC] px-3 py-2 text-[#0F172A] focus:border-[#16A34A] focus:bg-white focus:outline-none"
                  />
                </div>
                <div>
                  <label className="mb-1 block font-semibold text-[#475569]">{c.natLabel}</label>
                  <select
                    value={normalizeResignationNationality(formState.nationality)}
                    onChange={(e) =>
                      setFormState((s) => ({ ...s, nationality: e.target.value }))
                    }
                    className="w-full rounded-lg border border-[#CBD5E1] bg-white px-3 py-2 text-[#0F172A] focus:border-[#16A34A] focus:outline-none"
                  >
                    {formState.nationality &&
                      !RESIGNATION_NATIONALITIES.includes(
                        normalizeResignationNationality(
                          formState.nationality
                        ) as (typeof RESIGNATION_NATIONALITIES)[number]
                      ) && (
                        <option value={formState.nationality}>{formState.nationality}</option>
                      )}
                    {RESIGNATION_NATIONALITIES.map((nat) => (
                      <option key={nat} value={nat}>
                        {nat === '国籍不明' ? c.unknownNatOpt : nat}
                      </option>
                    ))}
                  </select>
                  {(() => {
                    const suggested = suggestNationalityFromName(formState.fullName);
                    const cur = normalizeResignationNationality(formState.nationality);
                    if (suggested && cur !== suggested && cur === '国籍不明') {
                      return (
                        <button
                          type="button"
                          onClick={() =>
                            setFormState((s) => ({ ...s, nationality: suggested }))
                          }
                          className="mt-1 inline-flex items-center gap-1 text-[11px] font-semibold text-sky-700 hover:text-sky-800 cursor-pointer"
                        >
                          <Sparkles className="h-3 w-3" aria-hidden="true" />
                          {c.suggestByNameBtn(suggested)}
                        </button>
                      );
                    }
                    return null;
                  })()}
                </div>
                <div className="sm:col-span-2">
                  <label className="mb-1 block font-semibold text-[#475569]">
                    {c.prevCompanyInputLabel}
                  </label>
                  <input
                    required
                    type="text"
                    value={formState.previousCompany}
                    onChange={(e) =>
                      setFormState((s) => ({ ...s, previousCompany: e.target.value }))
                    }
                    placeholder={c.prevCompanyPlaceholder}
                    className="w-full rounded-lg border border-[#CBD5E1] bg-[#F8FAFC] px-3 py-2 text-[#0F172A] focus:border-[#16A34A] focus:bg-white focus:outline-none"
                  />
                </div>
                <div>
                  <label className="mb-1 block font-semibold text-[#475569]">
                    {c.resigDateInputLabel}
                  </label>
                  <input
                    type="date"
                    value={formState.resignationDate}
                    onChange={(e) => {
                      const nextResig = e.target.value;
                      setFormState((s) => {
                        const prevAuto = subtractDaysFromYmd(s.resignationDate, 3);
                        const shouldAutoFillDeadline =
                          !s.paperworkDeadline ||
                          s.paperworkDeadline === s.resignationDate ||
                          (prevAuto && s.paperworkDeadline === prevAuto);
                        return {
                          ...s,
                          resignationDate: nextResig,
                          paperworkDeadline: shouldAutoFillDeadline
                            ? subtractDaysFromYmd(nextResig, 3)
                            : s.paperworkDeadline,
                        };
                      });
                    }}
                    className="w-full rounded-lg border border-[#CBD5E1] bg-[#F8FAFC] px-3 py-2 text-[#0F172A] focus:border-[#16A34A] focus:bg-white focus:outline-none"
                  />
                </div>
                <div>
                  <div className="mb-1 flex items-center justify-between gap-2">
                    <label className="block font-semibold text-[#475569]">
                      {c.deadlineInputLabel}
                    </label>
                    {formState.resignationDate && (
                      <button
                        type="button"
                        onClick={() =>
                          setFormState((s) => ({
                            ...s,
                            paperworkDeadline: subtractDaysFromYmd(s.resignationDate, 3),
                          }))
                        }
                        className="text-[10px] font-semibold text-[#16A34A] hover:underline cursor-pointer"
                      >
                        {c.calcMinus3Btn}
                      </button>
                    )}
                  </div>
                  <input
                    type="date"
                    value={formState.paperworkDeadline}
                    onChange={(e) =>
                      setFormState((s) => ({ ...s, paperworkDeadline: e.target.value }))
                    }
                    className="w-full rounded-lg border border-[#CBD5E1] bg-[#F8FAFC] px-3 py-2 text-[#0F172A] focus:border-[#16A34A] focus:bg-white focus:outline-none"
                  />
                  {formState.resignationDate && (
                    <p className="mt-1 text-[10px] text-[#64748B]">
                      {c.autoDeadlineHint(subtractDaysFromYmd(formState.resignationDate, 3))}
                    </p>
                  )}
                </div>
                <div>
                  <label className="mb-1 block font-semibold text-[#475569]">
                    {c.paperworkStatusLabel}
                  </label>
                  <select
                    value={formState.paperworkStatus}
                    onChange={(e) =>
                      setFormState((s) => ({ ...s, paperworkStatus: e.target.value }))
                    }
                    className="w-full rounded-lg border border-[#CBD5E1] bg-white px-3 py-2 text-[#0F172A] focus:border-[#16A34A] focus:outline-none"
                  >
                    {RESIGNATION_STATUSES.map((st) => (
                      <option key={st} value={st}>
                        {st}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="mb-1 block font-semibold text-[#475569]">{c.indJlptLabel}</label>
                  <div className="flex gap-2">
                    <select
                      value={formState.industry || ''}
                      onChange={(e) =>
                        setFormState((s) => ({ ...s, industry: e.target.value }))
                      }
                      className="flex-1 rounded-lg border border-[#CBD5E1] bg-white px-2 py-2 text-[#0F172A]"
                    >
                      <option value="">{c.selectIndOpt}</option>
                      {INDUSTRIES.map((ind) => (
                        <option key={ind} value={ind}>
                          {ind}
                        </option>
                      ))}
                    </select>
                    <select
                      value={formState.jlpt || 'N3'}
                      onChange={(e) =>
                        setFormState((s) => ({ ...s, jlpt: e.target.value }))
                      }
                      className="w-24 rounded-lg border border-[#CBD5E1] bg-white px-2 py-2 text-[#0F172A]"
                    >
                      {JLPT_LEVELS.map((j) => (
                        <option key={j} value={j}>
                          {j}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
                <div>
                  <label className="mb-1 block font-semibold text-[#475569]">{c.phoneLabel}</label>
                  <input
                    type="text"
                    value={formState.phone || ''}
                    onChange={(e) => setFormState((s) => ({ ...s, phone: e.target.value }))}
                    placeholder="080-xxxx-xxxx"
                    className="w-full rounded-lg border border-[#CBD5E1] bg-[#F8FAFC] px-3 py-2 text-[#0F172A] focus:border-[#16A34A] focus:bg-white focus:outline-none"
                  />
                </div>
                <div>
                  <label className="mb-1 block font-semibold text-[#475569]">{c.fbNameLabel}</label>
                  <input
                    type="text"
                    value={formState.fbName || ''}
                    onChange={(e) => setFormState((s) => ({ ...s, fbName: e.target.value }))}
                    placeholder={c.fbNamePlaceholder}
                    className="w-full rounded-lg border border-[#CBD5E1] bg-[#F8FAFC] px-3 py-2 text-[#0F172A] focus:border-[#16A34A] focus:bg-white focus:outline-none"
                  />
                </div>
              </div>

              <div>
                <div className="mb-1 flex items-center justify-between gap-2">
                  <label className="block font-semibold text-[#475569]">
                    {c.handoverNotesLabel}
                  </label>
                  {(() => {
                    const detected = extractResignationDateFromText(formState.handoverNotes);
                    if (detected && detected !== formState.resignationDate) {
                      return (
                        <button
                          type="button"
                          onClick={() =>
                            setFormState((s) => ({
                              ...s,
                              resignationDate: detected,
                              paperworkDeadline: subtractDaysFromYmd(detected, 3),
                            }))
                          }
                          className="inline-flex items-center gap-1 text-[10px] font-semibold text-sky-700 hover:underline cursor-pointer"
                        >
                          <Sparkles className="h-3 w-3" aria-hidden="true" />
                          {c.detectedDateBtn(detected, subtractDaysFromYmd(detected, 3))}
                        </button>
                      );
                    }
                    return null;
                  })()}
                </div>
                <textarea
                  rows={3}
                  value={formState.handoverNotes || ''}
                  onChange={(e) => {
                    const nextNotes = e.target.value;
                    const detected = extractResignationDateFromText(nextNotes);
                    setFormState((s) => {
                      const nextResig = !s.resignationDate && detected ? detected : s.resignationDate;
                      const nextDeadline =
                        !s.resignationDate &&
                        detected &&
                        (!s.paperworkDeadline || s.paperworkDeadline === detected)
                          ? subtractDaysFromYmd(detected, 3)
                          : s.paperworkDeadline;
                      return {
                        ...s,
                        handoverNotes: nextNotes,
                        resignationDate: nextResig,
                        paperworkDeadline: nextDeadline,
                      };
                    });
                  }}
                  placeholder={c.handoverPlaceholder}
                  className="w-full rounded-lg border border-[#CBD5E1] bg-[#F8FAFC] px-3 py-2 text-[#0F172A] focus:border-[#16A34A] focus:bg-white focus:outline-none"
                />
              </div>

              <div>
                <label className="mb-1 block font-semibold text-[#475569]">{c.newJobNeedLabel}</label>
                <textarea
                  rows={2}
                  value={formState.newJobNeed || ''}
                  onChange={(e) =>
                    setFormState((s) => ({ ...s, newJobNeed: e.target.value }))
                  }
                  placeholder={c.newJobPlaceholder}
                  className="w-full rounded-lg border border-[#CBD5E1] bg-[#F8FAFC] px-3 py-2 text-[#0F172A] focus:border-[#16A34A] focus:bg-white focus:outline-none"
                />
              </div>

              <div className="flex justify-end gap-2 border-t border-[#E2E8F0] pt-3">
                <button
                  type="button"
                  onClick={() => setIsModalOpen(false)}
                  className="rounded-lg border border-[#CBD5E1] bg-white px-3 py-1.5 font-semibold text-[#475569] hover:bg-slate-50 cursor-pointer"
                >
                  {c.cancelBtn}
                </button>
                <button
                  type="submit"
                  className="rounded-lg bg-[#16A34A] px-4 py-1.5 font-semibold text-white hover:bg-[#15803D] cursor-pointer"
                >
                  {c.saveSyncBtn}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Excel/CSV Preview Table Modal before Bulk Import */}
      {isImportModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4 backdrop-blur-[1px]">
          <div className="flex max-h-[88vh] w-full max-w-5xl flex-col rounded-xl border border-[#E2E8F0] bg-white p-5 text-[#0F172A] shadow-2xl">
            <div className="mb-3 flex items-center justify-between border-b border-[#E2E8F0] pb-3">
              <div>
                <h3 className="text-sm font-bold text-[#0F172A]">
                  {c.importTitle(importSheetName)}
                </h3>
                <p className="text-xs text-[#64748B]">{c.importDesc(importPreviewRows.length)}</p>
              </div>
              <button
                type="button"
                onClick={() => setIsImportModalOpen(false)}
                className="rounded p-1 text-[#64748B] hover:bg-slate-100 cursor-pointer"
              >
                <X className="h-4 w-4" aria-hidden="true" />
              </button>
            </div>

            <div className="flex-1 overflow-auto rounded-lg border border-[#E2E8F0]">
              <table className="w-full border-collapse text-left text-xs">
                <thead className="sticky top-0 bg-[#F8FAFC] text-[#475569]">
                  <tr>
                    <th className="border-b border-[#E2E8F0] px-2.5 py-2">#</th>
                    <th className="border-b border-[#E2E8F0] px-2.5 py-2">{c.thName}</th>
                    <th className="border-b border-[#E2E8F0] px-2.5 py-2">{c.thNat}</th>
                    <th className="border-b border-[#E2E8F0] px-2.5 py-2">{c.thPrevComp}</th>
                    <th className="border-b border-[#E2E8F0] px-2.5 py-2">{c.thResigDate}</th>
                    <th className="border-b border-[#E2E8F0] px-2.5 py-2">{c.thDeadline}</th>
                    <th className="border-b border-[#E2E8F0] px-2.5 py-2">{c.thStatus}</th>
                    <th className="border-b border-[#E2E8F0] px-2.5 py-2">{c.thNotes}</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[#E2E8F0] text-[#0F172A]">
                  {importPreviewRows.map((row, idx) => {
                    const rowNat = normalizeResignationNationality(row.nationality);
                    return (
                      <tr key={row.id} className="hover:bg-slate-50">
                        <td className="px-2.5 py-2 text-[#64748B]">{row._rowIndex}</td>
                        <td className="px-2.5 py-2 font-semibold">{row.fullName}</td>
                        <td className="px-2.5 py-2">
                          <select
                            aria-label={`${c.thNat} ${row._rowIndex}`}
                            value={rowNat}
                            onChange={(e) => {
                              const nextNat = normalizeResignationNationality(e.target.value);
                              setImportPreviewRows((prev) =>
                                prev.map((item, i) =>
                                  i === idx
                                    ? {
                                        ...item,
                                        nationality: nextNat,
                                        manualProtectedFields: ['nationality'],
                                        gasProtected: true,
                                      }
                                    : item
                                )
                              );
                            }}
                            className={`rounded border px-1.5 py-1 text-[11px] font-medium focus:border-[#16A34A] focus:outline-none ${
                              rowNat === '国籍不明'
                                ? 'border-amber-300 bg-amber-50 text-amber-800'
                                : 'border-[#CBD5E1] bg-white text-[#0F172A]'
                            }`}
                          >
                            {RESIGNATION_NATIONALITIES.map((nat) => (
                              <option key={nat} value={nat}>
                                {nat === '国籍不明' ? c.unknownNatShort : nat}
                              </option>
                            ))}
                          </select>
                        </td>
                        <td className="px-2.5 py-2">{row.previousCompany}</td>
                        <td className="px-2.5 py-2 tabular-nums">{row.resignationDate || '-'}</td>
                        <td className="px-2.5 py-2 font-semibold text-amber-700 tabular-nums">
                          {row.paperworkDeadline || '-'}
                        </td>
                        <td className="px-2.5 py-2">{row.paperworkStatus}</td>
                        <td className="max-w-xs px-2.5 py-2 text-[11px] text-[#64748B]">
                          {[row.handoverNotes, row.newJobNeed, row._warning]
                            .filter(Boolean)
                            .join(' | ') || '-'}
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <div className="mt-4 flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setIsImportModalOpen(false)}
                className="rounded-lg border border-[#CBD5E1] bg-white px-3 py-1.5 text-xs font-semibold text-[#475569] hover:bg-slate-50 cursor-pointer"
              >
                {c.closeBtn}
              </button>
              <button
                type="button"
                disabled={isImporting}
                onClick={handleConfirmImport}
                className="inline-flex items-center gap-1.5 rounded-lg bg-[#16A34A] px-4 py-1.5 text-xs font-semibold text-white hover:bg-[#15803D] disabled:opacity-50 cursor-pointer"
              >
                {isImporting && <Loader2 className="h-3.5 w-3.5 animate-spin" aria-hidden="true" />}
                {c.confirmImportBtn(importPreviewRows.length)}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Reminder Message Modal */}
      {reminderTarget && reminderMsgs && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/50 p-4 backdrop-blur-[1px]">
          <div className="w-full max-w-xl rounded-xl border border-[#E2E8F0] bg-white p-5 text-[#0F172A] shadow-2xl">
            <div className="mb-3 flex items-center justify-between border-b border-[#E2E8F0] pb-2.5">
              <h3 className="text-sm font-bold text-[#0F172A]">
                {c.reminderTitle(getResignationDisplayName(reminderTarget, lang))}
              </h3>
              <button
                type="button"
                onClick={() => setReminderTarget(null)}
                className="rounded p-1 text-[#64748B] hover:bg-slate-100 cursor-pointer"
              >
                <X className="h-4 w-4" aria-hidden="true" />
              </button>
            </div>
            <div className="space-y-3 text-xs">
              <div className="rounded-lg border border-[#E2E8F0] bg-[#F8FAFC] p-3">
                <div className="mb-1.5 flex items-center justify-between">
                  <span className="font-semibold text-[#16A34A]">{c.viLabel}</span>
                  <button
                    type="button"
                    onClick={async () => {
                      await copyTextSafely(reminderMsgs.vi);
                      onNotify(c.notifyCopiedVi, 'success');
                    }}
                    className="inline-flex items-center gap-1 rounded border border-emerald-200 bg-emerald-50 px-2 py-1 text-[11px] font-semibold text-emerald-700 hover:bg-emerald-100 cursor-pointer"
                  >
                    <Copy className="h-3 w-3" aria-hidden="true" />
                    {c.copyBtn}
                  </button>
                </div>
                <p className="whitespace-pre-wrap text-[#0F172A]">{reminderMsgs.vi}</p>
              </div>
              <div className="rounded-lg border border-[#E2E8F0] bg-[#F8FAFC] p-3">
                <div className="mb-1.5 flex items-center justify-between">
                  <span className="font-semibold text-sky-700">{c.jpLabel}</span>
                  <button
                    type="button"
                    onClick={async () => {
                      await copyTextSafely(reminderMsgs.jp);
                      onNotify(c.notifyCopiedJp, 'success');
                    }}
                    className="inline-flex items-center gap-1 rounded border border-sky-200 bg-sky-50 px-2 py-1 text-[11px] font-semibold text-sky-700 hover:bg-sky-100 cursor-pointer"
                  >
                    <Copy className="h-3 w-3" aria-hidden="true" />
                    {c.copyBtn}
                  </button>
                </div>
                <p className="whitespace-pre-wrap text-[#0F172A]">{reminderMsgs.jp}</p>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
