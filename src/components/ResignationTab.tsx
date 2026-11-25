import React, { useMemo, useRef, useState } from 'react';
import {
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
  AlertTriangle,
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
  normalizeResignationNationality,
  parseResignationSpreadsheetFile,
  suggestNationalityFromName,
  type ResignationImportPreviewRow,
} from '../utils/resignationHelpers';
import { copyTextSafely } from '../utils/crmHelpers';
import { isValidGoogleCalendarUrl } from '../utils/calendarHelpers';

interface ResignationTabProps {
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

export const ResignationTab: React.FC<ResignationTabProps> = ({
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
    setEditingRecord(rec);
    setFormState({
      fullName: rec.fullName,
      nationality: normalizeResignationNationality(rec.nationality),
      previousCompany: rec.previousCompany,
      resignationDate: rec.resignationDate || '',
      paperworkDeadline: rec.paperworkDeadline || '',
      paperworkStatus: rec.paperworkStatus || 'Chưa làm',
      handoverNotes: rec.handoverNotes || '',
      newJobNeed: rec.newJobNeed || '',
      phone: rec.phone || '',
      fbName: rec.fbName || '',
      fbProfileUrl: rec.fbProfileUrl || '',
      industry: rec.industry || '',
      jlpt: rec.jlpt || 'N3',
      convertedCandidateId: rec.convertedCandidateId,
      googleCalendarEventId: rec.googleCalendarEventId,
      googleCalendarEventLink: rec.googleCalendarEventLink,
      manualProtectedFields: rec.manualProtectedFields,
      gasProtected: rec.gasProtected,
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
      onNotify('Vui lòng nhập đầy đủ Họ tên và Công ty cũ.', 'error');
      return;
    }
    const isNew = !editingRecord;
    const rawSaved: ResignationRecord = {
      id: editingRecord?.id || `resig_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
      ...formState,
      nationality: normalizeResignationNationality(formState.nationality),
      fullName: formState.fullName.trim(),
      previousCompany: formState.previousCompany.trim(),
      updatedAt: new Date().toISOString(),
    };
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
        onNotify('Không tìm thấy dòng dữ liệu hợp lệ trong file Excel/CSV.', 'error');
        return;
      }
      setImportPreviewRows(rows);
      setImportSheetName(sheetName);
      setIsImportModalOpen(true);
    } catch (err) {
      onNotify(
        `Lỗi đọc file Excel/CSV: ${err instanceof Error ? err.message : 'Unknown error'}`,
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
      <div className="flex flex-col gap-3 rounded-xl border border-zinc-800 bg-[#141417] p-4 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <h2 className="text-base font-bold text-zinc-100">
            Quản lý Người nghỉ việc &amp; Thủ tục Giấy tờ (退職者管理)
          </h2>
          <p className="text-xs text-zinc-400">
            Theo dõi ngày nghỉ việc, hạn nộp giấy tờ (退職証明書・源泉徴収票), nhắc Google Calendar (3 ngày / 1 ngày / 2 giờ) &amp; chuyển tiếp ứng viên mới.
          </p>
        </div>
        <div className="flex flex-wrap items-center gap-2">
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
            className="inline-flex items-center gap-1.5 rounded-lg border border-zinc-700 bg-zinc-800/80 px-3 py-1.5 text-xs font-medium text-zinc-200 hover:bg-zinc-700"
          >
            <FileSpreadsheet className="h-3.5 w-3.5 text-emerald-400" />
            Nhập Excel / CSV hàng loạt
          </button>
          <button
            type="button"
            onClick={onSyncCalendarBulk}
            disabled={isBulkCalendarSyncing}
            className="inline-flex items-center gap-1.5 rounded-lg border border-sky-500/40 bg-sky-500/15 px-3 py-1.5 text-xs font-medium text-sky-300 hover:bg-sky-500/25 disabled:opacity-50"
          >
            {isBulkCalendarSyncing ? (
              <Loader2 className="h-3.5 w-3.5 animate-spin" />
            ) : (
              <Calendar className="h-3.5 w-3.5" />
            )}
            Đồng bộ toàn bộ Lịch Google
          </button>
          <button
            type="button"
            onClick={openAddModal}
            className="inline-flex items-center gap-1.5 rounded-lg bg-emerald-500 px-3 py-1.5 text-xs font-semibold text-zinc-950 hover:bg-emerald-400"
          >
            <Plus className="h-3.5 w-3.5" />
            Thêm người nghỉ việc
          </button>
        </div>
      </div>

      {/* KPI Strip */}
      <div className="grid grid-cols-2 gap-2 sm:grid-cols-4">
        <div className="rounded-lg border border-zinc-800 bg-[#141417] px-3 py-2.5">
          <div className="text-[11px] text-zinc-400">Tổng hồ sơ</div>
          <div className="mt-0.5 text-lg font-bold text-zinc-100">{stats.total}</div>
        </div>
        <div className="rounded-lg border border-rose-500/30 bg-rose-950/20 px-3 py-2.5">
          <div className="text-[11px] text-rose-300">Gấp ≤ 3 ngày / Quá hạn</div>
          <div className="mt-0.5 text-lg font-bold text-rose-400">{stats.overdueOrUrgent}</div>
        </div>
        <div className="rounded-lg border border-amber-500/30 bg-amber-950/20 px-3 py-2.5">
          <div className="text-[11px] text-amber-300">Sắp đến hạn (≤ 7 ngày)</div>
          <div className="mt-0.5 text-lg font-bold text-amber-400">{stats.warn7}</div>
        </div>
        <div className="rounded-lg border border-emerald-500/30 bg-emerald-950/20 px-3 py-2.5">
          <div className="text-[11px] text-emerald-300">Đã hoàn tất giấy tờ</div>
          <div className="mt-0.5 text-lg font-bold text-emerald-400">{stats.done}</div>
        </div>
      </div>

      {/* Search & Filter */}
      <div className="flex flex-col gap-2 rounded-xl border border-zinc-800 bg-[#141417] p-3 sm:flex-row sm:items-center">
        <div className="relative flex-1">
          <Search className="pointer-events-none absolute left-2.5 top-2.5 h-3.5 w-3.5 text-zinc-500" />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Tìm theo Họ tên, Công ty cũ, Quốc tịch, Ghi chú bàn giao..."
            className="w-full rounded-lg border border-zinc-800 bg-zinc-950 py-1.5 pl-8 pr-3 text-xs text-zinc-100 placeholder:text-zinc-500 focus:border-emerald-500 focus:outline-none"
          />
        </div>
        <select
          value={nationalityFilter}
          onChange={(e) => setNationalityFilter(e.target.value)}
          className="rounded-lg border border-zinc-800 bg-zinc-950 px-3 py-1.5 text-xs text-zinc-200 focus:border-emerald-500 focus:outline-none"
        >
          <option value="all">Tất cả quốc tịch</option>
          {stats.reviewNatCount > 0 && (
            <option value="needs_review">
              ⚠️ Cần rà soát quốc tịch ({stats.reviewNatCount})
            </option>
          )}
          {RESIGNATION_NATIONALITIES.map((nat) => (
            <option key={nat} value={nat}>
              {nat === '国籍不明' ? '国籍不明 (Chưa rõ)' : nat}
            </option>
          ))}
        </select>
        <select
          value={statusFilter}
          onChange={(e) => setStatusFilter(e.target.value)}
          className="rounded-lg border border-zinc-800 bg-zinc-950 px-3 py-1.5 text-xs text-zinc-200 focus:border-emerald-500 focus:outline-none"
        >
          <option value="all">Tất cả trạng thái ({records.length})</option>
          {RESIGNATION_STATUSES.map((st) => (
            <option key={st} value={st}>
              {st}
            </option>
          ))}
        </select>
      </div>

      {/* Cards List */}
      {filteredRecords.length === 0 ? (
        <div className="rounded-xl border border-dashed border-zinc-800 bg-[#141417]/60 p-8 text-center text-xs text-zinc-400">
          Chưa có hồ sơ nghỉ việc nào khớp bộ lọc. Nhấn &quot;Thêm người nghỉ việc&quot; hoặc &quot;Nhập Excel / CSV hàng loạt&quot;.
        </div>
      ) : (
        <div className="space-y-3">
          {filteredRecords.map((rec) => {
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
                ? 'border-rose-500/40 bg-rose-500/10 text-rose-300'
                : deadlineInfo.urgency === 'warn7'
                ? 'border-amber-500/40 bg-amber-500/10 text-amber-300'
                : deadlineInfo.urgency === 'done'
                ? 'border-emerald-500/40 bg-emerald-500/10 text-emerald-300'
                : 'border-zinc-700 bg-zinc-800/60 text-zinc-300';

            return (
              <div
                key={rec.id}
                className="rounded-xl border border-zinc-800 bg-[#141417] p-4 transition hover:border-zinc-700"
              >
                <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
                  <div className="space-y-1.5">
                    <div className="flex flex-wrap items-center gap-2">
                      <h3 className="text-sm font-bold text-zinc-100">{rec.fullName}</h3>
                      <select
                        aria-label={`Quốc tịch của ${rec.fullName}`}
                        value={currentNat}
                        onChange={(e) =>
                          void handleQuickNationalityChange(rec, e.target.value)
                        }
                        className={`rounded border px-1.5 py-0.5 text-[11px] font-medium focus:border-emerald-500 focus:outline-none ${
                          currentNat === '国籍不明'
                            ? 'border-amber-500/50 bg-amber-500/15 text-amber-300'
                            : 'border-zinc-700 bg-zinc-800 text-zinc-200'
                        }`}
                      >
                        {isCustomNat && <option value={currentNat}>{currentNat}</option>}
                        {RESIGNATION_NATIONALITIES.map((nat) => (
                          <option key={nat} value={nat}>
                            {nat === '国籍不明' ? '国籍不明 (Chưa rõ)' : nat}
                          </option>
                        ))}
                      </select>
                      {showNatSuggestion && suggestedNat && (
                        <button
                          type="button"
                          onClick={() => void handleQuickNationalityChange(rec, suggestedNat)}
                          title="Nhận diện theo tên — bấm để sửa nhanh và khóa bảo vệ GAS"
                          className="inline-flex items-center gap-1 rounded border border-sky-500/40 bg-sky-500/15 px-1.5 py-0.5 text-[10px] font-semibold text-sky-300 hover:bg-sky-500/25"
                        >
                          <Sparkles className="h-2.5 w-2.5" />
                          Gợi ý: {suggestedNat} (Sửa nhanh)
                        </button>
                      )}
                      {rec.jlpt && (
                        <span className="rounded bg-emerald-500/15 px-1.5 py-0.5 text-[11px] font-semibold text-emerald-300">
                          {rec.jlpt}
                        </span>
                      )}
                      {isProtected && (
                        <span
                          title={`Trường đã chỉnh thủ công không bị GAS ghi đè: ${(
                            rec.manualProtectedFields || ['nationality']
                          ).join(', ')}`}
                          className="inline-flex items-center gap-1 rounded border border-emerald-500/30 bg-emerald-500/10 px-1.5 py-0.5 text-[10px] text-emerald-300"
                        >
                          <Lock className="h-2.5 w-2.5" />
                          Bảo vệ GAS
                        </span>
                      )}
                      <span
                        className={`inline-flex items-center gap-1 rounded border px-2 py-0.5 text-[11px] font-semibold ${urgencyClass}`}
                      >
                        {(deadlineInfo.urgency === 'overdue' ||
                          deadlineInfo.urgency === 'urgent3' ||
                          deadlineInfo.urgency === 'warn7') && (
                          <AlertTriangle className="h-3 w-3" />
                        )}
                        {deadlineInfo.urgency === 'done' && (
                          <CheckCircle2 className="h-3 w-3" />
                        )}
                        {deadlineInfo.label}
                      </span>
                    </div>

                    <div className="grid grid-cols-1 gap-x-4 gap-y-1 text-xs text-zinc-300 sm:grid-cols-2">
                      <div>
                        <span className="text-zinc-500">Công ty cũ: </span>
                        <span className="font-medium text-zinc-100">{rec.previousCompany}</span>
                      </div>
                      <div>
                        <span className="text-zinc-500">Ngày nghỉ việc (退職日): </span>
                        <span className="font-medium text-zinc-100">
                          {rec.resignationDate || 'Chưa cập nhật'}
                        </span>
                      </div>
                      <div>
                        <span className="text-zinc-500">Hạn giấy tờ (提出期限): </span>
                        <span className="font-medium text-amber-300">
                          {rec.paperworkDeadline || 'Chưa đặt hạn'}
                        </span>
                      </div>
                      <div>
                        <span className="text-zinc-500">Liên hệ / Ngành: </span>
                        <span className="text-zinc-200">
                          {[rec.phone, rec.fbName, rec.industry].filter(Boolean).join(' • ') ||
                            'Chưa cập nhật'}
                        </span>
                      </div>
                    </div>

                    {rec.handoverNotes && (
                      <div className="rounded-lg border border-zinc-800/80 bg-zinc-950/70 p-2 text-xs text-zinc-300">
                        <span className="font-semibold text-zinc-400">Ghi chú bàn giao: </span>
                        <span className="whitespace-pre-wrap">{rec.handoverNotes}</span>
                      </div>
                    )}

                    {rec.newJobNeed && (
                      <div className="rounded-lg border border-emerald-500/20 bg-emerald-950/15 p-2 text-xs text-emerald-200">
                        <span className="font-semibold text-emerald-400">
                          Nhu cầu chuyển việc mới:{' '}
                        </span>
                        <span className="whitespace-pre-wrap">{rec.newJobNeed}</span>
                      </div>
                    )}
                  </div>

                  {/* Right Action Controls */}
                  <div className="flex shrink-0 flex-col items-stretch gap-2 sm:w-56">
                    <select
                      value={rec.paperworkStatus}
                      onChange={(e) => onStatusChange(rec, e.target.value)}
                      className="rounded-lg border border-zinc-700 bg-zinc-900 px-2.5 py-1.5 text-xs font-medium text-zinc-100 focus:border-emerald-500 focus:outline-none"
                    >
                      {RESIGNATION_STATUSES.map((st) => (
                        <option key={st} value={st}>
                          Trạng thái: {st}
                        </option>
                      ))}
                    </select>

                    <div className="flex items-center gap-1.5">
                      <button
                        type="button"
                        onClick={() => onSyncCalendarSingle(rec)}
                        disabled={isSyncing || rec.paperworkStatus === 'Hoàn tất'}
                        className={`flex flex-1 items-center justify-center gap-1 rounded-lg border px-2 py-1.5 text-[11px] font-medium transition disabled:opacity-45 ${
                          hasCalendarLink
                            ? 'border-emerald-500/40 bg-emerald-500/15 text-emerald-300 hover:bg-emerald-500/25'
                            : 'border-sky-500/40 bg-sky-500/15 text-sky-300 hover:bg-sky-500/25'
                        }`}
                      >
                        {isSyncing ? (
                          <Loader2 className="h-3 w-3 animate-spin" />
                        ) : hasCalendarLink ? (
                          <CalendarCheck className="h-3 w-3" />
                        ) : (
                          <Calendar className="h-3 w-3" />
                        )}
                        {hasCalendarLink ? 'Cập nhật Lịch GG' : 'Đồng bộ Lịch Google'}
                      </button>

                      {isValidGoogleCalendarUrl(rec.googleCalendarEventLink) && (
                        <a
                          href={rec.googleCalendarEventLink}
                          target="_blank"
                          rel="noopener noreferrer"
                          title="Mở sự kiện trên Google Calendar"
                          className="inline-flex items-center justify-center rounded-lg border border-zinc-700 bg-zinc-800 p-1.5 text-zinc-300 hover:text-white"
                        >
                          <ExternalLink className="h-3.5 w-3.5" />
                        </a>
                      )}
                    </div>

                    <button
                      type="button"
                      onClick={() => setReminderTarget(rec)}
                      className="inline-flex items-center justify-center gap-1.5 rounded-lg border border-amber-500/40 bg-amber-500/10 px-2.5 py-1.5 text-[11px] font-medium text-amber-300 hover:bg-amber-500/20"
                    >
                      <MessageSquare className="h-3 w-3" />
                      Tạo tin nhắn nhắc nộp giấy tờ
                    </button>

                    <button
                      type="button"
                      onClick={() => onConvertToCandidate(rec)}
                      className="inline-flex items-center justify-center gap-1.5 rounded-lg border border-emerald-500/40 bg-emerald-500/15 px-2.5 py-1.5 text-[11px] font-semibold text-emerald-300 hover:bg-emerald-500/25"
                    >
                      <UserPlus className="h-3 w-3" />
                      {rec.convertedCandidateId
                        ? 'Đã chuyển (Tạo lại UV mới)'
                        : 'Chuyển sang Ứng viên tìm việc mới'}
                    </button>

                    <div className="flex items-center justify-end gap-1.5 pt-1">
                      <button
                        type="button"
                        onClick={() => openEditModal(rec)}
                        className="inline-flex items-center gap-1 rounded border border-zinc-700 bg-zinc-800 px-2 py-1 text-[11px] text-zinc-200 hover:bg-zinc-700"
                      >
                        <Edit3 className="h-3 w-3" />
                        Sửa
                      </button>
                      <button
                        type="button"
                        onClick={() => onDeleteRecord(rec)}
                        className="inline-flex items-center gap-1 rounded border border-rose-500/30 bg-rose-500/10 px-2 py-1 text-[11px] text-rose-300 hover:bg-rose-500/20"
                      >
                        <Trash2 className="h-3 w-3" />
                        Xóa
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
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/75 p-4">
          <div className="max-h-[90vh] w-full max-w-2xl overflow-y-auto rounded-xl border border-zinc-700 bg-[#141417] p-5 shadow-2xl">
            <div className="mb-4 flex items-center justify-between border-b border-zinc-800 pb-3">
              <h3 className="text-sm font-bold text-zinc-100">
                {editingRecord ? 'Cập nhật hồ sơ nghỉ việc' : 'Thêm hồ sơ Người nghỉ việc'}
              </h3>
              <button
                type="button"
                onClick={() => setIsModalOpen(false)}
                className="rounded p-1 text-zinc-400 hover:bg-zinc-800 hover:text-white"
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            <form onSubmit={handleSubmitForm} className="space-y-3 text-xs">
              <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
                <div>
                  <label className="mb-1 block text-zinc-400">Họ và tên *</label>
                  <input
                    required
                    type="text"
                    value={formState.fullName}
                    onChange={(e) => setFormState((s) => ({ ...s, fullName: e.target.value }))}
                    placeholder="Ví dụ: LE THI BICH PHUONG"
                    className="w-full rounded-lg border border-zinc-700 bg-zinc-950 px-3 py-2 text-zinc-100 focus:border-emerald-500 focus:outline-none"
                  />
                </div>
                <div>
                  <label className="mb-1 block text-zinc-400">Quốc tịch (国籍)</label>
                  <select
                    value={normalizeResignationNationality(formState.nationality)}
                    onChange={(e) =>
                      setFormState((s) => ({ ...s, nationality: e.target.value }))
                    }
                    className="w-full rounded-lg border border-zinc-700 bg-zinc-950 px-3 py-2 text-zinc-100 focus:border-emerald-500 focus:outline-none"
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
                        {nat === '国籍不明' ? '国籍不明 (Chưa rõ quốc tịch)' : nat}
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
                          className="mt-1 inline-flex items-center gap-1 text-[11px] text-sky-400 hover:text-sky-300"
                        >
                          <Sparkles className="h-3 w-3" />
                          Gợi ý theo tên: chọn {suggested}
                        </button>
                      );
                    }
                    return null;
                  })()}
                </div>
                <div className="sm:col-span-2">
                  <label className="mb-1 block text-zinc-400">
                    Công ty cũ đang làm việc *
                  </label>
                  <input
                    required
                    type="text"
                    value={formState.previousCompany}
                    onChange={(e) =>
                      setFormState((s) => ({ ...s, previousCompany: e.target.value }))
                    }
                    placeholder="Tên công ty / cơ sở sắp nghỉ việc"
                    className="w-full rounded-lg border border-zinc-700 bg-zinc-950 px-3 py-2 text-zinc-100 focus:border-emerald-500 focus:outline-none"
                  />
                </div>
                <div>
                  <label className="mb-1 block text-zinc-400">
                    Ngày nghỉ việc chính thức (退職日)
                  </label>
                  <input
                    type="date"
                    value={formState.resignationDate}
                    onChange={(e) =>
                      setFormState((s) => ({ ...s, resignationDate: e.target.value }))
                    }
                    className="w-full rounded-lg border border-zinc-700 bg-zinc-950 px-3 py-2 text-zinc-100 focus:border-emerald-500 focus:outline-none"
                  />
                </div>
                <div>
                  <label className="mb-1 block text-zinc-400">
                    Hạn chót làm &amp; nộp giấy tờ (書類提出期限)
                  </label>
                  <input
                    type="date"
                    value={formState.paperworkDeadline}
                    onChange={(e) =>
                      setFormState((s) => ({ ...s, paperworkDeadline: e.target.value }))
                    }
                    className="w-full rounded-lg border border-zinc-700 bg-zinc-950 px-3 py-2 text-zinc-100 focus:border-emerald-500 focus:outline-none"
                  />
                </div>
                <div>
                  <label className="mb-1 block text-zinc-400">Trạng thái giấy tờ</label>
                  <select
                    value={formState.paperworkStatus}
                    onChange={(e) =>
                      setFormState((s) => ({ ...s, paperworkStatus: e.target.value }))
                    }
                    className="w-full rounded-lg border border-zinc-700 bg-zinc-950 px-3 py-2 text-zinc-100 focus:border-emerald-500 focus:outline-none"
                  >
                    {RESIGNATION_STATUSES.map((st) => (
                      <option key={st} value={st}>
                        {st}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="mb-1 block text-zinc-400">Ngành / JLPT (phục vụ chuyển UV)</label>
                  <div className="flex gap-2">
                    <select
                      value={formState.industry || ''}
                      onChange={(e) =>
                        setFormState((s) => ({ ...s, industry: e.target.value }))
                      }
                      className="flex-1 rounded-lg border border-zinc-700 bg-zinc-950 px-2 py-2 text-zinc-100"
                    >
                      <option value="">-- Ngành --</option>
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
                      className="w-24 rounded-lg border border-zinc-700 bg-zinc-950 px-2 py-2 text-zinc-100"
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
                  <label className="mb-1 block text-zinc-400">Số điện thoại</label>
                  <input
                    type="text"
                    value={formState.phone || ''}
                    onChange={(e) => setFormState((s) => ({ ...s, phone: e.target.value }))}
                    placeholder="080-xxxx-xxxx"
                    className="w-full rounded-lg border border-zinc-700 bg-zinc-950 px-3 py-2 text-zinc-100"
                  />
                </div>
                <div>
                  <label className="mb-1 block text-zinc-400">Tên Facebook</label>
                  <input
                    type="text"
                    value={formState.fbName || ''}
                    onChange={(e) => setFormState((s) => ({ ...s, fbName: e.target.value }))}
                    placeholder="Tên hiển thị Messenger"
                    className="w-full rounded-lg border border-zinc-700 bg-zinc-950 px-3 py-2 text-zinc-100"
                  />
                </div>
              </div>

              <div>
                <label className="mb-1 block text-zinc-400">Ghi chú bàn giao &amp; giấy tờ</label>
                <textarea
                  rows={3}
                  value={formState.handoverNotes || ''}
                  onChange={(e) =>
                    setFormState((s) => ({ ...s, handoverNotes: e.target.value }))
                  }
                  placeholder="Ví dụ: Đang chờ công ty cũ gửi 退職証明書, 源泉徴収票, 離職票; đã nhắc trả chìa khóa KTX..."
                  className="w-full rounded-lg border border-zinc-700 bg-zinc-950 px-3 py-2 text-zinc-100 focus:border-emerald-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="mb-1 block text-zinc-400">Nhu cầu chuyển việc mới</label>
                <textarea
                  rows={2}
                  value={formState.newJobNeed || ''}
                  onChange={(e) =>
                    setFormState((s) => ({ ...s, newJobNeed: e.target.value }))
                  }
                  placeholder="Ví dụ: Tìm đơn Kaigo lương thực nhận >= 17 man tại Saitama / Tochigi..."
                  className="w-full rounded-lg border border-zinc-700 bg-zinc-950 px-3 py-2 text-zinc-100 focus:border-emerald-500 focus:outline-none"
                />
              </div>

              <div className="flex justify-end gap-2 border-t border-zinc-800 pt-3">
                <button
                  type="button"
                  onClick={() => setIsModalOpen(false)}
                  className="rounded-lg border border-zinc-700 px-3 py-1.5 text-zinc-300 hover:bg-zinc-800"
                >
                  Hủy
                </button>
                <button
                  type="submit"
                  className="rounded-lg bg-emerald-500 px-4 py-1.5 font-semibold text-zinc-950 hover:bg-emerald-400"
                >
                  Lưu &amp; Đồng bộ nhắc lịch
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Excel/CSV Preview Table Modal before Bulk Import */}
      {isImportModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 p-4">
          <div className="flex max-h-[88vh] w-full max-w-5xl flex-col rounded-xl border border-zinc-700 bg-[#141417] p-5 shadow-2xl">
            <div className="mb-3 flex items-center justify-between border-b border-zinc-800 pb-3">
              <div>
                <h3 className="text-sm font-bold text-zinc-100">
                  Xem trước dữ liệu nhập từ Excel / CSV ({importSheetName})
                </h3>
                <p className="text-xs text-zinc-400">
                  Đã tự động ánh xạ {importPreviewRows.length} dòng. Kiểm tra trước khi bấm Xác nhận nhập hàng loạt.
                </p>
              </div>
              <button
                type="button"
                onClick={() => setIsImportModalOpen(false)}
                className="rounded p-1 text-zinc-400 hover:bg-zinc-800"
              >
                <X className="h-4 w-4" />
              </button>
            </div>

            <div className="flex-1 overflow-auto rounded-lg border border-zinc-800">
              <table className="w-full border-collapse text-left text-xs">
                <thead className="sticky top-0 bg-zinc-900 text-zinc-300">
                  <tr>
                    <th className="border-b border-zinc-800 px-2.5 py-2">#</th>
                    <th className="border-b border-zinc-800 px-2.5 py-2">Họ tên</th>
                    <th className="border-b border-zinc-800 px-2.5 py-2">Quốc tịch</th>
                    <th className="border-b border-zinc-800 px-2.5 py-2">Công ty cũ</th>
                    <th className="border-b border-zinc-800 px-2.5 py-2">Ngày nghỉ (退職日)</th>
                    <th className="border-b border-zinc-800 px-2.5 py-2">Hạn giấy tờ</th>
                    <th className="border-b border-zinc-800 px-2.5 py-2">Trạng thái</th>
                    <th className="border-b border-zinc-800 px-2.5 py-2">Ghi chú / Nhu cầu</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-zinc-800/80 text-zinc-200">
                  {importPreviewRows.map((row, idx) => {
                    const rowNat = normalizeResignationNationality(row.nationality);
                    return (
                      <tr key={row.id} className="hover:bg-zinc-900/50">
                        <td className="px-2.5 py-2 text-zinc-500">{row._rowIndex}</td>
                        <td className="px-2.5 py-2 font-semibold">{row.fullName}</td>
                        <td className="px-2.5 py-2">
                          <select
                            aria-label={`Quốc tịch dòng ${row._rowIndex}`}
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
                            className={`rounded border px-1.5 py-1 text-[11px] focus:border-emerald-500 focus:outline-none ${
                              rowNat === '国籍不明'
                                ? 'border-amber-500/50 bg-amber-500/15 text-amber-300'
                                : 'border-zinc-700 bg-zinc-950 text-zinc-100'
                            }`}
                          >
                            {RESIGNATION_NATIONALITIES.map((nat) => (
                              <option key={nat} value={nat}>
                                {nat === '国籍不明' ? '国籍不明 (Chưa rõ)' : nat}
                              </option>
                            ))}
                          </select>
                        </td>
                        <td className="px-2.5 py-2">{row.previousCompany}</td>
                        <td className="px-2.5 py-2">{row.resignationDate || '-'}</td>
                        <td className="px-2.5 py-2 text-amber-300">
                          {row.paperworkDeadline || '-'}
                        </td>
                        <td className="px-2.5 py-2">{row.paperworkStatus}</td>
                        <td className="max-w-xs px-2.5 py-2 text-[11px] text-zinc-400">
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
                className="rounded-lg border border-zinc-700 px-3 py-1.5 text-xs text-zinc-300 hover:bg-zinc-800"
              >
                Đóng
              </button>
              <button
                type="button"
                disabled={isImporting}
                onClick={handleConfirmImport}
                className="inline-flex items-center gap-1.5 rounded-lg bg-emerald-500 px-4 py-1.5 text-xs font-semibold text-zinc-950 hover:bg-emerald-400 disabled:opacity-50"
              >
                {isImporting && <Loader2 className="h-3.5 w-3.5 animate-spin" />}
                Xác nhận nhập {importPreviewRows.length} hồ sơ
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Reminder Message Modal */}
      {reminderTarget && reminderMsgs && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/75 p-4">
          <div className="w-full max-w-xl rounded-xl border border-zinc-700 bg-[#141417] p-5 shadow-2xl">
            <div className="mb-3 flex items-center justify-between border-b border-zinc-800 pb-2.5">
              <h3 className="text-sm font-bold text-zinc-100">
                Tin nhắn nhắc nộp giấy tờ nghỉ việc — {reminderTarget.fullName}
              </h3>
              <button
                type="button"
                onClick={() => setReminderTarget(null)}
                className="rounded p-1 text-zinc-400 hover:bg-zinc-800"
              >
                <X className="h-4 w-4" />
              </button>
            </div>
            <div className="space-y-3 text-xs">
              <div className="rounded-lg border border-zinc-800 bg-zinc-950 p-3">
                <div className="mb-1.5 flex items-center justify-between">
                  <span className="font-semibold text-emerald-400">Tiếng Việt</span>
                  <button
                    type="button"
                    onClick={async () => {
                      await copyTextSafely(reminderMsgs.vi);
                      onNotify('Đã sao chép tin nhắn nhắc giấy tờ (Tiếng Việt)!', 'success');
                    }}
                    className="inline-flex items-center gap-1 rounded bg-emerald-500/20 px-2 py-1 text-[11px] text-emerald-300 hover:bg-emerald-500/30"
                  >
                    <Copy className="h-3 w-3" />
                    Sao chép
                  </button>
                </div>
                <p className="whitespace-pre-wrap text-zinc-200">{reminderMsgs.vi}</p>
              </div>
              <div className="rounded-lg border border-zinc-800 bg-zinc-950 p-3">
                <div className="mb-1.5 flex items-center justify-between">
                  <span className="font-semibold text-sky-400">日本語</span>
                  <button
                    type="button"
                    onClick={async () => {
                      await copyTextSafely(reminderMsgs.jp);
                      onNotify('Đã sao chép tin nhắn nhắc giấy tờ (Tiếng Nhật)!', 'success');
                    }}
                    className="inline-flex items-center gap-1 rounded bg-sky-500/20 px-2 py-1 text-[11px] text-sky-300 hover:bg-sky-500/30"
                  >
                    <Copy className="h-3 w-3" />
                    Sao chép
                  </button>
                </div>
                <p className="whitespace-pre-wrap text-zinc-200">{reminderMsgs.jp}</p>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
