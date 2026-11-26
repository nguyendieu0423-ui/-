export type ChannelType = 'personal' | 'company' | 'school';

export type MainTabId =
  | 'all'
  | 'personal'
  | 'company'
  | 'school'
  | 'mendan'
  | 'proposed'
  | 'resignation'
  | 'pt'
  | 'kyujin'
  | 'report';

export type CareFilterType =
  | 'all'
  | 'recent7'
  | 'scheduled'
  | 'recentMendan'
  | 'urgent'
  | 'declined';

export type ReplyScenario =
  | 'first_mendan'
  | 'recommend_job'
  | 'remind_interview'
  | 'followup_silent'
  | 'missing_info_request'
  | 'post_interview_followup';

export type CandidateViewMode = 'list' | 'card';

export type ProposedGroupFilter = 'in_progress' | 'naitei' | 'jitai' | 'all';

export type JitaiSubFilter = 'all' | '内定辞退' | '選考中辞退' | '不合格' | '辞退・終了';

export interface Candidate {
  id: string;
  fbSenderId?: string;
  fbProfileUrl?: string;
  name: string;
  fbName?: string;
  furigana?: string;
  channel: ChannelType | string;
  schoolName?: string;
  graduationDate?: string;
  nationality?: string;
  gender?: string;
  age?: string;
  birthDate?: string;
  postalCode?: string;
  jlpt?: string;
  desiredIndustry?: string;
  desiredArea?: string;
  conversionGoal?: string;
  visaStatus?: string;
  visaExpiry?: string;
  passedExam?: string;
  japanExperience?: string;
  salaryExpectation?: string;
  negotiationLog?: string;
  startDate?: string;
  phone?: string;
  email?: string;
  address?: string;
  progressStatus?: string;
  proposedCompany?: string;
  naiteiCompany?: string;
  expectedJoinDate?: string;
  visaSubmissionDeadline?: string;
  mendanDate?: string;
  interviewDate?: string;
  lastContactDate?: string;
  declineReason?: string;
  memo?: string;
  googleCalendarEventId?: string;
  googleCalendarEventLink?: string;
  savedRirekisho?: SavedRirekishoRecord;
  _isSchool?: boolean;
}

export type KyujinEmploymentType =
  | 'Tokutei Ginou (特定技能)'
  | 'Nhân viên chính thức / Chuyển đổi visa (正社員・登用)'
  | 'Baito / Làm thêm (アルバイト・パート)'
  | string;

export interface KyujinSlot {
  slotId: string;
  sector?: string;
  industry?: string;
  employmentType?: KyujinEmploymentType;
  jobCode?: string;
  area?: string;
  shiftTime?: string;
  salary?: string;
  nightShiftAllowance?: string;
  estimatedNetSalary?: string;
  hourlyWage?: string;
  commuteAllowance?: string;
  shiftsPerWeek?: string;
  bonus?: string;
  bonusDetail?: string;
  minJlpt?: string;
  housingSupport?: string;
  recruitmentStatus?: string;
  careQualification?: string;
  facilityType?: string;
  experienceRequirement?: string;
  experienceDetail?: string;
  jobDescription?: string;
  excludedNationalities?: string[];
  excludedGenders?: string[];
  requiresInJapan?: boolean;
  requiresTokutei?: boolean;
  requiredConditions?: string[];
  ngConditions?: string[];
  salaryManuallyEdited?: boolean;
  manualProtectedFields?: string[];
}

export interface KyujinCompany {
  id: string;
  companyName: string;
  sector?: string;
  industry: string;
  employmentType?: KyujinEmploymentType;
  recruitmentStatus: string;
  excludedNationalities?: string[];
  excludedGenders?: string[];
  requiresInJapan?: boolean;
  requiresTokutei?: boolean;
  requiredConditions?: string[];
  ngConditions?: string[];
  areaBreakdown: KyujinSlot[];
}

export interface PtPartner {
  id: string;
  partnerName: string;
  contactPerson: string;
  phone: string;
  email: string;
  status: string;
  specialty: string;
  commissionTerms: string;
  memo: string;
}

export type ResignationPaperworkStatus =
  | 'Chưa làm'
  | 'Đang thu thập'
  | 'Đã gửi công ty'
  | 'Hoàn tất';

export type ResignationProtectedField =
  | 'nationality'
  | 'resignationDate'
  | 'paperworkDeadline';

export interface ResignationRecord {
  id: string;
  fullName: string;
  nationality?: string;
  previousCompany: string;
  resignationDate: string; // YYYY-MM-DD (退職日)
  paperworkDeadline: string; // YYYY-MM-DD (書類提出期限)
  paperworkStatus: ResignationPaperworkStatus | string;
  handoverNotes?: string;
  newJobNeed?: string;
  phone?: string;
  fbName?: string;
  fbProfileUrl?: string;
  industry?: string;
  jlpt?: string;
  convertedCandidateId?: string;
  googleCalendarEventId?: string;
  googleCalendarEventLink?: string;
  manualProtectedFields?: ResignationProtectedField[];
  gasProtected?: boolean;
  updatedAt?: string;
}

export type DailyFocusCompletionMap = Record<string, boolean>;

export interface DailyFocusItem {
  id: string;
  category: 'mendan' | 'uncontacted' | 'visa_expiry' | 'resignation' | 'visa_submission';
  urgency: 'critical' | 'warning';
  title: string;
  subtitle: string;
  targetTab: MainTabId;
  entityId?: string;
}

export interface RirekishoRow {
  year: string;
  month: string;
  content: string;
}

export interface RirekishoDraft {
  dateStr: string;
  furigana: string;
  fullName: string;
  birthDate: string;
  age: string;
  gender: string;
  postalCode: string;
  address: string;
  phone: string;
  email: string;
  nationality: string;
  visaExpiry: string;
  visaStatus: string;
  educationRows: RirekishoRow[];
  workRows: RirekishoRow[];
  licenseRows: RirekishoRow[];
  nativeLanguage: string;
  availableLanguages: string;
  caSpeaking: string;
  caListening: string;
  motivation: string;
  selfPr: string;
  personalRequests: string;
  allergy: string;
  dependentsCount: string;
  spouse: string;
  spouseSupport: string;
  photoDataUrl?: string;
}

export interface SavedRirekishoRecord {
  draft: RirekishoDraft;
  sourceFileName?: string;
  exportedFileName?: string;
  exportedAt?: string;
  updatedAt: string;
}

export interface MissingInfoItem {
  key: string;
  ja: string;
  vi: string;
  en: string;
}

export interface KyujinMatchResult {
  score: number;
  reasons: string[];
  company: KyujinCompany;
  slot: KyujinSlot;
  industryMatched?: boolean;
  areaMatched?: boolean;
}

export interface KyujinMatchEvaluation {
  matches: KyujinMatchResult[];
  hasCandidateIndustry: boolean;
  hasCandidateArea: boolean;
  hasExactIndustryAndAreaMatch: boolean;
  noticeMessage: string;
}

export interface ExcelAnchor {
  c1: number;
  r1: number;
  c2: number;
  r2: number;
}

