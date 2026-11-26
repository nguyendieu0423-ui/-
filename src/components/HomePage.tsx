import React, { useState } from 'react';
import {
  ArrowRight,
  Building2,
  FileSpreadsheet,
  MessageSquareText,
  ShieldCheck,
  Users,
} from 'lucide-react';
import type { MainTabId } from '../types/crm';
import type { UiLang } from '../i18n/crmTranslations';

interface HomePageProps {
  lang: UiLang;
  onChangeLang: (lang: UiLang) => void;
  candidateCount: number;
  schoolCount: number;
  companyCount: number;
  slotCount: number;
  ptCount: number;
  upcomingMendanCount: number;
  userEmail: string | null;
  authLoading: boolean;
  onStartWorkspace: (targetTab?: MainTabId) => void;
  onSignInGoogle: () => void;
  onSignOut: () => void;
  onOpenChatbot: () => void;
}

const HERO_IMAGE_PATH = '/src/assets/images/hero_tokutei_office_1795481152089.jpg';

const HOME_COPY = {
  vi: {
    navAria: 'Điều hướng trang chủ',
    navOverview: 'Giới thiệu',
    navCapabilities: 'Tính năng',
    navWorkflow: 'Quy trình',
    navAi: 'Trợ lý AI',
    startBtn: 'Bắt đầu',
    badgeRegion: 'Nhật Bản',
    heroTitle: 'Quản lý ứng viên Kỹ năng Đặc định từ tin nhắn đầu tiên đến khi nhận việc.',
    heroDesc:
      'Tổng hợp ứng viên từ Facebook cá nhân, Fanpage công ty và trường Nhật ngữ trong một không gian điều phối mật độ cao. Trích xuất phiếu tuyển dụng bằng AI, tạo Rirekisho chuẩn Nhật và đồng bộ bảo mật trên Cloud SQL.',
    startNowBtn: 'Bắt đầu ngay',
    signedInLabel: 'Đã đăng nhập:',
    signOut: 'Đăng xuất',
    connecting: 'Đang kết nối...',
    signInGoogle: 'Sign in with Google',
    stripItems: [
      'Tiếp nhận đa kênh FB',
      'Rirekisho 2 cột',
      'Trích xuất Kyujin PDF & Excel',
      'Đồng bộ Cloud SQL',
    ],
    heroFallbackTitle: 'Không gian điều phối tuyển dụng Tokutei Ginou',
    heroFallbackDesc: 'Quản lý hồ sơ ứng viên, lịch Mendan và phiếu tuyển dụng tập trung.',
    heroImgAlt: 'Văn phòng tư vấn tuyển dụng Tokutei Ginou tại Nhật Bản',
    metricCand: 'Ứng viên & Tư vấn FB',
    metricSchool: 'Du học sinh trường Nhật',
    metricJobs: 'Đơn hàng đang mở',
    metricSlots: (n: number) => `(${n} vị trí)`,
    metricMendan: 'Lịch Mendan & Đối tác PT',
    capTitle: 'Mọi công cụ điều phối trong một bàn làm việc thống nhất',
    capDesc:
      'Thiết kế 3 cột sáng rõ giúp xử lý danh sách ứng viên, ghép đơn hàng và theo dõi cảnh báo khẩn trong cùng một màn hình.',
    card1Tag: '01. Quản lý ứng viên & Cảnh báo Mendan',
    card1Title: 'Theo dõi ứng viên FB cá nhân, Fanpage và Trường Nhật ngữ',
    card1Desc:
      'Tự động bóc tách đoạn chat Facebook thành JLPT, SĐT, hạn thẻ ngoại kiều, ngành nghề và khu vực mong muốn. Bảng cảnh báo bên phải luôn hiển thị lịch Mendan sắp tới và hồ sơ thiếu thông tin.',
    card1Cta: 'Mở danh sách ứng viên',
    card2Tag: '02. Phiếu tuyển dụng',
    card2Title: 'Đọc Excel, PDF và khớp đơn hàng tự động',
    card2Desc:
      'Tải lên nhiều phiếu tuyển dụng cùng lúc để trích xuất mức lương, thưởng, ký túc xá và điều kiện JLPT. Xuất bản tóm tắt ẩn danh tiếng Việt hoặc tiếng Nhật.',
    card2Cta: 'Quản lý đơn hàng',
    card3Tag: '03. Rirekisho & Soạn tin',
    card3Title: 'Sơ yếu lý lịch 2 cột & Tin nhắn đa ngôn ngữ',
    card3Desc:
      'Điền tự động Rirekisho từ file hoặc đoạn hội thoại, in hoặc xuất HTML độc lập. Tạo nhanh mẫu nhắc lịch bằng tiếng Việt, Nhật, Anh.',
    card3Cta: 'Xem tiến độ tiến cử',
    card4Tag: '04. Trợ lý AI Gemini & Cloud SQL PostgreSQL',
    card4Title: 'Hội thoại AI nhiều lượt theo vai trò & Lưu trữ bền vững',
    card4Desc:
      'Chuyển đổi linh hoạt giữa Tư vấn chiến lược CRM, Soạn phản hồi nhanh (Việt/Nhật/Anh) và Cố vấn sàng lọc Mendan. Khi đăng nhập tài khoản Google, toàn bộ hồ sơ ứng viên, đơn hàng và lịch sử trò chuyện được đồng bộ lên Cloud SQL.',
    card4CtaAi: 'Mở Trợ lý AI Gemini',
    card4CtaReport: 'Xem báo cáo phân tích',
    workflowTitle: 'Quy trình làm việc tinh gọn cho điều phối viên',
    step1Label: 'Bước 1',
    step1Title: 'Nhập đoạn chat Messenger hoặc danh sách du học sinh',
    step1Desc:
      'Trích xuất tự động thông tin cá nhân, trình độ tiếng Nhật và khu vực làm việc mong muốn.',
    step2Label: 'Bước 2',
    step2Title: 'Đối chiếu phiếu tuyển dụng và hoàn thiện Rirekisho 2 cột',
    step2Desc:
      'Xem bảng gợi ý job ngay trong thẻ ứng viên và tạo hồ sơ chuẩn Nhật chỉ trong vài phút.',
    step3Label: 'Bước 3',
    step3Title: 'Xử lý cảnh báo khẩn và đồng bộ dữ liệu Cloud SQL',
    step3Desc:
      'Đăng nhập Google để lưu trữ lâu dài toàn bộ ứng viên, đơn hàng và lịch sử trò chuyện AI giữa các phiên.',
    readyTitle: 'Sẵn sàng bắt đầu phiên điều phối?',
    readyDesc:
      'Truy cập trực tiếp giao diện 3 cột sáng rõ với dữ liệu hiện tại hoặc đăng nhập Google để đồng bộ cơ sở dữ liệu PostgreSQL cá nhân.',
    enterWorkspaceBtn: 'Vào không gian làm việc CRM',
    chatAiBtn: 'Trò chuyện với Trợ lý AI Gemini',
    footerDesc:
      'Hệ thống quản lý ứng viên, phiếu tuyển dụng và đối tác Kỹ năng Đặc định Nhật Bản.',
    footerCand: 'Ứng viên',
    footerJobs: 'Tin tuyển dụng',
    footerReport: 'Báo cáo',
    footerAi: 'Trợ lý AI',
  },
  jp: {
    navAria: 'ホームナビゲーション',
    navOverview: '概要',
    navCapabilities: '機能一覧',
    navWorkflow: '業務フロー',
    navAi: 'AIアシスタント',
    startBtn: '開く',
    badgeRegion: '日本国内向け',
    heroTitle: '特定技能の求職者を初回チャットから内定・入社まで一元管理。',
    heroDesc:
      '個人Facebook、会社公式Fanpage、日本語学校からの相談者をひとつの高密度3カラム画面に統合。AIによる求人票抽出、2カラム履歴書作成、Cloud SQL連携までスムーズに対応します。',
    startNowBtn: '今すぐ開始',
    signedInLabel: 'ログイン中:',
    signOut: 'ログアウト',
    connecting: '接続中...',
    signInGoogle: 'Googleでログイン',
    stripItems: [
      'FBマルチチャネル受付',
      '2カラム履歴書生成',
      'PDF・Excel求人一括読込',
      'Cloud SQL自動同期',
    ],
    heroFallbackTitle: '特定技能人材紹介 統合CRMワークスペース',
    heroFallbackDesc: '求職者管理・面談日程・求人マッチングを効率的に一元管理。',
    heroImgAlt: '日本の特定技能人材紹介オフィスイメージ',
    metricCand: 'FB求職者・相談者',
    metricSchool: '日本語学校・留学生',
    metricJobs: '登録求人企業数',
    metricSlots: (n: number) => `(${n}枠)`,
    metricMendan: '直近面談・PT提携先',
    capTitle: '紹介コーディネーターの全業務を1つの画面に集約',
    capDesc:
      '見やすいライトモードの3カラム構成で、求職者リスト、求人マッチング、要至急アラートを同時に処理できます。',
    card1Tag: '01. 求職者管理 & 面談アラート',
    card1Title: '個人FB・会社Fanpage・日本語学校からの相談者を横断管理',
    card1Desc:
      'Facebookの会話からJLPT・連絡先・在留期限・希望分野とエリアを自動抽出。右側のシステムアラートで直近の面談予定や情報不足を即座に把握できます。',
    card1Cta: '求職者一覧を開く',
    card2Tag: '02. 求人票管理',
    card2Title: 'Excel・PDFを自動読込し、最適求人を高精度マッチング',
    card2Desc:
      '複数ファイルを一括アップロードして給与・賞与・寮情報・JLPT条件を抽出。ベトナム語・日本語の匿名紹介文もワンクリックで作成できます。',
    card2Cta: '求人一覧を管理',
    card3Tag: '03. 履歴書 & 返信テンプレート',
    card3Title: '2カラム標準履歴書 & 多言語返信ジェネレーター',
    card3Desc:
      'ファイルやチャット内容から履歴書を自動作成し、HTML保存・印刷に対応。ベトナム語・日本語・英語のリマインド文も即座に生成します。',
    card3Cta: '推薦・選考状況を見る',
    card4Tag: '04. Gemini AI & Cloud SQL PostgreSQL',
    card4Title: '役割別AIチャット支援 & クラウド永続保存',
    card4Desc:
      'CRM戦略相談、日越英の返信添削、面談スクリーニング助言を切り替え可能。Googleアカウントでログインすると全データと会話履歴がCloud SQLに安全に保存されます。',
    card4CtaAi: 'Gemini AIを開く',
    card4CtaReport: '分析レポートを見る',
    workflowTitle: 'コーディネーター向け 3ステップ業務フロー',
    step1Label: 'ステップ 1',
    step1Title: 'Messengerチャットや留学生リストを取り込む',
    step1Desc: '基本情報・日本語レベル・希望勤務地を自動で構造化して登録します。',
    step2Label: 'ステップ 2',
    step2Title: '最適求人を比較し、2カラム履歴書を完成させる',
    step2Desc: 'カード内のマッチング候補を確認し、企業提出用の履歴書を数分で出力できます。',
    step3Label: 'ステップ 3',
    step3Title: '要至急フォローを処理し、Cloud SQLへ同期保存',
    step3Desc:
      'Googleログインにより、求職者・求人・退職書類リマインドやAI履歴を永続的に共有できます。',
    readyTitle: 'CRMワークスペースを開始しますか？',
    readyDesc:
      '既存データをそのまま保持した3カラム画面へ移動するか、GoogleログインでPostgreSQL同期を有効にできます。',
    enterWorkspaceBtn: 'CRMワークスペースへ進む',
    chatAiBtn: 'Gemini AIアシスタントと話す',
    footerDesc: '特定技能人材紹介・求人票・提携パートナー統合管理システム',
    footerCand: '求職者一覧',
    footerJobs: '求人一覧',
    footerReport: 'レポート',
    footerAi: 'AIアシスタント',
  },
} as const;

export function HomePage({
  lang,
  onChangeLang,
  candidateCount,
  schoolCount,
  companyCount,
  slotCount,
  ptCount,
  upcomingMendanCount,
  userEmail,
  authLoading,
  onStartWorkspace,
  onSignInGoogle,
  onSignOut,
  onOpenChatbot,
}: HomePageProps) {
  const [heroImgFailed, setHeroImgFailed] = useState(false);
  const c = HOME_COPY[lang];

  return (
    <div className="min-h-screen bg-[#F8FAFC] text-[#0F172A]">
      {/* Top Bar */}
      <header className="sticky top-0 z-30 border-b border-[#E2E8F0] bg-white/95 shadow-[0_1px_3px_rgba(0,0,0,0.05)] backdrop-blur">
        <div className="mx-auto flex max-w-6xl items-center justify-between gap-6 px-6 py-3.5">
          <a
            href="#overview"
            className="text-base font-bold tracking-tight text-[#0F172A] whitespace-nowrap shrink-0"
          >
            Tokutei CRM
          </a>

          <nav
            aria-label={c.navAria}
            className="hidden md:flex items-center gap-6 text-sm font-medium text-[#64748B]"
          >
            <a
              href="#overview"
              className="hover:text-[#0F172A] transition-colors whitespace-nowrap shrink-0"
            >
              {c.navOverview}
            </a>
            <a
              href="#capabilities"
              className="hover:text-[#0F172A] transition-colors whitespace-nowrap shrink-0"
            >
              {c.navCapabilities}
            </a>
            <a
              href="#workflow"
              className="hover:text-[#0F172A] transition-colors whitespace-nowrap shrink-0"
            >
              {c.navWorkflow}
            </a>
            <a
              href="#assistant"
              className="hover:text-[#0F172A] transition-colors whitespace-nowrap shrink-0"
            >
              {c.navAi}
            </a>
          </nav>

          <div className="flex items-center gap-2.5 shrink-0">
            <div
              role="group"
              aria-label="Language Switcher"
              className="inline-flex items-center rounded-md border border-[#E2E8F0] bg-slate-100 p-0.5 shrink-0"
            >
              <button
                type="button"
                onClick={() => onChangeLang('vi')}
                className={`rounded px-2 py-1 text-[11px] font-bold transition-all whitespace-nowrap cursor-pointer ${
                  lang === 'vi'
                    ? 'bg-white text-[#0F172A] shadow-sm'
                    : 'text-[#64748B] hover:text-[#0F172A]'
                }`}
              >
                🇻🇳 VI
              </button>
              <button
                type="button"
                onClick={() => onChangeLang('jp')}
                className={`rounded px-2 py-1 text-[11px] font-bold transition-all whitespace-nowrap cursor-pointer ${
                  lang === 'jp'
                    ? 'bg-white text-[#0F172A] shadow-sm'
                    : 'text-[#64748B] hover:text-[#0F172A]'
                }`}
              >
                🇯🇵 JP
              </button>
            </div>

            <button
              type="button"
              onClick={() => onStartWorkspace('all')}
              className="inline-flex items-center gap-2 rounded-md bg-[#16A34A] px-4 py-2 text-xs font-semibold text-white transition-colors hover:bg-[#15803D] whitespace-nowrap shrink-0 cursor-pointer"
            >
              {c.startBtn}
              <ArrowRight className="h-3.5 w-3.5" aria-hidden="true" />
            </button>
          </div>
        </div>
      </header>

      <main>
        {/* Hero Section */}
        <section id="overview" className="mx-auto max-w-6xl px-6 pt-12 pb-10 md:pt-16 md:pb-12">
          <div className="grid grid-cols-1 gap-10 lg:grid-cols-12 lg:items-stretch">
            <div className="flex flex-col justify-center lg:col-span-6">
              <div className="mb-4 flex flex-wrap items-center gap-2 font-mono text-xs text-[#64748B]">
                <span className="font-semibold text-[#16A34A]">Tokutei Ginou Workstation</span>
                <span aria-hidden="true">·</span>
                <span>{c.badgeRegion}</span>
              </div>

              <h1
                className="font-bold tracking-tight text-[#0F172A] text-balance"
                style={{ fontSize: 'clamp(2rem, 3.6vw, 3.05rem)', lineHeight: 1.15 }}
              >
                {c.heroTitle}
              </h1>

              <p className="mt-5 max-w-xl text-base leading-relaxed text-[#64748B]">
                {c.heroDesc}
              </p>

              <div className="mt-8 flex flex-wrap items-center gap-4">
                <button
                  type="button"
                  onClick={() => onStartWorkspace('all')}
                  className="inline-flex items-center gap-2 rounded-md bg-[#16A34A] px-6 py-3 text-sm font-semibold text-white shadow-sm transition-colors hover:bg-[#15803D] whitespace-nowrap shrink-0 cursor-pointer"
                >
                  {c.startNowBtn}
                  <ArrowRight className="h-4 w-4" aria-hidden="true" />
                </button>

                {userEmail ? (
                  <div className="flex items-center gap-3 text-xs text-[#64748B]">
                    <span className="truncate max-w-[220px]" title={userEmail}>
                      {c.signedInLabel}{' '}
                      <strong className="font-semibold text-[#0F172A]">{userEmail}</strong>
                    </span>
                    <button
                      type="button"
                      onClick={onSignOut}
                      className="font-medium text-[#64748B] underline underline-offset-4 hover:text-[#0F172A] whitespace-nowrap shrink-0 cursor-pointer"
                    >
                      {c.signOut}
                    </button>
                  </div>
                ) : (
                  <button
                    type="button"
                    disabled={authLoading}
                    onClick={onSignInGoogle}
                    className="inline-flex items-center gap-2.5 rounded-md border border-[#CBD5E1] bg-white px-5 py-3 text-sm font-medium text-[#0F172A] shadow-sm transition-colors hover:bg-slate-50 disabled:opacity-60 whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <svg className="h-4 w-4 shrink-0" viewBox="0 0 48 48" aria-hidden="true">
                      <path
                        fill="#EA4335"
                        d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"
                      />
                      <path
                        fill="#4285F4"
                        d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"
                      />
                      <path
                        fill="#FBBC05"
                        d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"
                      />
                      <path
                        fill="#34A853"
                        d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"
                      />
                    </svg>
                    {authLoading ? c.connecting : c.signInGoogle}
                  </button>
                )}
              </div>

              <div className="mt-8 flex flex-wrap items-center gap-x-4 gap-y-2 border-t border-[#E2E8F0] pt-5 font-mono text-[11px] text-[#64748B]">
                {c.stripItems.map((item, idx) => (
                  <React.Fragment key={item}>
                    {idx > 0 && <span aria-hidden="true">·</span>}
                    <span>{item}</span>
                  </React.Fragment>
                ))}
              </div>
            </div>

            <div className="lg:col-span-6">
              <div className="h-full min-h-[280px] overflow-hidden rounded-xl border border-[#E2E8F0] bg-white shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
                {heroImgFailed ? (
                  <div className="flex h-full min-h-[320px] flex-col items-center justify-center bg-white p-8 text-center">
                    <Building2 className="mb-3 h-10 w-10 text-[#64748B]" aria-hidden="true" />
                    <p className="text-sm font-semibold text-[#0F172A]">{c.heroFallbackTitle}</p>
                    <p className="mt-1 max-w-xs text-xs text-[#64748B]">{c.heroFallbackDesc}</p>
                  </div>
                ) : (
                  <img
                    src={HERO_IMAGE_PATH}
                    alt={c.heroImgAlt}
                    referrerPolicy="no-referrer"
                    onError={() => setHeroImgFailed(true)}
                    className="h-full w-full object-cover lg:min-h-[420px]"
                  />
                )}
              </div>
            </div>
          </div>

          {/* Full-width Operational Metrics Strip below Hero Split */}
          <div className="mt-10 grid grid-cols-2 gap-6 rounded-lg border border-[#E2E8F0] bg-white px-6 py-5 shadow-[0_1px_3px_rgba(0,0,0,0.05)] sm:grid-cols-4">
            <div>
              <p className="font-mono text-[10px] tracking-wider text-[#64748B]">{c.metricCand}</p>
              <p className="mt-1 font-mono text-2xl font-bold tracking-tight text-[#0F172A] tabular-nums">
                {candidateCount}
              </p>
            </div>
            <div>
              <p className="font-mono text-[10px] tracking-wider text-[#64748B]">{c.metricSchool}</p>
              <p className="mt-1 font-mono text-2xl font-bold tracking-tight text-[#0F172A] tabular-nums">
                {schoolCount}
              </p>
            </div>
            <div>
              <p className="font-mono text-[10px] tracking-wider text-[#64748B]">{c.metricJobs}</p>
              <p className="mt-1 font-mono text-2xl font-bold tracking-tight text-[#16A34A] tabular-nums">
                {companyCount}{' '}
                <span className="text-xs font-medium text-[#64748B]">
                  {c.metricSlots(slotCount)}
                </span>
              </p>
            </div>
            <div>
              <p className="font-mono text-[10px] tracking-wider text-[#64748B]">{c.metricMendan}</p>
              <p className="mt-1 font-mono text-2xl font-bold tracking-tight text-[#0F172A] tabular-nums">
                {upcomingMendanCount}{' '}
                <span className="text-xs font-medium text-[#64748B]">/ {ptCount} PT</span>
              </p>
            </div>
          </div>
        </section>

        {/* Capabilities Bento Grid */}
        <section id="capabilities" className="mx-auto max-w-6xl px-6 py-12">
          <div className="mb-8 max-w-2xl">
            <h2 className="text-2xl font-bold tracking-tight text-[#0F172A] sm:text-3xl text-balance">
              {c.capTitle}
            </h2>
            <p className="mt-2 text-sm leading-relaxed text-[#64748B]">{c.capDesc}</p>
          </div>

          <div className="grid grid-cols-1 gap-4 lg:grid-cols-3">
            {/* Card 01 */}
            <div className="flex flex-col justify-between rounded-lg border border-[#E2E8F0] bg-white p-6 shadow-[0_1px_3px_rgba(0,0,0,0.05)] lg:col-span-2">
              <div>
                <div className="flex items-center justify-between font-mono text-xs text-[#64748B]">
                  <span>{c.card1Tag}</span>
                  <Users className="h-4 w-4 text-[#16A34A]" aria-hidden="true" />
                </div>
                <h3 className="mt-3 text-lg font-bold text-[#0F172A]">{c.card1Title}</h3>
                <p className="mt-2 text-sm leading-relaxed text-[#64748B]">{c.card1Desc}</p>
              </div>
              <div className="mt-6">
                <button
                  type="button"
                  onClick={() => onStartWorkspace('all')}
                  className="inline-flex items-center gap-1.5 text-xs font-semibold text-[#16A34A] hover:underline cursor-pointer"
                >
                  {c.card1Cta}
                  <ArrowRight className="h-3.5 w-3.5" aria-hidden="true" />
                </button>
              </div>
            </div>

            {/* Card 02 */}
            <div className="flex flex-col justify-between rounded-lg border border-[#E2E8F0] bg-white p-6 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
              <div>
                <div className="flex items-center justify-between font-mono text-xs text-[#64748B]">
                  <span>{c.card2Tag}</span>
                  <FileSpreadsheet className="h-4 w-4 text-[#16A34A]" aria-hidden="true" />
                </div>
                <h3 className="mt-3 text-lg font-bold text-[#0F172A]">{c.card2Title}</h3>
                <p className="mt-2 text-sm leading-relaxed text-[#64748B]">{c.card2Desc}</p>
              </div>
              <div className="mt-6">
                <button
                  type="button"
                  onClick={() => onStartWorkspace('kyujin')}
                  className="inline-flex items-center gap-1.5 text-xs font-semibold text-[#16A34A] hover:underline cursor-pointer"
                >
                  {c.card2Cta}
                  <ArrowRight className="h-3.5 w-3.5" aria-hidden="true" />
                </button>
              </div>
            </div>

            {/* Card 03 */}
            <div className="flex flex-col justify-between rounded-lg border border-[#E2E8F0] bg-white p-6 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
              <div>
                <div className="flex items-center justify-between font-mono text-xs text-[#64748B]">
                  <span>{c.card3Tag}</span>
                  <ShieldCheck className="h-4 w-4 text-[#16A34A]" aria-hidden="true" />
                </div>
                <h3 className="mt-3 text-lg font-bold text-[#0F172A]">{c.card3Title}</h3>
                <p className="mt-2 text-sm leading-relaxed text-[#64748B]">{c.card3Desc}</p>
              </div>
              <div className="mt-6">
                <button
                  type="button"
                  onClick={() => onStartWorkspace('proposed')}
                  className="inline-flex items-center gap-1.5 text-xs font-semibold text-[#16A34A] hover:underline cursor-pointer"
                >
                  {c.card3Cta}
                  <ArrowRight className="h-3.5 w-3.5" aria-hidden="true" />
                </button>
              </div>
            </div>

            {/* Card 04 */}
            <div className="flex flex-col justify-between rounded-lg border border-[#E2E8F0] bg-white p-6 shadow-[0_1px_3px_rgba(0,0,0,0.05)] lg:col-span-2">
              <div>
                <div className="flex items-center justify-between font-mono text-xs text-[#64748B]">
                  <span>{c.card4Tag}</span>
                  <MessageSquareText className="h-4 w-4 text-[#16A34A]" aria-hidden="true" />
                </div>
                <h3 className="mt-3 text-lg font-bold text-[#0F172A]">{c.card4Title}</h3>
                <p className="mt-2 text-sm leading-relaxed text-[#64748B]">{c.card4Desc}</p>
              </div>
              <div className="mt-6 flex flex-wrap items-center gap-4">
                <button
                  type="button"
                  onClick={onOpenChatbot}
                  className="inline-flex items-center gap-1.5 text-xs font-semibold text-[#16A34A] hover:underline cursor-pointer"
                >
                  {c.card4CtaAi}
                  <ArrowRight className="h-3.5 w-3.5" aria-hidden="true" />
                </button>
                <button
                  type="button"
                  onClick={() => onStartWorkspace('report')}
                  className="inline-flex items-center gap-1.5 text-xs font-semibold text-[#64748B] hover:text-[#0F172A] cursor-pointer"
                >
                  {c.card4CtaReport}
                </button>
              </div>
            </div>
          </div>
        </section>

        {/* Workflow & Interactive Assistant Entry */}
        <section id="workflow" className="mx-auto max-w-6xl px-6 py-10">
          <div className="rounded-xl border border-[#E2E8F0] bg-white p-8 shadow-[0_1px_3px_rgba(0,0,0,0.05)]">
            <div className="grid grid-cols-1 gap-8 lg:grid-cols-12 lg:items-center">
              <div className="lg:col-span-7">
                <h2 className="text-xl font-bold tracking-tight text-[#0F172A] sm:text-2xl">
                  {c.workflowTitle}
                </h2>
                <div className="mt-6 space-y-5">
                  <div className="border-b border-[#E2E8F0] pb-4">
                    <p className="font-mono text-[11px] font-semibold text-[#16A34A]">
                      {c.step1Label}
                    </p>
                    <p className="mt-1 text-sm font-semibold text-[#0F172A]">{c.step1Title}</p>
                    <p className="mt-1 text-xs text-[#64748B]">{c.step1Desc}</p>
                  </div>
                  <div className="border-b border-[#E2E8F0] pb-4">
                    <p className="font-mono text-[11px] font-semibold text-[#16A34A]">
                      {c.step2Label}
                    </p>
                    <p className="mt-1 text-sm font-semibold text-[#0F172A]">{c.step2Title}</p>
                    <p className="mt-1 text-xs text-[#64748B]">{c.step2Desc}</p>
                  </div>
                  <div>
                    <p className="font-mono text-[11px] font-semibold text-[#16A34A]">
                      {c.step3Label}
                    </p>
                    <p className="mt-1 text-sm font-semibold text-[#0F172A]">{c.step3Title}</p>
                    <p className="mt-1 text-xs text-[#64748B]">{c.step3Desc}</p>
                  </div>
                </div>
              </div>

              <div
                id="assistant"
                className="flex flex-col justify-between rounded-lg border border-[#E2E8F0] bg-[#F8FAFC] p-6 lg:col-span-5"
              >
                <div>
                  <h3 className="text-base font-bold text-[#0F172A]">{c.readyTitle}</h3>
                  <p className="mt-2 text-xs leading-relaxed text-[#64748B]">{c.readyDesc}</p>
                </div>

                <div className="mt-6 flex flex-col gap-3">
                  <button
                    type="button"
                    onClick={() => onStartWorkspace('all')}
                    className="inline-flex w-full items-center justify-center gap-2 rounded-md bg-[#16A34A] px-4 py-2.5 text-sm font-semibold text-white transition-colors hover:bg-[#15803D] cursor-pointer"
                  >
                    {c.enterWorkspaceBtn}
                    <ArrowRight className="h-4 w-4" aria-hidden="true" />
                  </button>
                  <button
                    type="button"
                    onClick={onOpenChatbot}
                    className="inline-flex w-full items-center justify-center gap-2 rounded-md border border-[#E2E8F0] bg-white px-4 py-2.5 text-xs font-semibold text-[#0F172A] transition-colors hover:border-[#94A3B8] hover:bg-slate-50 cursor-pointer"
                  >
                    <MessageSquareText className="h-4 w-4 text-[#16A34A]" aria-hidden="true" />
                    {c.chatAiBtn}
                  </button>
                </div>
              </div>
            </div>
          </div>
        </section>
      </main>

      {/* Quiet Footer */}
      <footer className="mt-12 border-t border-[#E2E8F0] bg-white">
        <div className="mx-auto flex max-w-6xl flex-col items-start justify-between gap-4 px-6 py-8 sm:flex-row sm:items-center">
          <div>
            <p className="text-sm font-bold text-[#0F172A]">Tokutei CRM</p>
            <p className="mt-1 text-xs text-[#64748B]">{c.footerDesc}</p>
          </div>
          <div className="flex flex-wrap items-center gap-5 text-xs font-medium text-[#64748B]">
            <button
              type="button"
              onClick={() => onStartWorkspace('all')}
              className="hover:text-[#0F172A] cursor-pointer"
            >
              {c.footerCand}
            </button>
            <button
              type="button"
              onClick={() => onStartWorkspace('kyujin')}
              className="hover:text-[#0F172A] cursor-pointer"
            >
              {c.footerJobs}
            </button>
            <button
              type="button"
              onClick={() => onStartWorkspace('report')}
              className="hover:text-[#0F172A] cursor-pointer"
            >
              {c.footerReport}
            </button>
            <button
              type="button"
              onClick={onOpenChatbot}
              className="hover:text-[#0F172A] cursor-pointer"
            >
              {c.footerAi}
            </button>
          </div>
        </div>
      </footer>
    </div>
  );
}
