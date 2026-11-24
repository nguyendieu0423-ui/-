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

interface HomePageProps {
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

export function HomePage({
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

  return (
    <div className="min-h-screen bg-[#0c0c0e] text-[#edeeef]">
      {/* Top Bar Contract: 3 exact zones (Wordmark — 4 Links — 1 Primary CTA) */}
      <header className="sticky top-0 z-30 border-b border-[#27272a] bg-[#141417]/95 backdrop-blur">
        <div className="mx-auto flex max-w-6xl items-center justify-between gap-8 px-6 py-4">
          <a
            href="#overview"
            className="text-base font-bold tracking-tight text-[#edeeef] whitespace-nowrap shrink-0"
          >
            Tokutei CRM
          </a>

          <nav
            aria-label="Điều hướng trang chủ"
            className="hidden md:flex items-center gap-6 text-sm font-medium text-[#94949e]"
          >
            <a
              href="#overview"
              className="hover:text-[#edeeef] transition-colors whitespace-nowrap shrink-0"
            >
              Giới thiệu
            </a>
            <a
              href="#capabilities"
              className="hover:text-[#edeeef] transition-colors whitespace-nowrap shrink-0"
            >
              Tính năng
            </a>
            <a
              href="#workflow"
              className="hover:text-[#edeeef] transition-colors whitespace-nowrap shrink-0"
            >
              Quy trình
            </a>
            <a
              href="#assistant"
              className="hover:text-[#edeeef] transition-colors whitespace-nowrap shrink-0"
            >
              Trợ lý AI
            </a>
          </nav>

          <div className="shrink-0">
            <button
              type="button"
              onClick={() => onStartWorkspace('all')}
              className="inline-flex items-center gap-2 rounded-md bg-[#22c55e] px-4 py-2 text-xs font-semibold text-[#0c0c0e] transition-opacity hover:opacity-90 whitespace-nowrap shrink-0 cursor-pointer"
            >
              Bắt đầu
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
              <div className="mb-4 flex flex-wrap items-center gap-2 font-mono text-xs text-[#94949e]">
                <span className="text-[#22c55e]">Tokutei Ginou Workstation</span>
                <span aria-hidden="true">·</span>
                <span>Nhật Bản</span>
              </div>

              <h1
                className="font-bold tracking-tight text-[#edeeef] text-balance"
                style={{ fontSize: 'clamp(2rem, 3.6vw, 3.1rem)', lineHeight: 1.15 }}
              >
                Quản lý ứng viên Kỹ năng Đặc định từ tin nhắn đầu tiên đến khi nhận việc.
              </h1>

              <p className="mt-5 max-w-xl text-base leading-relaxed text-[#94949e]">
                Tổng hợp ứng viên từ Facebook cá nhân, Fanpage công ty và trường Nhật ngữ trong một không gian điều phối mật độ cao. Trích xuất phiếu tuyển dụng bằng AI, tạo Rirekisho chuẩn Nhật và đồng bộ bảo mật trên Cloud SQL.
              </p>

              <div className="mt-8 flex flex-wrap items-center gap-4">
                <button
                  type="button"
                  onClick={() => onStartWorkspace('all')}
                  className="inline-flex items-center gap-2 rounded-md bg-[#22c55e] px-6 py-3 text-sm font-semibold text-[#0c0c0e] transition-opacity hover:opacity-90 whitespace-nowrap shrink-0 cursor-pointer"
                >
                  Bắt đầu ngay
                  <ArrowRight className="h-4 w-4" aria-hidden="true" />
                </button>

                {userEmail ? (
                  <div className="flex items-center gap-3 text-xs text-[#94949e]">
                    <span className="truncate max-w-[220px]" title={userEmail}>
                      Đã đăng nhập: <strong className="font-semibold text-[#edeeef]">{userEmail}</strong>
                    </span>
                    <button
                      type="button"
                      onClick={onSignOut}
                      className="font-medium text-[#94949e] underline underline-offset-4 hover:text-[#edeeef] whitespace-nowrap shrink-0 cursor-pointer"
                    >
                      Đăng xuất
                    </button>
                  </div>
                ) : (
                  <button
                    type="button"
                    disabled={authLoading}
                    onClick={onSignInGoogle}
                    className="inline-flex items-center gap-2.5 rounded-md border border-[#dadce0] bg-white px-5 py-3 text-sm font-medium text-[#3c4043] shadow-sm transition-colors hover:bg-[#f8f9fa] disabled:opacity-60 whitespace-nowrap shrink-0 cursor-pointer"
                  >
                    <svg className="h-4 w-4 shrink-0" viewBox="0 0 48 48" aria-hidden="true">
                      <path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z" />
                      <path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z" />
                      <path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z" />
                      <path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z" />
                    </svg>
                    {authLoading ? 'Đang kết nối...' : 'Sign in with Google'}
                  </button>
                )}
              </div>

              <div className="mt-8 flex flex-wrap items-center gap-x-4 gap-y-2 border-t border-[#27272a] pt-5 font-mono text-[11px] text-[#94949e]">
                <span>Tiếp nhận đa kênh FB</span>
                <span aria-hidden="true">·</span>
                <span>Rirekisho 2 cột</span>
                <span aria-hidden="true">·</span>
                <span>Trích xuất求人 PDF &amp; Excel</span>
                <span aria-hidden="true">·</span>
                <span>Đồng bộ Cloud SQL</span>
              </div>
            </div>

            <div className="lg:col-span-6">
              <div className="h-full min-h-[280px] overflow-hidden rounded-xl border border-[#27272a] bg-[#141417]">
                {heroImgFailed ? (
                  <div className="flex h-full min-h-[320px] flex-col items-center justify-center bg-[#141417] p-8 text-center">
                    <Building2 className="mb-3 h-10 w-10 text-[#94949e]" aria-hidden="true" />
                    <p className="text-sm font-semibold text-[#edeeef]">
                      Không gian điều phối tuyển dụng Tokutei Ginou
                    </p>
                    <p className="mt-1 max-w-xs text-xs text-[#94949e]">
                      Quản lý hồ sơ ứng viên, lịch Mendan và phiếu tuyển dụng tập trung.
                    </p>
                  </div>
                ) : (
                  <img
                    src={HERO_IMAGE_PATH}
                    alt="Văn phòng tư vấn tuyển dụng Tokutei Ginou tại Nhật Bản"
                    referrerPolicy="no-referrer"
                    onError={() => setHeroImgFailed(true)}
                    className="h-full w-full object-cover lg:min-h-[420px]"
                  />
                )}
              </div>
            </div>
          </div>

          {/* Full-width Operational Metrics Strip below Hero Split */}
          <div className="mt-10 grid grid-cols-2 gap-6 rounded-lg border border-[#27272a] bg-[#141417] px-6 py-5 sm:grid-cols-4">
            <div>
              <p className="font-mono text-[10px] tracking-wider text-[#94949e]">
                Ứng viên &amp; Tư vấn FB
              </p>
              <p className="mt-1 font-mono text-2xl font-bold tracking-tight text-[#edeeef] tabular-nums">
                {candidateCount}
              </p>
            </div>
            <div>
              <p className="font-mono text-[10px] tracking-wider text-[#94949e]">
                Du học sinh trường Nhật
              </p>
              <p className="mt-1 font-mono text-2xl font-bold tracking-tight text-[#edeeef] tabular-nums">
                {schoolCount}
              </p>
            </div>
            <div>
              <p className="font-mono text-[10px] tracking-wider text-[#94949e]">
                Đơn hàng đang mở
              </p>
              <p className="mt-1 font-mono text-2xl font-bold tracking-tight text-[#22c55e] tabular-nums">
                {companyCount}{' '}
                <span className="text-xs font-medium text-[#94949e]">({slotCount} suất)</span>
              </p>
            </div>
            <div>
              <p className="font-mono text-[10px] tracking-wider text-[#94949e]">
                Lịch Mendan &amp; Đối tác PT
              </p>
              <p className="mt-1 font-mono text-2xl font-bold tracking-tight text-[#edeeef] tabular-nums">
                {upcomingMendanCount}{' '}
                <span className="text-xs font-medium text-[#94949e]">/ {ptCount} PT</span>
              </p>
            </div>
          </div>
        </section>

        {/* Capabilities Bento Grid */}
        <section id="capabilities" className="mx-auto max-w-6xl px-6 py-12">
          <div className="mb-8 max-w-2xl">
            <h2 className="text-2xl font-bold tracking-tight text-[#edeeef] sm:text-3xl text-balance">
              Mọi công cụ điều phối trong một bàn làm việc thống nhất
            </h2>
            <p className="mt-2 text-sm leading-relaxed text-[#94949e]">
              Thiết kế 3 cột mật độ cao giúp xử lý danh sách ứng viên, ghép đơn hàng và theo dõi cảnh báo khẩn trong cùng một màn hình.
            </p>
          </div>

          <div className="grid grid-cols-1 gap-4 lg:grid-cols-3">
            {/* Card 01 */}
            <div className="flex flex-col justify-between rounded-lg border border-[#27272a] bg-[#141417] p-6 lg:col-span-2">
              <div>
                <div className="flex items-center justify-between font-mono text-xs text-[#94949e]">
                  <span>01. Quản lý ứng viên &amp; Cảnh báo Mendan</span>
                  <Users className="h-4 w-4 text-[#22c55e]" aria-hidden="true" />
                </div>
                <h3 className="mt-3 text-lg font-bold text-[#edeeef]">
                  Theo dõi ứng viên FB cá nhân, Fanpage và Trường Nhật ngữ
                </h3>
                <p className="mt-2 text-sm leading-relaxed text-[#94949e]">
                  Tự động bóc tách đoạn chat Facebook thành JLPT, SĐT, hạn thẻ ngoại kiều, ngành nghề và khu vực mong muốn. Bảng cảnh báo bên phải luôn hiển thị lịch Mendan sắp tới và hồ sơ thiếu thông tin.
                </p>
              </div>
              <div className="mt-6">
                <button
                  type="button"
                  onClick={() => onStartWorkspace('all')}
                  className="inline-flex items-center gap-1.5 text-xs font-semibold text-[#22c55e] hover:underline cursor-pointer"
                >
                  Mở danh sách ứng viên
                  <ArrowRight className="h-3.5 w-3.5" aria-hidden="true" />
                </button>
              </div>
            </div>

            {/* Card 02 */}
            <div className="flex flex-col justify-between rounded-lg border border-[#27272a] bg-[#141417] p-6">
              <div>
                <div className="flex items-center justify-between font-mono text-xs text-[#94949e]">
                  <span>02. Phiếu tuyển dụng</span>
                  <FileSpreadsheet className="h-4 w-4 text-[#22c55e]" aria-hidden="true" />
                </div>
                <h3 className="mt-3 text-lg font-bold text-[#edeeef]">
                  Đọc Excel, PDF và khớp đơn hàng tự động
                </h3>
                <p className="mt-2 text-sm leading-relaxed text-[#94949e]">
                  Tải lên nhiều phiếu tuyển dụng cùng lúc để trích xuất mức lương, thưởng, ký túc xá và điều kiện JLPT. Xuất bản tóm tắt ẩn danh tiếng Việt hoặc tiếng Nhật.
                </p>
              </div>
              <div className="mt-6">
                <button
                  type="button"
                  onClick={() => onStartWorkspace('kyujin')}
                  className="inline-flex items-center gap-1.5 text-xs font-semibold text-[#22c55e] hover:underline cursor-pointer"
                >
                  Quản lý đơn hàng
                  <ArrowRight className="h-3.5 w-3.5" aria-hidden="true" />
                </button>
              </div>
            </div>

            {/* Card 03 */}
            <div className="flex flex-col justify-between rounded-lg border border-[#27272a] bg-[#141417] p-6">
              <div>
                <div className="flex items-center justify-between font-mono text-xs text-[#94949e]">
                  <span>03. Rirekisho &amp; Soạn tin</span>
                  <ShieldCheck className="h-4 w-4 text-[#22c55e]" aria-hidden="true" />
                </div>
                <h3 className="mt-3 text-lg font-bold text-[#edeeef]">
                  Sơ yếu lý lịch 2 cột &amp; Tin nhắn đa ngôn ngữ
                </h3>
                <p className="mt-2 text-sm leading-relaxed text-[#94949e]">
                  Điền tự động Rirekisho từ file hoặc đoạn hội thoại, in hoặc xuất HTML độc lập. Tạo nhanh mẫu nhắc lịch bằng tiếng Việt, Nhật, Anh.
                </p>
              </div>
              <div className="mt-6">
                <button
                  type="button"
                  onClick={() => onStartWorkspace('proposed')}
                  className="inline-flex items-center gap-1.5 text-xs font-semibold text-[#22c55e] hover:underline cursor-pointer"
                >
                  Xem tiến độ tiến cử
                  <ArrowRight className="h-3.5 w-3.5" aria-hidden="true" />
                </button>
              </div>
            </div>

            {/* Card 04 */}
            <div className="flex flex-col justify-between rounded-lg border border-[#27272a] bg-[#141417] p-6 lg:col-span-2">
              <div>
                <div className="flex items-center justify-between font-mono text-xs text-[#94949e]">
                  <span>04. Trợ lý AI Gemini &amp; Cloud SQL PostgreSQL</span>
                  <MessageSquareText className="h-4 w-4 text-[#22c55e]" aria-hidden="true" />
                </div>
                <h3 className="mt-3 text-lg font-bold text-[#edeeef]">
                  Hội thoại AI nhiều lượt theo vai trò &amp; Lưu trữ bền vững
                </h3>
                <p className="mt-2 text-sm leading-relaxed text-[#94949e]">
                  Chuyển đổi linh hoạt giữa Tư vấn chiến lược CRM, Soạn phản hồi nhanh (Việt/Nhật/Anh) và Cố vấn sàng lọc Mendan. Khi đăng nhập tài khoản Google, toàn bộ hồ sơ ứng viên, đơn hàng và lịch sử trò chuyện được đồng bộ lên Cloud SQL.
                </p>
              </div>
              <div className="mt-6 flex flex-wrap items-center gap-4">
                <button
                  type="button"
                  onClick={onOpenChatbot}
                  className="inline-flex items-center gap-1.5 text-xs font-semibold text-[#22c55e] hover:underline cursor-pointer"
                >
                  Mở Trợ lý AI Gemini
                  <ArrowRight className="h-3.5 w-3.5" aria-hidden="true" />
                </button>
                <button
                  type="button"
                  onClick={() => onStartWorkspace('report')}
                  className="inline-flex items-center gap-1.5 text-xs font-semibold text-[#94949e] hover:text-[#edeeef] cursor-pointer"
                >
                  Xem báo cáo phân tích
                </button>
              </div>
            </div>
          </div>
        </section>

        {/* Workflow & Interactive Assistant Entry */}
        <section id="workflow" className="mx-auto max-w-6xl px-6 py-10">
          <div className="rounded-xl border border-[#27272a] bg-[#141417] p-8">
            <div className="grid grid-cols-1 gap-8 lg:grid-cols-12 lg:items-center">
              <div className="lg:col-span-7">
                <h2 className="text-xl font-bold tracking-tight text-[#edeeef] sm:text-2xl">
                  Quy trình làm việc tinh gọn cho điều phối viên
                </h2>
                <div className="mt-6 space-y-5">
                  <div className="border-b border-[#27272a] pb-4">
                    <p className="font-mono text-[11px] font-semibold text-[#22c55e]">Bước 1</p>
                    <p className="mt-1 text-sm font-semibold text-[#edeeef]">
                      Nhập đoạn chat Messenger hoặc danh sách du học sinh
                    </p>
                    <p className="mt-1 text-xs text-[#94949e]">
                      Trích xuất tự động thông tin cá nhân, trình độ tiếng Nhật và khu vực làm việc mong muốn.
                    </p>
                  </div>
                  <div className="border-b border-[#27272a] pb-4">
                    <p className="font-mono text-[11px] font-semibold text-[#22c55e]">Bước 2</p>
                    <p className="mt-1 text-sm font-semibold text-[#edeeef]">
                      Đối chiếu phiếu tuyển dụng và hoàn thiện Rirekisho 2 cột
                    </p>
                    <p className="mt-1 text-xs text-[#94949e]">
                      Xem bảng gợi ý job ngay trong thẻ ứng viên và tạo hồ sơ chuẩn Nhật chỉ trong vài phút.
                    </p>
                  </div>
                  <div>
                    <p className="font-mono text-[11px] font-semibold text-[#22c55e]">Bước 3</p>
                    <p className="mt-1 text-sm font-semibold text-[#edeeef]">
                      Xử lý cảnh báo khẩn và đồng bộ dữ liệu Cloud SQL
                    </p>
                    <p className="mt-1 text-xs text-[#94949e]">
                      Đăng nhập Google để lưu trữ lâu dài toàn bộ ứng viên, đơn hàng và lịch sử trò chuyện AI giữa các phiên.
                    </p>
                  </div>
                </div>
              </div>

              <div
                id="assistant"
                className="flex flex-col justify-between rounded-lg border border-[#27272a] bg-[#0c0c0e] p-6 lg:col-span-5"
              >
                <div>
                  <h3 className="text-base font-bold text-[#edeeef]">
                    Sẵn sàng bắt đầu phiên điều phối?
                  </h3>
                  <p className="mt-2 text-xs leading-relaxed text-[#94949e]">
                    Truy cập trực tiếp giao diện 3 cột mật độ cao với dữ liệu mẫu hoặc đăng nhập Google để đồng bộ cơ sở dữ liệu PostgreSQL cá nhân.
                  </p>
                </div>

                <div className="mt-6 flex flex-col gap-3">
                  <button
                    type="button"
                    onClick={() => onStartWorkspace('all')}
                    className="inline-flex w-full items-center justify-center gap-2 rounded-md bg-[#22c55e] px-4 py-2.5 text-sm font-semibold text-[#0c0c0e] transition-opacity hover:opacity-90 cursor-pointer"
                  >
                    Vào không gian làm việc CRM
                    <ArrowRight className="h-4 w-4" aria-hidden="true" />
                  </button>
                  <button
                    type="button"
                    onClick={onOpenChatbot}
                    className="inline-flex w-full items-center justify-center gap-2 rounded-md border border-[#27272a] bg-[#141417] px-4 py-2.5 text-xs font-semibold text-[#edeeef] transition-colors hover:border-[#94949e] cursor-pointer"
                  >
                    <MessageSquareText className="h-4 w-4 text-[#22c55e]" aria-hidden="true" />
                    Trò chuyện với Trợ lý AI Gemini
                  </button>
                </div>
              </div>
            </div>
          </div>
        </section>
      </main>

      {/* Quiet Footer */}
      <footer className="mt-12 border-t border-[#27272a] bg-[#141417]">
        <div className="mx-auto flex max-w-6xl flex-col items-start justify-between gap-4 px-6 py-8 sm:flex-row sm:items-center">
          <div>
            <p className="text-sm font-bold text-[#edeeef]">Tokutei CRM</p>
            <p className="mt-1 text-xs text-[#94949e]">
              Hệ thống quản lý ứng viên, phiếu tuyển dụng và đối tác Kỹ năng Đặc định Nhật Bản.
            </p>
          </div>
          <div className="flex flex-wrap items-center gap-5 text-xs font-medium text-[#94949e]">
            <button
              type="button"
              onClick={() => onStartWorkspace('all')}
              className="hover:text-[#edeeef] cursor-pointer"
            >
              Ứng viên
            </button>
            <button
              type="button"
              onClick={() => onStartWorkspace('kyujin')}
              className="hover:text-[#edeeef] cursor-pointer"
            >
              Đơn hàng
            </button>
            <button
              type="button"
              onClick={() => onStartWorkspace('report')}
              className="hover:text-[#edeeef] cursor-pointer"
            >
              Báo cáo
            </button>
            <button
              type="button"
              onClick={onOpenChatbot}
              className="hover:text-[#edeeef] cursor-pointer"
            >
              Trợ lý AI
            </button>
          </div>
        </div>
      </footer>
    </div>
  );
}
