import * as pdfjsLib from 'pdfjs-dist';

const PDFJS_VERSION = pdfjsLib.version || '6.5.438';
const PRIMARY_CDN_BASE = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${PDFJS_VERSION}`;
const FALLBACK_CDN_BASE = `https://unpkg.com/pdfjs-dist@${PDFJS_VERSION}`;
const LOCAL_ASSETS_BASE = '/pdfjs-assets';

// Configure PDF.js worker using exact matching library version via stable CDN
pdfjsLib.GlobalWorkerOptions.workerSrc = `${PRIMARY_CDN_BASE}/build/pdf.worker.min.mjs`;

export interface RenderedPdfPageImage {
  pageNumber: number;
  mimeType: 'image/jpeg' | 'image/png';
  base64Data: string;
}

export interface PdfExtractionResult {
  mode: 'text' | 'ocr_images';
  text: string;
  meaningfulCharCount: number;
  numPages: number;
  renderedImages: RenderedPdfPageImage[];
}

interface TextRunItem {
  str: string;
  x: number;
  y: number;
  width: number;
  height: number;
}

const MAX_OCR_PAGES = 10;
const MIN_MEANINGFUL_TEXT_CHARS = 30;

const countMeaningfulPdfChars = (text: string): number => {
  // Count Japanese (Hiragana, Katakana, Kanji) + ASCII alphanumeric + ¥/円
  // Exclude replacement characters (\uFFFD), whitespace, punctuation-only noise
  const matched = String(text || '').match(
    /[A-Za-z0-9\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF\uFF10-\uFF19\uFF21-\uFF3A\uFF41-\uFF5A]/g
  );
  return matched ? matched.length : 0;
};

const reconstructPageLines = (items: unknown[]): string => {
  const runs: TextRunItem[] = [];
  for (const raw of items) {
    if (!raw || typeof raw !== 'object') continue;
    const candidate = raw as {
      str?: unknown;
      transform?: unknown;
      width?: unknown;
      height?: unknown;
    };
    if (typeof candidate.str !== 'string' || !candidate.str) continue;
    const tr = Array.isArray(candidate.transform) ? candidate.transform : [1, 0, 0, 1, 0, 0];
    const x = typeof tr[4] === 'number' ? tr[4] : 0;
    const y = typeof tr[5] === 'number' ? tr[5] : 0;
    const h =
      typeof candidate.height === 'number' && candidate.height > 0
        ? candidate.height
        : Math.abs(typeof tr[3] === 'number' ? tr[3] : 10) || 10;
    const w = typeof candidate.width === 'number' && candidate.width >= 0 ? candidate.width : 0;
    runs.push({ str: candidate.str, x, y, width: w, height: h });
  }

  if (runs.length === 0) return '';

  // Sort top-to-bottom (higher Y in PDF coords comes first), then left-to-right
  runs.sort((a, b) => {
    if (Math.abs(a.y - b.y) > 3.5) return b.y - a.y;
    return a.x - b.x;
  });

  const lineGroups: TextRunItem[][] = [];
  for (const run of runs) {
    const lastGroup = lineGroups[lineGroups.length - 1];
    if (!lastGroup) {
      lineGroups.push([run]);
      continue;
    }
    const refY = lastGroup[0].y;
    const tolerance = Math.max(3.5, Math.min(8, lastGroup[0].height * 0.45));
    if (Math.abs(run.y - refY) <= tolerance) {
      lastGroup.push(run);
    } else {
      lineGroups.push([run]);
    }
  }

  return lineGroups
    .map((group) => {
      group.sort((a, b) => a.x - b.x);
      let line = '';
      let prevEndX: number | null = null;
      let prevHeight = 10;
      for (const item of group) {
        const piece = item.str;
        if (!piece) continue;
        if (prevEndX !== null) {
          const gap = item.x - prevEndX;
          const unit = Math.max(6, Math.min(prevHeight, item.height || prevHeight));
          if (gap > unit * 1.8) {
            line += ' | ';
          } else if (gap > unit * 0.38) {
            line += ' ';
          }
        }
        line += piece;
        prevEndX = item.x + item.width;
        if (item.height > 0) prevHeight = item.height;
      }
      return line.replace(/\s+/g, ' ').trim();
    })
    .filter(Boolean)
    .join('\n');
};

const classifyPdfLoadError = (err: unknown, fileName: string): Error => {
  const errName =
    err && typeof err === 'object' && 'name' in err ? String((err as { name?: unknown }).name) : '';
  const errMsg = err instanceof Error ? err.message : String(err || '');

  if (errName === 'PasswordException' || /password|encrypted/i.test(errMsg)) {
    return new Error(
      `File PDF "${fileName}" đang được đặt mật khẩu bảo vệ. Vui lòng mở khóa mật khẩu trước khi tải lên.`
    );
  }
  if (
    errName === 'InvalidPDFException' ||
    errName === 'MissingPDFException' ||
    /invalid pdf|corrupt|startxref|stream must have data|unexpected end/i.test(errMsg)
  ) {
    return new Error(
      `File PDF "${fileName}" bị hỏng hoặc không đúng định dạng PDF chuẩn. Vui lòng kiểm tra hoặc xuất lại file PDF.`
    );
  }
  return new Error(
    `Không thể đọc file PDF "${fileName}": ${errMsg || 'Lỗi khởi tạo trình đọc PDF'}`
  );
};

const openPdfDocumentWithFallback = async (
  rawBytes: Uint8Array,
  fileName: string
) => {
  const bases = [PRIMARY_CDN_BASE, FALLBACK_CDN_BASE, LOCAL_ASSETS_BASE];
  let lastError: unknown = null;

  for (const base of bases) {
    try {
      pdfjsLib.GlobalWorkerOptions.workerSrc = `${base}/build/pdf.worker.min.mjs`;
      const loadingTask = pdfjsLib.getDocument({
        data: rawBytes.slice(),
        cMapUrl: `${base}/cmaps/`,
        cMapPacked: true,
        standardFontDataUrl: `${base}/standard_fonts/`,
        useSystemFonts: true,
      });
      return await loadingTask.promise;
    } catch (err) {
      lastError = err;
      const errName =
        err && typeof err === 'object' && 'name' in err
          ? String((err as { name?: unknown }).name)
          : '';
      const errMsg = err instanceof Error ? err.message : String(err || '');
      // Do not retry other CDNs for password or invalid PDF content errors
      if (
        errName === 'PasswordException' ||
        errName === 'InvalidPDFException' ||
        errName === 'MissingPDFException' ||
        /password|invalid pdf/i.test(errMsg)
      ) {
        throw classifyPdfLoadError(err, fileName);
      }
      // Otherwise (e.g. CDN/worker fetch failed), try next base
    }
  }

  throw classifyPdfLoadError(lastError, fileName);
};

const MAX_VISUAL_PAGES = 8;

const renderPageToBase64Image = async (
  pdfDoc: Awaited<ReturnType<typeof openPdfDocumentWithFallback>>,
  pageNumber: number,
  maxTargetDim = 1700,
  jpegQuality = 0.85
): Promise<RenderedPdfPageImage> => {
  const page = await pdfDoc.getPage(pageNumber);
  const baseViewport = page.getViewport({ scale: 1 });
  const maxDim = Math.max(baseViewport.width, baseViewport.height, 1);
  // Target crisp resolution (clamped between 1.35x and 2.2x) for Japanese text, checkboxes & radio circles
  const scale = Math.min(2.2, Math.max(1.35, maxTargetDim / maxDim));
  const viewport = page.getViewport({ scale });

  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, Math.floor(viewport.width));
  canvas.height = Math.max(1, Math.floor(viewport.height));
  const ctx = canvas.getContext('2d');
  if (!ctx) {
    throw new Error('Trình duyệt không hỗ trợ Canvas 2D để xử lý PDF ảnh scan.');
  }

  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  await page.render({
    canvas,
    canvasContext: ctx,
    viewport,
    background: '#ffffff',
  }).promise;

  const dataUrl = canvas.toDataURL('image/jpeg', jpegQuality);
  const commaIdx = dataUrl.indexOf(',');
  const base64Data = commaIdx >= 0 ? dataUrl.slice(commaIdx + 1) : dataUrl;

  // Free canvas memory immediately
  canvas.width = 0;
  canvas.height = 0;

  return {
    pageNumber,
    mimeType: 'image/jpeg',
    base64Data,
  };
};

export const extractPdfWithOcrFallback = async (
  file: File,
  onProgress?: (statusText: string) => void,
  options?: { includeVisualImages?: boolean }
): Promise<PdfExtractionResult> => {
  if (file.size === 0) {
    throw new Error(`File PDF "${file.name}" bị rỗng (0 bytes).`);
  }

  onProgress?.('Đang mở file PDF...');
  const arrayBuffer = await file.arrayBuffer();
  const rawBytes = new Uint8Array(arrayBuffer);
  const pdfDoc = await openPdfDocumentWithFallback(rawBytes, file.name);
  const numPages = pdfDoc.numPages || 0;

  if (numPages === 0) {
    throw new Error(`File PDF "${file.name}" không có trang nào.`);
  }

  const pageTexts: string[] = [];
  for (let p = 1; p <= numPages; p++) {
    onProgress?.(`Đang đọc trang ${p}/${numPages}...`);
    const page = await pdfDoc.getPage(p);
    const textContent = await page.getTextContent({ includeMarkedContent: false });
    const pageStr = reconstructPageLines(textContent.items || []);
    if (pageStr) {
      pageTexts.push(`--- Trang ${p}/${numPages} ---\n${pageStr}`);
    }
    await new Promise((r) => setTimeout(r, 0));
  }

  const combinedText = pageTexts.join('\n\n').trim();
  const meaningfulCharCount = countMeaningfulPdfChars(combinedText);

  if (meaningfulCharCount >= MIN_MEANINGFUL_TEXT_CHARS) {
    const renderedImages: RenderedPdfPageImage[] = [];
    if (options?.includeVisualImages) {
      const visualCount = Math.min(numPages, MAX_VISUAL_PAGES);
      for (let p = 1; p <= visualCount; p++) {
        onProgress?.(`Đang chụp bố cục bảng & checkbox trang ${p}/${visualCount}...`);
        try {
          const rendered = await renderPageToBase64Image(pdfDoc, p, 1450, 0.8);
          renderedImages.push(rendered);
        } catch {
          // Keep extracted text even if canvas render fails on a page
        }
        await new Promise((r) => setTimeout(r, 0));
      }
    }
    try {
      await pdfDoc.cleanup(false);
    } catch {
      // ignore cleanup error
    }
    return {
      mode: 'text',
      text: combinedText,
      meaningfulCharCount,
      numPages,
      renderedImages,
    };
  }

  // Fallback to Canvas -> Image OCR / Vision AI for scanned / outline-only PDFs
  const pagesToRender = Math.min(numPages, MAX_OCR_PAGES);
  onProgress?.(
    `Phát hiện PDF ảnh scan (${meaningfulCharCount} ký tự), đang chuyển sang nhận diện OCR...`
  );

  const renderedImages: RenderedPdfPageImage[] = [];
  for (let p = 1; p <= pagesToRender; p++) {
    onProgress?.(`Phát hiện PDF ảnh scan, đang render OCR trang ${p}/${pagesToRender}...`);
    const rendered = await renderPageToBase64Image(pdfDoc, p, 1700, 0.86);
    renderedImages.push(rendered);
    await new Promise((r) => setTimeout(r, 0));
  }

  try {
    await pdfDoc.cleanup(false);
  } catch {
    // ignore cleanup error
  }

  return {
    mode: 'ocr_images',
    text: combinedText,
    meaningfulCharCount,
    numPages,
    renderedImages,
  };
};
