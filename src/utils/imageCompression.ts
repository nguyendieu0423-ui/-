export interface CompressedImagePayload {
  mimeType: string;
  base64Data: string;
  dataUrl: string;
}

export const fileToBase64Raw = (file: Blob): Promise<string> =>
  new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const res = String(reader.result || '');
      const idx = res.indexOf(',');
      resolve(idx >= 0 ? res.slice(idx + 1) : res);
    };
    reader.onerror = () => reject(new Error('Không thể đọc file từ trình duyệt.'));
    reader.readAsDataURL(file);
  });

export const compressImageFileToBase64 = async (
  file: Blob,
  maxDimension = 1600,
  quality = 0.78
): Promise<CompressedImagePayload> => {
  const rawMime = (file.type || 'image/png').toLowerCase();
  if (!rawMime.startsWith('image/')) {
    const base64Data = await fileToBase64Raw(file);
    return {
      mimeType: rawMime || 'application/octet-stream',
      base64Data,
      dataUrl: `data:${rawMime};base64,${base64Data}`,
    };
  }

  const objectUrl = URL.createObjectURL(file);
  try {
    const img = await new Promise<HTMLImageElement>((resolve, reject) => {
      const el = new Image();
      el.onload = () => resolve(el);
      el.onerror = () => reject(new Error('Không thể giải mã ảnh để nén.'));
      el.src = objectUrl;
    });

    const width = img.naturalWidth || img.width || 1000;
    const height = img.naturalHeight || img.height || 1000;
    const scale = Math.min(1, maxDimension / Math.max(width, height));
    const targetW = Math.max(1, Math.round(width * scale));
    const targetH = Math.max(1, Math.round(height * scale));

    const canvas = document.createElement('canvas');
    canvas.width = targetW;
    canvas.height = targetH;
    const ctx = canvas.getContext('2d');
    if (!ctx) {
      const base64Data = await fileToBase64Raw(file);
      return {
        mimeType: rawMime,
        base64Data,
        dataUrl: `data:${rawMime};base64,${base64Data}`,
      };
    }

    // Fill white background for transparent PNG screenshots so OCR remains crisp in JPEG
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, targetW, targetH);
    ctx.drawImage(img, 0, 0, targetW, targetH);

    const dataUrl = canvas.toDataURL('image/jpeg', quality);
    const commaIdx = dataUrl.indexOf(',');
    const base64Data = commaIdx >= 0 ? dataUrl.slice(commaIdx + 1) : dataUrl;
    return {
      mimeType: 'image/jpeg',
      base64Data,
      dataUrl,
    };
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
};

export const extractFirstImageFromClipboard = (
  clipboardData: DataTransfer | null | undefined
): File | null => {
  if (!clipboardData) return null;
  const items = Array.from(clipboardData.items || []);
  for (const item of items) {
    if (item.type && item.type.startsWith('image/')) {
      const f = item.getAsFile();
      if (f) return f;
    }
  }
  const files = Array.from(clipboardData.files || []);
  for (const f of files) {
    if (f.type && f.type.startsWith('image/')) {
      return f;
    }
  }
  return null;
};
