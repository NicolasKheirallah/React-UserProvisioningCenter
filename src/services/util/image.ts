/** Data URL length cap matching IPersonalInfo.photoDataUrl's ~50KB budget (base64 is ~1.37x binary size). */
export const MAX_PHOTO_DATA_URL_LENGTH: number = 70_000;

/**
 * Downscales an image file to a square-ish thumbnail and re-encodes it as a
 * JPEG data URL, so a multi-megabyte camera photo doesn't blow up the job's
 * PayloadJson (a SharePoint Note field). Throws if the result is still over
 * MAX_PHOTO_DATA_URL_LENGTH after compression (e.g. a very busy/noisy image).
 */
export async function resizeImageToDataUrl(
  file: File,
  maxDimension: number = 200,
  quality: number = 0.75
): Promise<string> {
  const bitmap: ImageBitmap = await createImageBitmap(file);
  try {
    const scale: number = Math.min(1, maxDimension / Math.max(bitmap.width, bitmap.height));
    const width: number = Math.max(1, Math.round(bitmap.width * scale));
    const height: number = Math.max(1, Math.round(bitmap.height * scale));

    const canvas: HTMLCanvasElement = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    const ctx: CanvasRenderingContext2D | null = canvas.getContext('2d');
    if (!ctx) {
      throw new Error('Canvas 2D context unavailable');
    }
    ctx.drawImage(bitmap, 0, 0, width, height);

    const dataUrl: string = canvas.toDataURL('image/jpeg', quality);
    if (dataUrl.length > MAX_PHOTO_DATA_URL_LENGTH) {
      throw new Error('photoTooLarge');
    }
    return dataUrl;
  } finally {
    bitmap.close();
  }
}
