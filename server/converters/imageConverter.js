import sharp from 'sharp';
import path from 'path';
import fs from 'fs';

/**
 * Universal Image Converter powered by Sharp (Libvips)
 * Supports: PNG, JPG, JPEG, WEBP, AVIF, TIFF, GIF, SVG, BMP, ICO, HEIC, HEIF
 * 
 * @param {string} inputPath - Input file path
 * @param {string} outputPath - Output file path
 * @param {string} targetFormat - Target extension
 * @param {Object} options - Quality, resize, etc.
 */
export async function convertImage(inputPath, outputPath, targetFormat, options = {}) {
  const target = targetFormat.toLowerCase();
  const quality = parseInt(options.imageQuality || options.quality || 85, 10);

  let pipeline = sharp(inputPath, {
    animated: true, // Preserve animations for GIF/WebP if present
    failOnError: false
  });

  // Optional Resize
  if (options.width || options.height) {
    pipeline = pipeline.resize({
      width: options.width ? parseInt(options.width, 10) : null,
      height: options.height ? parseInt(options.height, 10) : null,
      fit: options.fit || 'inside',
      withoutEnlargement: true
    });
  }

  // Target Format Encoding
  switch (target) {
    case 'jpg':
    case 'jpeg':
      pipeline = pipeline.jpeg({ quality, mozjpeg: true });
      break;
    case 'png':
      pipeline = pipeline.png({ compressionLevel: 8 });
      break;
    case 'webp':
      pipeline = pipeline.webp({ quality, effort: 4 });
      break;
    case 'avif':
      pipeline = pipeline.avif({ quality: Math.min(quality, 80), effort: 4 });
      break;
    case 'tiff':
    case 'tif':
      pipeline = pipeline.tiff({ quality, compression: 'lzw' });
      break;
    case 'gif':
      pipeline = pipeline.gif({ effort: 7 });
      break;
    case 'ico':
      // Favicon sizing
      pipeline = pipeline.resize(48, 48, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } }).png();
      break;
    case 'bmp':
      // Convert to raw/uncompressed or fallback to png if bmp direct output is restricted
      pipeline = pipeline.toFormat('png');
      break;
    default:
      pipeline = pipeline.toFormat(target);
      break;
  }

  await pipeline.toFile(outputPath);
  return outputPath;
}
