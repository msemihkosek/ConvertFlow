import ffmpeg from 'fluent-ffmpeg';
import ffmpegPath from 'ffmpeg-static';
import ffprobeStatic from 'ffprobe-static';
import path from 'path';
import fs from 'fs';

// Binaries cannot be executed from inside app.asar, use the unpacked copy when packaged
const unpacked = (p) => (p ? p.replace('app.asar' + path.sep, 'app.asar.unpacked' + path.sep) : p);

// Configure binary paths
if (ffmpegPath) {
  ffmpeg.setFfmpegPath(unpacked(ffmpegPath));
}
if (ffprobeStatic && ffprobeStatic.path) {
  ffmpeg.setFfprobePath(unpacked(ffprobeStatic.path));
}

/**
 * Converts audio/video media files using the bundled high-performance FFmpeg binary
 * 
 * @param {string} inputPath - Absolute path to uploaded media
 * @param {string} outputPath - Target output path
 * @param {string} targetFormat - Target extension (e.g. 'mp3', 'mp4', 'wav', 'gif')
 * @param {Object} options - Quality, resolution, bitrate options
 * @param {Function} onProgress - Progress callback (percentage)
 */
export function convertMedia(inputPath, outputPath, targetFormat, options = {}, onProgress = () => {}) {
  return new Promise((resolve, reject) => {
    const target = targetFormat.toLowerCase();
    let command = ffmpeg(inputPath);

    // Audio format specific optimizations
    if (['mp3', 'wav', 'ogg', 'aac', 'flac', 'm4a', 'wma', 'opus'].includes(target)) {
      command.noVideo();

      if (target === 'mp3') {
        command.audioCodec('libmp3lame');
        command.audioBitrate(options.audioBitrate || '192k');
      } else if (target === 'wav') {
        command.audioCodec('pcm_s16le');
      } else if (target === 'aac' || target === 'm4a') {
        command.audioCodec('aac');
        command.audioBitrate(options.audioBitrate || '192k');
      } else if (target === 'ogg') {
        command.audioCodec('libvorbis');
      } else if (target === 'flac') {
        command.audioCodec('flac');
      }
    } 
    // Video to Animated GIF
    else if (target === 'gif') {
      command
        .fps(options.fps || 15)
        .size(options.resolution || '480x?')
        .outputOptions([
          '-vf', 'split[s0][s1];[s0]palettegen[p];[s1][p]paletteuse'
        ]);
    } 
    // Video to Video
    else if (['mp4', 'mkv', 'webm', 'avi', 'mov', 'flv'].includes(target)) {
      if (target === 'mp4') {
        command.videoCodec('libx264').audioCodec('aac');
        command.outputOptions(['-preset fast', '-crf 23', '-pix_fmt yuv420p', '-movflags +faststart']);
      } else if (target === 'webm') {
        command.videoCodec('libvpx-vp9').audioCodec('libopus');
        command.outputOptions(['-crf 30', '-b:v 0']);
      } else if (target === 'mkv') {
        command.videoCodec('libx264').audioCodec('aac');
      } else if (target === 'avi') {
        command.videoCodec('mpeg4').audioCodec('mp3');
      }

      if (options.resolution && options.resolution !== 'original') {
        command.size(options.resolution);
      }
    }

    command
      .on('progress', (progress) => {
        if (progress.percent) {
          onProgress(Math.min(99, Math.round(progress.percent)));
        }
      })
      .on('end', () => {
        onProgress(100);
        resolve(outputPath);
      })
      .on('error', (err) => {
        console.error('FFmpeg Conversion Error:', err);
        reject(err);
      })
      .save(outputPath);
  });
}
