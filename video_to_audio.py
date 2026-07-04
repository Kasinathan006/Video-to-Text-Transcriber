#!/usr/bin/env python3
"""
video_to_audio.py - Memory-Efficient Batch Video to Audio Converter

Designed to process massive video files (e.g., 50+ GB recordings) without loading
them into RAM. Streams audio directly via FFmpeg subprocess.

Modes:
  wav16k : 16kHz mono 16-bit PCM WAV (Ideal & optimal input for OpenAI Whisper)
  copy   : Fast lossless stream copy (extracts audio stream as-is)
  mp3    : High-quality VBR MP3 (libmp3lame -q:a 2)
  ogg    : Ogg Opus voice note format (libopus -b:a 64k)

Usage:
  python video_to_audio.py --input-dir videos/ --output-dir audio_files/ --mode wav16k
"""

import os
import sys
import time
import shutil
import argparse
import subprocess
from pathlib import Path

VIDEO_EXTENSIONS = {
    ".mp4", ".mkv", ".avi", ".mov", ".flv", ".wmv", 
    ".webm", ".m4v", ".ts", ".mts", ".m2ts"
}

MODE_CONFIGS = {
    "wav16k": {
        "ext": ".wav",
        "desc": "16kHz Mono 16-bit PCM WAV (Optimized for Whisper AI)",
        "args": ["-vn", "-acodec", "pcm_s16le", "-ar", "16000", "-ac", "1"]
    },
    "copy": {
        "ext": ".m4a",  # Default fallback container for stream copy
        "desc": "Lossless stream copy (fastest, keeps original encoding)",
        "args": ["-vn", "-acodec", "copy"]
    },
    "mp3": {
        "ext": ".mp3",
        "desc": "High-quality VBR MP3 (~190 kbps)",
        "args": ["-vn", "-acodec", "libmp3lame", "-q:a", "2"]
    },
    "ogg": {
        "ext": ".ogg",
        "desc": "Compressed Ogg Opus voice format (64 kbps)",
        "args": ["-vn", "-acodec", "libopus", "-b:a", "64k"]
    }
}


def check_ffmpeg():
    """Verify ffmpeg is installed and accessible on PATH."""
    if not shutil.which("ffmpeg"):
        print("[ERROR] FFmpeg is not installed or not found in system PATH.")
        print("Please install FFmpeg:")
        print("  Windows : https://ffmpeg.org/download.html (add bin/ to PATH)")
        print("  Mac     : brew install ffmpeg")
        print("  Linux   : sudo apt install ffmpeg")
        sys.exit(1)


def get_video_duration(filepath):
    """Attempt to get video duration in seconds using ffprobe."""
    ffprobe = shutil.which("ffprobe")
    if not ffprobe:
        return None
    try:
        cmd = [
            ffprobe, "-v", "error", "-show_entries", "format=duration",
            "-of", "default=noprint_wrappers=1:nokey=1", str(filepath)
        ]
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=True)
        return float(result.stdout.strip())
    except Exception:
        return None


def format_duration(seconds):
    """Format seconds into HH:MM:SS."""
    if seconds is None:
        return "Unknown"
    h = int(seconds // 3600)
    m = int((seconds % 3600) // 60)
    s = int(seconds % 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def convert_file(input_path, output_path, mode, overwrite=False):
    """Convert a single video file to audio streaming via FFmpeg."""
    if output_path.exists() and not overwrite:
        print(f"  [SKIP] Output already exists: {output_path.name}")
        return True

    cfg = MODE_CONFIGS[mode]
    cmd = [
        "ffmpeg", "-y" if overwrite else "-n",
        "-i", str(input_path),
        *cfg["args"],
        str(output_path)
    ]

    duration = get_video_duration(input_path)
    file_size_mb = input_path.stat().st_size / (1024 * 1024)
    file_size_str = f"{file_size_mb:.2f} MB" if file_size_mb < 1024 else f"{file_size_mb/1024:.2f} GB"

    print(f"  [CONVERTING] {input_path.name} ({file_size_str}, {format_duration(duration)}) -> {output_path.name}")

    start_time = time.time()
    try:
        # Run ffmpeg redirecting output so it doesn't clutter terminal unless error occurs
        proc = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        elapsed = time.time() - start_time

        if proc.returncode != 0:
            print(f"  [FAILED] FFmpeg error on {input_path.name}:")
            # Print last few lines of stderr
            err_lines = proc.stderr.strip().split("\n")[-5:]
            for line in err_lines:
                print(f"    | {line}")
            if output_path.exists():
                output_path.unlink()  # Remove partial file
            return False

        out_size_mb = output_path.stat().st_size / (1024 * 1024)
        print(f"  [SUCCESS] Created {output_path.name} ({out_size_mb:.2f} MB) in {elapsed:.1f}s")
        return True

    except Exception as e:
        print(f"  [ERROR] Exception during conversion: {e}")
        if output_path.exists():
            output_path.unlink()
        return False


def main():
    parser = argparse.ArgumentParser(
        description="Memory-Efficient Batch Video to Audio Converter (Powered by FFmpeg)"
    )
    parser.add_argument(
        "-i", "--input-dir",
        default=".",
        help="Directory containing source video files (default: current directory)"
    )
    parser.add_argument(
        "-o", "--output-dir",
        default="audio_files",
        help="Destination directory for extracted audio files (default: audio_files/)"
    )
    parser.add_argument(
        "-m", "--mode",
        choices=list(MODE_CONFIGS.keys()),
        default="wav16k",
        help="Conversion mode (default: wav16k - best for Whisper AI)"
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Scan subdirectories inside input-dir recursively"
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite existing output files"
    )

    args = parser.parse_args()

    check_ffmpeg()

    in_dir = Path(args.input_dir).resolve()
    out_dir = Path(args.output_dir).resolve()

    if not in_dir.exists():
        print(f"[ERROR] Input directory not found: {in_dir}")
        sys.exit(1)

    out_dir.mkdir(parents=True, exist_ok=True)

    print("=" * 60)
    print(" VIDEO TO AUDIO BATCH CONVERTER")
    print("=" * 60)
    print(f" Input Directory  : {in_dir}")
    print(f" Output Directory : {out_dir}")
    print(f" Mode             : {args.mode} ({MODE_CONFIGS[args.mode]['desc']})")
    print("=" * 60)

    # Gather video files
    if args.recursive:
        files = [p for p in in_dir.rglob("*") if p.suffix.lower() in VIDEO_EXTENSIONS and p.is_file()]
    else:
        files = [p for p in in_dir.iterdir() if p.suffix.lower() in VIDEO_EXTENSIONS and p.is_file()]

    if not files:
        print(f"\nNo video files found in {in_dir}.")
        print(f"Supported extensions: {', '.join(sorted(VIDEO_EXTENSIONS))}")
        sys.exit(0)

    print(f"\nFound {len(files)} video file(s). Starting conversion...\n")

    success_count = 0
    fail_count = 0

    total_start = time.time()
    for idx, fpath in enumerate(sorted(files), 1):
        print(f"[{idx}/{len(files)}]")
        ext = MODE_CONFIGS[args.mode]["ext"]
        out_path = out_dir / (fpath.stem + ext)
        
        ok = convert_file(fpath, out_path, mode=args.mode, overwrite=args.overwrite)
        if ok:
            success_count += 1
        else:
            fail_count += 1
        print()

    total_elapsed = time.time() - total_start
    print("=" * 60)
    print(f"BATCH COMPLETE in {total_elapsed:.1f}s")
    print(f"  Successful : {success_count}")
    print(f"  Failed     : {fail_count}")
    print(f"  Output Dir : {out_dir}")
    print("=" * 60)
    
    if success_count > 0:
        print(f"\n[TIP] Ready for transcription! You can now run:")
        print(f"      python transcribe_to_docx.py")


if __name__ == "__main__":
    main()
