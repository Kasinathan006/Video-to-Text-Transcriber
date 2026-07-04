#!/usr/bin/env python3
"""
run_sarvam_pipeline.py - 1-Click Master Automation Pipeline

This script automates the entire end-to-end workflow:
  1. Detects any video files (.mp4, .mkv, .mov, etc.) in the current directory or specified input folder.
  2. Extracts audio from videos into 'audio_files/' using video_to_audio.py (fast & memory-efficient).
  3. Uses Sarvam AI ('saaras:v3') to transcribe all extracted/available audio files with automatic chunking.
  4. Generates the formatted Microsoft Word document ('sarvam_transcript.docx').

Usage:
  python run_sarvam_pipeline.py
"""

import os
import sys
import glob
import argparse
import subprocess

VIDEO_EXTENSIONS = ('.mp4', '.mkv', '.avi', '.mov', '.flv', '.wmv', '.webm', '.m4v', '.ts', '.mts')


def main():
    parser = argparse.ArgumentParser(description="1-Click Sarvam AI Video-to-Word Pipeline")
    parser.add_argument("--input-dir", default=".", help="Directory containing video or audio files (default: current dir)")
    parser.add_argument("--audio-dir", default="audio_files", help="Directory to store intermediate audio files")
    parser.add_argument("--output-docx", default="sarvam_transcript.docx", help="Output Word document filename")
    args = parser.parse_args()

    input_dir = args.input_dir
    audio_dir = args.audio_dir

    # Step 1: Check for video files
    video_files = []
    for ext in VIDEO_EXTENSIONS:
        video_files.extend(glob.glob(os.path.join(input_dir, f"*{ext}")))
        video_files.extend(glob.glob(os.path.join(input_dir, f"*{ext.upper()}")))
    video_files = sorted(set(video_files))

    if video_files:
        print(f"============================================================")
        print(f"STEP 1: Found {len(video_files)} video file(s). Extracting audio...")
        print(f"============================================================")
        cmd = [
            sys.executable, "video_to_audio.py",
            "--input-dir", input_dir,
            "--output-dir", audio_dir,
            "--mode", "wav16k"
        ]
        res = subprocess.run(cmd)
        if res.returncode != 0:
            print("[ERROR] Video-to-Audio extraction failed.")
            sys.exit(res.returncode)
        target_audio_dir = audio_dir
    else:
        print("No video files found in input directory. Checking for audio files directly...")
        if os.path.exists(audio_dir) and os.listdir(audio_dir):
            target_audio_dir = audio_dir
        else:
            target_audio_dir = input_dir

    # Step 2: Check for audio files
    audio_files = []
    for ext in ("*.wav", "*.ogg", "*.mp3", "*.m4a", "*.flac", "*.aac"):
        audio_files.extend(glob.glob(os.path.join(target_audio_dir, ext)))
    audio_files = sorted(set(audio_files))

    if not audio_files:
        print(f"[ERROR] No video or audio files found to process in '{target_audio_dir}'.")
        sys.exit(1)

    print(f"\n============================================================")
    print(f"STEP 2: Transcribing {len(audio_files)} audio file(s) with Sarvam AI...")
    print(f"============================================================")
    
    # Set environment variable or pass parameters to sarvam_transcribe_to_docx.py
    import sarvam_transcribe_to_docx as transcriber
    transcriber.AUDIO_DIR = target_audio_dir
    transcriber.OUTPUT_DOCX = args.output_docx
    transcriber.main()


if __name__ == "__main__":
    main()
