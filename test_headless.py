"""Headless test — validates the SegmentPro engine and report generator
against the GAP dataset without launching the Tkinter UI."""

import sys
import os

# Make the app package importable
sys.path.insert(0, '/home/claude/app')

# Import the relevant pieces (without running main())
from SegmentPro import SegmentationEngine, ReportGenerator

INPUT = '/home/claude/data.xlsx'
OUT_DIR = '/home/claude/app/test_output'
os.makedirs(OUT_DIR, exist_ok=True)

def log(msg, level="info"):
    prefix = {"info": "  ", "success": "✓ ", "warn": "! ", "error": "✗ "}.get(level, "  ")
    print(f"{prefix}{msg}")

print("="*70)
print("STAGE 1: Load file")
print("="*70)
engine = SegmentationEngine(log_callback=log)
engine.load_file(INPUT)

print("\n"+"="*70)
print("STAGE 2: Detect demographics")
print("="*70)
demos = engine.detect_demographics()
print(f"  Demographics detected: {demos}")

print("\n"+"="*70)
print("STAGE 3: Auto-detect segmentation variables")
print("="*70)
detected = engine.auto_detect_seg_cols()
print(f"  Detected {len(detected)} variables")
print(f"  First 5: {detected[:5]}")

print("\n"+"="*70)
print("STAGE 4: Clean data")
print("="*70)
engine.clean_data(detected)

print("\n"+"="*70)
print("STAGE 5: Run clustering (auto k)")
print("="*70)
engine.run_clustering()

print(f"\n  Final k: {engine.k}")
print(f"  Segments:")
for seg, size in engine.segment_sizes().items():
    name = engine.seg_names.get(seg, "?")
    desc = engine.seg_descriptions.get(seg, "")[:80]
    print(f"    {seg}: {name} (n={size}) — {desc}...")

print("\n"+"="*70)
print("STAGE 6: Write Excel report")
print("="*70)
gen = ReportGenerator(engine, log_callback=log)
xlsx_path = os.path.join(OUT_DIR, "Test_Segments_Report.xlsx")
gen.write_excel(xlsx_path)
print(f"  Excel file size: {os.path.getsize(xlsx_path):,} bytes")

print("\n"+"="*70)
print("STAGE 7: Write PowerPoint deck")
print("="*70)
pptx_path = os.path.join(OUT_DIR, "Test_Segments_Deck.pptx")
gen.write_pptx(pptx_path)
print(f"  PPTX file size: {os.path.getsize(pptx_path):,} bytes")

print("\n"+"="*70)
print("ALL STAGES PASSED ✓")
print("="*70)
