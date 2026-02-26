"""
HeartGraph Screenshot → Excel Extractor

Reads HeartGraph Premium screenshots and extracts heart rate zone data
into an Excel spreadsheet matching Carol's HRM Daily format.

USAGE:
    python heartgraph_extractor.py <screenshot_folder> [output.xlsx]

SCREENSHOT NAMING:
    - Files can be named anything (.png, .jpg, .jpeg)
    - The script reads the date from the screenshot header (e.g., "20 Feb 22:46")
    - Two screenshots per day: one with zone times, one with summary stats
    - Screenshots from the same date are automatically paired

FOLDER STRUCTURE (year-based):
    Because the screenshot header omits the year, images should be placed in
    year-named subfolders so the correct year is used automatically:

        screenshots/
            2025/   ← images taken in 2025
            2026/   ← images taken in 2026

    The script can also be pointed directly at a year folder:
        python heartgraph_extractor.py screenshots/2026

DATE OFFSET:
    - Since tracking starts at 10:45 PM, a screenshot dated "20 Feb"
      corresponds to February 21 in the spreadsheet (next day).

OUTPUT COLUMNS:
    A: Date | B: red (Zone 5) | C: orange (Zone 4) | D: yellow (Zone 3)
    E: green (Zone 2) | F: blue (Zone 1) | G: Max HR | H: Mean 24hr HR
"""

import pytesseract
from PIL import Image
import re
import os
import sys
from datetime import datetime, timedelta
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# On Windows, Tesseract is typically installed to a fixed location.
# Set the path automatically if it exists so users don't need to add it to PATH manually.
_win_tesseract = Path(r'C:\Program Files\Tesseract-OCR\tesseract.exe')
if sys.platform == 'win32' and _win_tesseract.exists():
    pytesseract.pytesseract.tesseract_cmd = str(_win_tesseract)


def _check_tesseract():
    """Verify Tesseract is available and print a helpful message if not."""
    try:
        pytesseract.get_tesseract_version()
    except pytesseract.TesseractNotFoundError:
        print("ERROR: Tesseract OCR is not installed or cannot be found.")
        print()
        print("Install it from: https://github.com/UB-Mannheim/tesseract/wiki")
        print("  (choose the Windows 64-bit installer)")
        print()
        print("After installing, either:")
        print("  • Restart this terminal so PATH is updated, or")
        print("  • Install to the default location:")
        print("    C:\\Program Files\\Tesseract-OCR\\  (detected automatically)")
        sys.exit(1)


def extract_text(image_path):
    """Extract text from a screenshot using OCR."""
    img = Image.open(image_path)
    return pytesseract.image_to_string(img)


def _ocr_crop(image_path, top_frac, bottom_frac):
    """OCR a horizontal strip of the image between top_frac and bottom_frac (0→1).

    Converts to greyscale and binarises before OCR.  The zone table and summary
    stats sit on top of a coloured graph-paper grid (teal lines, zone-band fills).
    Text luminance is ≲160; grid/background luminance is ≳190.  Binarising at
    160 gives Tesseract a clean black-on-white image and dramatically improves
    recognition on these backgrounds.
    """
    img = Image.open(image_path)
    crop = img.crop((0, int(img.height * top_frac), img.width, int(img.height * bottom_frac)))
    gray = crop.convert('L')
    binary = gray.point(lambda x: 255 if x > 160 else 0)
    return pytesseract.image_to_string(binary)


def extract_date(image_path, year=None):
    """Extract the date from the screenshot header by cropping the top.

    year - if provided, use this year; otherwise fall back to the current
           calendar year (legacy flat-folder behaviour).
    """
    img = Image.open(image_path)
    header = img.crop((0, 0, img.width, int(img.height * 0.12)))
    header_text = pytesseract.image_to_string(header)

    # Fall back to full image if header crop yields nothing useful
    full_text = pytesseract.image_to_string(img)

    date_pattern = r'(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2}):(\d{2})'

    # Try header crop first, then full image
    match = re.search(date_pattern, header_text, re.IGNORECASE)
    if not match:
        match = re.search(date_pattern, full_text, re.IGNORECASE)

    if not match:
        print(f"    Header OCR output: {repr(header_text.strip())}")

    if match:
        day = int(match.group(1))
        month_str = match.group(2)
        hour = int(match.group(3))
        month_map = {
            'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
            'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
        }
        month = month_map[month_str.lower()]
        if year is None:
            year = datetime.now().year

        screenshot_date = datetime(year, month, day)

        # Only add +1 day for evening sessions (started at ~10:45 PM).
        # If the session started during the day (e.g., 12:56), the date
        # on the screenshot is already the correct tracking date.
        if hour >= 20:
            actual_date = screenshot_date + timedelta(days=1)
        else:
            actual_date = screenshot_date
        return actual_date

    return None


def _year_from_name(name):
    """Return the year as int if name is a 4-digit year string (19xx or 20xx), else None."""
    m = re.fullmatch(r'((?:19|20)\d{2})', name)
    return int(m.group(1)) if m else None


def extract_zones(text):
    """Extract zone times from the zones screenshot.

    Zones in HeartGraph:
        Zone 5 (red)    → highest HR
        Zone 4 (orange)
        Zone 3 (yellow)
        Zone 2 (green)
        Zone 1 (blue)   → lowest HR

    Returns dict with keys: zone5, zone4, zone3, zone2, zone1
    """
    zones = {}
    lines = text.split('\n')

    zone_data = []
    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Strategy 1: line has percentage AND time (e.g., "0.3% 4:48" or "85.1% 12:03:22")
        # Also allow '°' which OCR sometimes substitutes for '%'
        match = re.search(
            r'(\d{1,3}\.?\d*)\s*[%°]\s+(?:.*?\s+)?(\d{1,2}:\d{2}(?::\d{2})?)', line
        )
        if match:
            zone_data.append(match.group(2))
            continue

        # Strategy 2: line is mostly non-letter characters ending with a time (H:MM:SS or M:SS).
        # Handles cases where OCR drops the '%' entirely, e.g. "0.3 4:17" or "@ 8:18:31".
        match = re.search(r'^[^a-zA-Z]*(\d{1,2}:\d{2}(?::\d{2})?)\s*$', line)
        if match:
            zone_data.append(match.group(1))
            continue

        # Strategy 3: line has a zone marker symbol and a time (M:SS or H:MM:SS)
        match = re.search(r'[©@®⑤④③②①\(\)]\s*.*?(\d{1,2}:\d{2}(?::\d{2})?)\s*$', line)
        if match:
            zone_data.append(match.group(1))
            continue

    if len(zone_data) >= 5:
        # Take the first 5 matches - zones are listed top-to-bottom: 5, 4, 3, 2, 1
        zones['zone5'] = zone_data[0]
        zones['zone4'] = zone_data[1]
        zones['zone3'] = zone_data[2]
        zones['zone2'] = zone_data[3]
        zones['zone1'] = zone_data[4]
    elif len(zone_data) > 0:
        # Partial extraction - log what we got for debugging
        print(f"    ⚠ Only found {len(zone_data)} zone values: {zone_data}")

    return zones


def extract_summary(text):
    """Extract summary stats from the summary screenshot.

    Looks for:
        - Heart rate range: XX - YYY (max HR = YYY)
        - Mean heart rate: XX.X
    """
    summary = {}

    # Format 1: "Heart rate range: XX - YYY" (newer) → max HR = YYY
    match = re.search(
        r'(?:Heart\s+rate\s+range|rate\s+range)[:\s]+(\d+)\s*[-–]\s*(\d+)',
        text,
        re.IGNORECASE,
    )
    if match:
        summary['max_hr'] = int(match.group(2))

    # Format 2: "Maximum heart rate: YYY" (older) → max HR = YYY
    if 'max_hr' not in summary:
        match = re.search(
            r'(?:Maximum\s+heart\s+rate|Max\w*\s+heart\s+rate)[:\s]+(\d+)',
            text,
            re.IGNORECASE,
        )
        if match:
            summary['max_hr'] = int(match.group(1))

    # Mean heart rate
    match = re.search(
        r'(?:Mean\s+heart\s+rate|mean\s+rate)[:\s]+(\d+\.?\d*)', text, re.IGNORECASE
    )
    if match:
        summary['mean_hr'] = float(match.group(1))

    return summary


def classify_screenshot(text):
    """Determine if a screenshot is a 'zones' or 'summary' type."""
    # --- Summary indicators checked first (unambiguous; absent from zones screens) ---
    if re.search(r'\bDuration\b', text, re.IGNORECASE):
        return 'summary'
    if re.search(r'Heart\s+rate\s+range', text, re.IGNORECASE):
        return 'summary'
    if re.search(r'Maximum\s+heart\s+rate', text, re.IGNORECASE):
        return 'summary'
    if re.search(r'Mean\s+heart', text, re.IGNORECASE):
        return 'summary'

    # --- Zones indicators ---
    # "Zone" and "Time" on the same line (compact layout)
    if re.search(r'Zone\s+Time', text, re.IGNORECASE):
        return 'zones'
    # Circled zone-number characters (⑤④③②①) are unique to the zones table
    if re.search(r'[⑤④③②①]', text):
        return 'zones'
    # "Zone" header present + percentages (OCR split "Zone" / "Time" onto separate lines,
    # and also tolerate '°' which OCR sometimes substitutes for '%')
    pct_count = text.count('%') + text.count('°')
    if re.search(r'\bZone\b', text, re.IGNORECASE) and pct_count >= 3:
        return 'zones'
    # Fallback percentage count (tolerate '°' for '%')
    if pct_count >= 4:
        return 'zones'
    # "Set Reference" / "Zoom" footer buttons appear on both zones and summary screens.
    # All summary keywords have already been checked above, so reaching here means
    # this is almost certainly a zones screen.
    if re.search(r'Set\s+Reference', text, re.IGNORECASE):
        return 'zones'
    if re.search(r'\bZoom\b', text) and re.search(r'\bZone\b', text, re.IGNORECASE):
        return 'zones'
    return 'unknown'


def normalize_time(time_str):
    """Ensure time is in H:MM:SS format for the spreadsheet."""
    parts = time_str.split(':')
    if len(parts) == 2:
        # M:SS → 0:MM:SS
        return f"0:{parts[0].zfill(2)}:{parts[1]}"
    elif len(parts) == 3:
        return f"{parts[0]}:{parts[1].zfill(2)}:{parts[2].zfill(2)}"
    return time_str


def _process_folder(folder, year, image_extensions, daily_data):
    """Process all images in a single folder, appending results to daily_data.

    year - int year to use for dates, or None to fall back to current year.
    """
    all_files = sorted(folder.iterdir())
    if not all_files:
        print(f"  ⚠ Folder appears to be empty or inaccessible: {folder}")
        return

    non_image = [f.name for f in all_files if f.suffix.lower() not in image_extensions and f.is_file()]
    if non_image:
        print(f"  Files found but not matched as images: {non_image}")

    image_files = sorted([f for f in all_files if f.suffix.lower() in image_extensions])

    year_label = str(year) if year is not None else "current year"
    print(f"Found {len(image_files)} image files to process (year: {year_label})...")

    for img_path in image_files:
        print(f"\nProcessing: {img_path.name}")

        try:
            text = extract_text(img_path)
            date = extract_date(img_path, year=year)

            if date is None:
                print(f"  ⚠ Could not extract date, skipping")
                continue

            date_key = date.strftime('%Y-%m-%d')
            screenshot_type = classify_screenshot(text)

            # When full-image OCR can't classify (the large heart-rate graph
            # dominates and buries the small zone-table / stats text), crop
            # just the data region — below the nav bar, above the graph
            # (~8-32 % of image height).  Both the zone table and the summary
            # stats live in that strip, so a single crop covers both types.
            data_region_text = None
            if screenshot_type == 'unknown':
                data_region_text = _ocr_crop(img_path, 0.08, 0.32)
                screenshot_type = classify_screenshot(data_region_text)

            if screenshot_type == 'unknown' and data_region_text is not None:
                # Last-resort: keyword patterns still failed (e.g. "Zone"/"Time"
                # garbled by the grid).  Count time-format values in the data
                # region: a zones screen has 5 (one per zone row) while a
                # summary screen has at most 1 (Duration).  No x-axis labels
                # appear in this narrow crop so false matches aren't a concern.
                time_matches = re.findall(r'\b\d{1,2}:\d{2}(?::\d{2})?\b', data_region_text)
                if len(time_matches) >= 5:
                    screenshot_type = 'zones'

            print(f"  Date: {date.strftime('%A, %B %d, %Y')} | Type: {screenshot_type}")

            if date_key not in daily_data:
                daily_data[date_key] = {
                    'date': date,
                    'zone_sessions': [],
                    'summary_sessions': [],
                }

            if screenshot_type == 'zones':
                # Always extract from the data-region crop: full-image OCR
                # picks up graph x-axis time labels (e.g. "9:00") as false
                # zone times, while the crop contains only the zone table.
                if data_region_text is None:
                    data_region_text = _ocr_crop(img_path, 0.08, 0.32)
                zones = extract_zones(data_region_text)
                if len(zones) < 5:
                    # data-region OCR still missed some zones; try full-image
                    # text as a last resort.
                    zones_from_full = extract_zones(text)
                    if len(zones_from_full) > len(zones):
                        zones = zones_from_full
                if zones:
                    daily_data[date_key]['zone_sessions'].append(zones)
                    print(f"  Zones extracted: {zones}")
                else:
                    print(f"  ⚠ Could not parse zone data")

            elif screenshot_type == 'summary':
                # Use the binarised data-region crop for extraction; the same
                # grid noise that breaks zone extraction also garbles numbers
                # in "Heart rate range: 45 - 95" / "Mean heart rate: 57.4".
                if data_region_text is None:
                    data_region_text = _ocr_crop(img_path, 0.08, 0.32)
                summary = extract_summary(data_region_text)
                if not summary:
                    summary = extract_summary(text)
                if summary:
                    daily_data[date_key]['summary_sessions'].append(summary)
                    print(f"  Summary extracted: {summary}")
                else:
                    print(f"  ⚠ Could not parse summary data")
            else:
                print(f"  ⚠ Unknown screenshot type")

        except Exception as e:
            print(f"  ✗ Error: {e}")


def process_screenshots(folder_path):
    """Process screenshots and group by date.

    Supports three layouts:

    1. Year subfolders (recommended):
           screenshots/2025/  screenshots/2026/
       The year is taken from the subfolder name.

    2. Year folder passed directly:
           python heartgraph_extractor.py screenshots/2026
       The year is taken from the folder name.

    3. Flat folder (legacy):
           screenshots/IMG_001.png  …
       The current calendar year is used (may be wrong for old images).
    """
    folder = Path(folder_path)
    image_extensions = {'.png', '.jpg', '.jpeg', '.bmp', '.tiff'}
    daily_data = {}

    # Case 1: folder contains year-named subfolders
    year_subdirs = sorted(
        [d for d in folder.iterdir() if d.is_dir() and _year_from_name(d.name) is not None]
    )
    if year_subdirs:
        for subdir in year_subdirs:
            year = _year_from_name(subdir.name)
            print(f"\n--- Processing year {year} ({subdir.name}/) ---")
            _process_folder(subdir, year, image_extensions, daily_data)
        return daily_data

    # Case 2: the folder itself is a year folder
    folder_year = _year_from_name(folder.name)

    # Case 3 (flat) or Case 2: process images directly in the folder
    _process_folder(folder, folder_year, image_extensions, daily_data)
    return daily_data


def time_to_seconds(time_str):
    """Convert H:MM:SS or M:SS to total seconds."""
    parts = time_str.split(':')
    if len(parts) == 3:
        return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
    elif len(parts) == 2:
        return int(parts[0]) * 60 + int(parts[1])
    return 0


def seconds_to_time(total_seconds):
    """Convert total seconds to H:MM:SS string."""
    h = total_seconds // 3600
    m = (total_seconds % 3600) // 60
    s = total_seconds % 60
    return f"{h}:{m:02d}:{s:02d}"


def aggregate_day(data):
    """Aggregate multiple sessions for a single day.

    - Zone times: summed across all sessions
    - Mean HR: averaged across sessions
    - Max HR: highest value across sessions

    Returns (zones_dict, summary_dict, num_sessions)
    """
    zone_sessions = data.get('zone_sessions', [])
    summary_sessions = data.get('summary_sessions', [])

    # Sum zone times
    zones = {}
    if zone_sessions:
        for zone_key in ['zone5', 'zone4', 'zone3', 'zone2', 'zone1']:
            total_secs = 0
            for session in zone_sessions:
                if zone_key in session:
                    total_secs += time_to_seconds(session[zone_key])
            if total_secs > 0:
                zones[zone_key] = seconds_to_time(total_secs)
            else:
                # Still record zero if there were sessions (sensor was active, just no time in this zone)
                if zone_sessions:
                    zones[zone_key] = "0:00:00"

    # Aggregate summary stats
    summary = {}
    if summary_sessions:
        max_hrs = [s['max_hr'] for s in summary_sessions if 'max_hr' in s]
        mean_hrs = [s['mean_hr'] for s in summary_sessions if 'mean_hr' in s]

        if max_hrs:
            summary['max_hr'] = max(max_hrs)
        if mean_hrs:
            summary['mean_hr'] = round(sum(mean_hrs) / len(mean_hrs), 1)

    num_sessions = max(len(zone_sessions), len(summary_sessions))
    return zones, summary, num_sessions


def create_excel(daily_data, output_path):
    """Create an Excel workbook matching Carol's HRM Daily format."""
    wb = Workbook()
    ws = wb.active
    ws.title = "HRM Daily"

    # === Header styling ===
    header_font = Font(name='Arial', bold=True, color='008080', size=11)
    date_font = Font(name='Arial', size=11)
    data_font = Font(name='Arial', size=11)

    # Color fills matching the zone columns
    zone_fills = {
        'B': PatternFill('solid', fgColor='FFE0E0'),  # red zone
        'C': PatternFill('solid', fgColor='FFE8CC'),  # orange zone
        'D': PatternFill('solid', fgColor='FFFFCC'),  # yellow zone
        'E': PatternFill('solid', fgColor='CCFFCC'),  # green zone
        'F': PatternFill('solid', fgColor='CCE5FF'),  # blue zone
    }

    # === Title row ===
    ws.merge_cells('A1:H1')
    ws['A1'] = 'HRM Daily'
    ws['A1'].font = Font(name='Arial', bold=True, color='FFFFFF', size=14)
    ws['A1'].fill = PatternFill('solid', fgColor='00B0F0')
    ws['A1'].alignment = Alignment(horizontal='right')

    # === Headers (row 2) ===
    headers = ['Date', 'red', 'orange', 'yellow', 'green', 'blue', 'Max HR\n(bpm)', 'Mean 24 hr\nHR (bpm)']
    header_colors = ['00B0F0', 'FF8080', 'FFB366', 'FFFF80', '80FF80', '80B3FF', '00B0F0', '00B0F0']

    for col_idx, (header, color) in enumerate(zip(headers, header_colors), 1):
        cell = ws.cell(row=2, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = PatternFill('solid', fgColor=color)
        cell.alignment = Alignment(horizontal='center', wrap_text=True)

    # === Column widths ===
    ws.column_dimensions['A'].width = 32
    for col in ['B', 'C', 'D', 'E', 'F']:
        ws.column_dimensions[col].width = 14
    ws.column_dimensions['G'].width = 10
    ws.column_dimensions['H'].width = 12

    # === Data rows ===
    sorted_dates = sorted(daily_data.keys())

    for row_idx, date_key in enumerate(sorted_dates, 3):
        data = daily_data[date_key]
        date_obj = data['date']
        zones, summary, num_sessions = aggregate_day(data)

        # Date column
        cell = ws.cell(row=row_idx, column=1, value=date_obj.strftime('%A, %B %d, %Y'))
        cell.font = date_font
        cell.alignment = Alignment(horizontal='right')

        # Zone times (B-F)
        zone_map = {'B': 'zone5', 'C': 'zone4', 'D': 'zone3', 'E': 'zone2', 'F': 'zone1'}
        for col_letter, zone_key in zone_map.items():
            if zone_key in zones:
                time_val = normalize_time(zones[zone_key])
                cell = ws.cell(row=row_idx, column=ord(col_letter) - 64, value=time_val)
                cell.font = data_font
                cell.alignment = Alignment(horizontal='center')
                cell.fill = zone_fills[col_letter]

        # Max HR (G)
        if 'max_hr' in summary:
            cell = ws.cell(row=row_idx, column=7, value=summary['max_hr'])
            cell.font = data_font
            cell.alignment = Alignment(horizontal='center')

        # Mean HR (H)
        if 'mean_hr' in summary:
            cell = ws.cell(row=row_idx, column=8, value=summary['mean_hr'])
            cell.font = data_font
            cell.alignment = Alignment(horizontal='center')

    wb.save(output_path)
    return len(sorted_dates)


def main():
    _check_tesseract()

    if len(sys.argv) < 2:
        print("Usage: python heartgraph_extractor.py <screenshot_folder> [output.xlsx]")
        print()
        print("  screenshot_folder  - Folder containing HeartGraph screenshots")
        print("  output.xlsx        - Output Excel file (default: heartgraph_data.xlsx)")
        sys.exit(1)

    folder = sys.argv[1]
    output = sys.argv[2] if len(sys.argv) > 2 else 'heartgraph_data.xlsx'

    if not os.path.isdir(folder):
        print(f"Error: '{folder}' is not a valid directory")
        sys.exit(1)

    print("=" * 60)
    print("HeartGraph Screenshot Extractor")
    print("=" * 60)

    daily_data = process_screenshots(folder)

    if not daily_data:
        print("\nNo data extracted. Check that screenshots are valid HeartGraph exports.")
        sys.exit(1)

    # Report multi-session days
    multi_session_days = []
    for date_key in sorted(daily_data.keys()):
        data = daily_data[date_key]
        n_zones = len(data['zone_sessions'])
        n_summary = len(data['summary_sessions'])
        n = max(n_zones, n_summary)
        if n > 1:
            multi_session_days.append((data['date'], n))

    num_days = create_excel(daily_data, output)

    print(f"\n{'=' * 60}")
    print(f"✓ Successfully extracted {num_days} day(s) of data")
    if multi_session_days:
        print(f"  ℹ {len(multi_session_days)} day(s) had multiple sessions (zones summed, mean HR averaged, max HR = highest):")
        for d, n in multi_session_days:
            print(f"    • {d.strftime('%A, %B %d, %Y')}: {n} sessions")
    print(f"✓ Saved to: {output}")
    print(f"{'=' * 60}")


if __name__ == '__main__':
    main()
