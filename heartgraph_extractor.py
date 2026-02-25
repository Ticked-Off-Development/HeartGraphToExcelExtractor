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


def extract_text(image_path):
    """Extract text from a screenshot using OCR."""
    img = Image.open(image_path)
    return pytesseract.image_to_string(img)


def extract_date(image_path):
    """Extract the date from the screenshot header by cropping the top."""
    img = Image.open(image_path)
    header = img.crop((0, 0, img.width, int(img.height * 0.12)))
    text = pytesseract.image_to_string(header)

    # Look for patterns like "20 Feb 22:46" or "16 Dec 12:56"
    match = re.search(
        r'(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2}):(\d{2})',
        text,
        re.IGNORECASE,
    )
    if match:
        day = int(match.group(1))
        month_str = match.group(2)
        hour = int(match.group(3))
        month_map = {
            'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
            'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
        }
        month = month_map[month_str.lower()]
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
        match = re.search(
            r'(\d{1,3}\.?\d*)%\s+(?:.*?\s+)?(\d{1,2}:\d{2}(?::\d{2})?)', line
        )
        if match:
            zone_data.append(match.group(2))
            continue

        # Strategy 2: line has a zone marker (@, ©, ⑤, etc.) and a time but no percentage
        # e.g., "@ 8:18:31" or "© 1:42:51"
        # Only match if line looks like a zone row (short, with a time at the end)
        match = re.search(r'^[^a-zA-Z]*(\d{1,2}:\d{2}:\d{2})\s*$', line)
        if match:
            zone_data.append(match.group(1))
            continue

        # Strategy 3: line has just M:SS at the end (short zone times like "0:07", "8:45")
        # but only if we're already collecting zone data or line has zone-like markers
        match = re.search(r'[©@®⑤④③②①\(\)]\s*.*?(\d{1,2}:\d{2})\s*$', line)
        if match and ':' in match.group(1):
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
    if re.search(r'Zone\s+Time', text, re.IGNORECASE):
        return 'zones'
    if re.search(r'Duration:', text, re.IGNORECASE):
        return 'summary'
    if re.search(r'Maximum\s+heart\s+rate', text, re.IGNORECASE):
        return 'summary'
    # Fallback: if it has percentages with zone markers, it's zones
    if text.count('%') >= 4:
        return 'zones'
    if re.search(r'Mean\s+heart', text, re.IGNORECASE):
        return 'summary'
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


def process_screenshots(folder_path):
    """Process all screenshots in a folder and group by date."""
    folder = Path(folder_path)
    image_extensions = {'.png', '.jpg', '.jpeg', '.bmp', '.tiff'}

    # Collect all data grouped by date
    daily_data = {}

    image_files = sorted(
        [f for f in folder.iterdir() if f.suffix.lower() in image_extensions]
    )

    print(f"Found {len(image_files)} image files to process...")

    for img_path in image_files:
        print(f"\nProcessing: {img_path.name}")

        try:
            text = extract_text(img_path)
            date = extract_date(img_path)

            if date is None:
                print(f"  ⚠ Could not extract date, skipping")
                continue

            date_key = date.strftime('%Y-%m-%d')
            screenshot_type = classify_screenshot(text)
            print(f"  Date: {date.strftime('%A, %B %d, %Y')} | Type: {screenshot_type}")

            if date_key not in daily_data:
                daily_data[date_key] = {
                    'date': date,
                    'zone_sessions': [],
                    'summary_sessions': [],
                }

            if screenshot_type == 'zones':
                zones = extract_zones(text)
                if zones:
                    daily_data[date_key]['zone_sessions'].append(zones)
                    print(f"  Zones extracted: {zones}")
                else:
                    print(f"  ⚠ Could not parse zone data")

            elif screenshot_type == 'summary':
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
