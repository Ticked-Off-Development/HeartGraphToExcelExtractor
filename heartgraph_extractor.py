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


def _ocr_crop(image_path, top_frac, bottom_frac, binarize=True, left_frac=0.0):
    """OCR a rectangular region of the image.

    top_frac / bottom_frac: vertical bounds (0 → 1, fraction of image height).
    left_frac: left edge (0 → 1, fraction of image width); right edge is always
        the full image width.  Defaults to 0 (full width).  Pass e.g. 0.60 to
        restrict to the rightmost 40 % — useful for isolating the time column of
        the zone table without ⓘ icon interference.

    binarize=True (default): converts to greyscale then applies a fixed
        threshold (luminance > 160 → white).  Works well when the text sits on
        a uniform teal grid background (luminance ≳ 190).

    binarize=False: passes raw greyscale to Tesseract and lets it use its own
        adaptive (Otsu) thresholding.  Better when text sits on coloured zone
        bands whose luminance varies widely (red ≈ 120, blue ≈ 130, teal ≈ 195)
        — a fixed threshold turns those bands solid black and masks the text.
    """
    img = Image.open(image_path)
    crop = img.crop((
        int(img.width * left_frac),
        int(img.height * top_frac),
        img.width,
        int(img.height * bottom_frac),
    ))
    gray = crop.convert('L')
    if binarize:
        gray = gray.point(lambda x: 255 if x > 160 else 0)
    return pytesseract.image_to_string(gray)


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

    def _plausible(t):
        """Return True if t is a plausible zone duration.

        Rejects H:MM:SS times where H > 23.  A 24-hour HeartGraph session
        cannot have more than 23 h 59 min in any single zone.  H > 23 arises
        when the ⓘ info icon is OCR'd as a digit and fused with the leading
        digit of the hour (e.g. ⓘ + "11:08:35" → "41:08:35").
        """
        parts = t.split(':')
        if len(parts) == 3:
            if int(parts[0]) > 23:
                return False
            # Allow SS=60: HeartWatch occasionally displays 60 seconds due to
            # rounding (e.g. "21:09:60").  The excel-serial conversion handles
            # it correctly (same total seconds as 21:10:00).
            if int(parts[1]) > 59 or int(parts[2]) > 60:
                return False
        elif len(parts) == 2:
            if int(parts[1]) > 59:
                return False
        return True

    zone_data = []
    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Strategy 1: line has percentage AND time (e.g., "0.3% 4:48" or "85.1% 12:03:22")
        # Also allow '°' which OCR sometimes substitutes for '%'
        #
        # The [©il]? prefix absorbs the ⓘ info icon when it is OCR'd as a
        # letter ('i', 'l') or copyright symbol ('©') and fused directly with
        # the leading digit of the time, producing e.g. "88.2% i21:09:60".
        # The existing (?:.*?\s+)? handles ⓘ when a space follows it ("© ");
        # [©il]? handles the no-space fused case.
        match = re.search(
            r'(\d{1,3}\.?\d*)\s*[%°]\s+(?:.*?\s+)?[©il]?(\d{1,2}:\d{2}(?::\d{2})?)', line
        )
        if match and _plausible(match.group(2)):
            zone_data.append(match.group(2))
            continue

        # Strategy 2: line contains no letters and ends with a time (H:MM:SS or M:SS).
        # Handles OCR-dropped '%' (e.g. "0.3 4:17") and bare time-only lines.
        #
        # The previous pattern '^[^a-zA-Z]*(\d{1,2}:\d{2}...)' had a backtracking
        # bug: the greedy [^a-zA-Z]* consumed leading digits (e.g. the "1" in
        # "10:12" or "22:3" in "22:38:39"), leaving a shorter tail ("0:12", "8:39")
        # for the capture group.  Fix: check for letters separately, then use a
        # plain end-anchored search so re finds the time from the leftmost digit.
        #
        # Also exclude whole-hour clock times (e.g. "3:00") — x-axis labels.
        if not re.search(r'[a-zA-Z]', line):
            m2 = re.search(r'(\d{1,2}:\d{2}(?::\d{2})?)\s*$', line)
            if m2 and not re.match(r'^[1-9]\d?:00$', m2.group(1)) and _plausible(m2.group(1)):
                zone_data.append(m2.group(1))
                continue

        # Strategy 2b: ⓘ info icon OCR'd as a letter ('i', 'l') or '©' and
        # fused at the start of an otherwise letter-free time-only line (e.g.
        # "i21:09:60" in a right-side-only crop that excluded the percentage
        # column).  Strategy 2 skips the line because it sees a letter; this
        # strategy strips the single leading icon character and re-validates.
        m2b = re.search(r'^[©il](\d{1,2}:\d{2}(?::\d{2})?)\s*$', line)
        if m2b and _plausible(m2b.group(1)):
            zone_data.append(m2b.group(1))
            continue

        # Strategy 3: line has a zone marker symbol and a time (M:SS or H:MM:SS)
        match = re.search(r'[©@®⑤④③②①\(\)]\s*.*?(\d{1,2}:\d{2}(?::\d{2})?)\s*$', line)
        if match and _plausible(match.group(1)):
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

    # Format 1: "Heart rate range: 44 - 85" (newer) → max HR = 85
    # Forward search handles labels-before-values OCR column order.
    # re.DOTALL + non-greedy .*? bridges the gap when Tesseract emits all
    # labels first and all values second in a two-column layout.
    match = re.search(
        r'(?:Heart\s+rate\s+range|rate\s+range).*?(\d{2,3})\s*[-–—]\s*(\d{2,3})',
        text,
        re.IGNORECASE | re.DOTALL,
    )
    if match:
        summary['max_hr'] = int(match.group(2))

    # Backward search for Format 1 (values-before-labels column order):
    # Tesseract sometimes reads the right column (values) before the left
    # column (labels), so "46 - 109" appears *before* "Heart rate range:" in
    # the OCR text and the forward search above finds nothing after the label.
    # Look for the range pair "XX - YYY" in the text that precedes the label.
    if 'max_hr' not in summary:
        label_m = re.search(r'(?:Heart\s+rate\s+range|rate\s+range)', text, re.IGNORECASE)
        if label_m:
            # Use the LAST match before the label — the first match may be an
            # unrelated number pair from graph axis labels earlier in the OCR
            # text; the true HR range sits immediately before the label.
            matches = list(re.finditer(
                r'(\d{2,3})\s*[-–—]\s*(\d{2,3})',
                text[:label_m.start()],
            ))
            if matches:
                summary['max_hr'] = int(matches[-1].group(2))

    # Format 2: "Maximum heart rate: YYY" (older) → max HR = YYY
    #
    # Uses the same DOTALL + range-check approach as Format 1 to handle all
    # three OCR column orderings:
    #   single-line    "Maximum heart rate: 75\n..."
    #   labels-first   "Maximum heart rate:\nMean heart rate:\n...\n75\n55.8"
    #   values-first   "75\n55.8\n...\nMaximum heart rate:\nMean heart rate:"
    #
    # For labels-first, [:\s]+(\d+) stopped at the 'M' of "Mean" and failed.
    # DOTALL lets .*? cross line-boundaries; taking the first \d{2,3} in the
    # plausible HR range [60–220] skips Duration time components (4, 24, 06)
    # that all fall below 60.
    if 'max_hr' not in summary:
        label_m = re.search(
            r'(?:Maximum\s+heart\s+rate|Max\w*\s+heart\s+rate)',
            text,
            re.IGNORECASE,
        )
        if label_m:
            # Forward: scan text after the label
            for n in re.findall(r'\b(\d{2,3})\b', text[label_m.end():]):
                if 60 <= int(n) <= 220:
                    summary['max_hr'] = int(n)
                    break

    # Backward search for Format 2 (values-before-labels column order):
    # max HR integer sits before the label; take the first plausible value.
    if 'max_hr' not in summary:
        label_m = re.search(
            r'(?:Maximum\s+heart\s+rate|Max\w*\s+heart\s+rate)',
            text,
            re.IGNORECASE,
        )
        if label_m:
            integers = [int(n) for n in re.findall(r'\b(\d{2,3})\b', text[:label_m.start()])]
            candidates = [v for v in integers if 60 <= v <= 220]
            if candidates:
                summary['max_hr'] = candidates[0]

    # Mean heart rate — three strategies in priority order.
    #
    # Strategy 1 (normal): decimal appears after the label.  Accept both '.'
    # and ',' as the decimal separator (OCR occasionally substitutes a comma).
    # Guards:
    #   • [30–200]: prevents DOTALL from grabbing spurious large graph values
    #     (e.g. "560.2" when the real value is "50.2").
    #   • val < max_hr: mean HR must be strictly less than max HR
    #     (physiological law).  Without this, DOTALL can scan past the actual
    #     mean HR integer (no decimal point) and land on the max HR expressed
    #     as "90.0" or similar, making mean == max.
    match = re.search(
        r'Mean\s+heart.*?(\d+[.,]\d+)',
        text,
        re.IGNORECASE | re.DOTALL,
    )
    if match:
        val = float(match.group(1).replace(',', '.'))
        if 30 <= val <= 200 and ('max_hr' not in summary or val < summary['max_hr']):
            summary['mean_hr'] = val

    # Strategy 2 (values-before-labels): Tesseract read the right column
    # first, so "59.7" sits before "Mean heart rate:" in the OCR text and
    # the forward search above misses it.  Take the last decimal in the
    # text that precedes the label and is a plausible HR (30–200 bpm).
    if 'mean_hr' not in summary:
        label_m = re.search(r'Mean\s+heart', text, re.IGNORECASE)
        if label_m:
            floats = re.findall(r'\b(\d+[.,]\d+)\b', text[:label_m.start()])
            if floats:
                val = float(floats[-1].replace(',', '.'))
                if 30 <= val <= 200 and ('max_hr' not in summary or val < summary['max_hr']):
                    summary['mean_hr'] = val

    # Strategy 3 (decimal dropped, label-first): OCR lost the decimal separator
    # (e.g. "59.7" → "59").  Accept a 2–3 digit whole number on the same OCR
    # line as "Mean heart rate:" that falls in a plausible HR range.
    if 'mean_hr' not in summary:
        match = re.search(
            r'Mean\s+heart\s+rate\s*:?\s*(\d{2,3})\b',
            text,
            re.IGNORECASE,
        )
        if match:
            val = int(match.group(1))
            if 30 <= val <= 200 and ('max_hr' not in summary or val < summary['max_hr']):
                summary['mean_hr'] = float(val)

    # Strategy 4 (decimal dropped, values-before-labels): Tesseract read the
    # right column first AND dropped the decimal point, so the integer mean HR
    # sits before the "Mean heart rate:" label and strategies 1–3 all miss it.
    # Take the last plausible integer in the text preceding the label.
    #
    # Exclude HR-range bounds: in a values-first layout the range "48 - 95"
    # also appears before the label.  Excluding only max_hr left the min bound
    # (48) as the last candidate, producing mean_hr == min_hr.  Instead, detect
    # every lo–hi pair in the prefix and exclude both numbers from candidates.
    if 'mean_hr' not in summary:
        label_m = re.search(r'Mean\s+heart', text, re.IGNORECASE)
        if label_m:
            prefix = text[:label_m.start()]
            range_values = set()
            for lo, hi in re.findall(r'\b(\d{2,3})\s*[-\u2013\u2014]\s*(\d{2,3})\b', prefix):
                range_values.add(int(lo))
                range_values.add(int(hi))
            integers = [int(n) for n in re.findall(r'\b(\d{2,3})\b', prefix)]
            candidates = [v for v in integers
                          if 30 <= v <= 200 and v not in range_values
                          and ('max_hr' not in summary or v < summary['max_hr'])]
            if candidates:
                summary['mean_hr'] = float(candidates[-1])

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
    # "Zones" (plural) is the page title — unique to the zones screen.
    if re.search(r'\bZones\b', text, re.IGNORECASE):
        return 'zones'
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
        return f"0:{parts[0].zfill(2)}:{parts[1].zfill(2)}"
    elif len(parts) == 3:
        return f"{parts[0]}:{parts[1].zfill(2)}:{parts[2].zfill(2)}"
    return time_str


def _time_str_to_excel_serial(time_str):
    """Convert a H:MM:SS string to an Excel time serial number (fraction of a day).

    Excel stores times as fractions of 24 hours (e.g. 6 hours = 0.25).
    Writing the numeric value instead of a string lets Excel treat the cell
    as a true duration without the user having to press Enter to re-evaluate.
    """
    parts = time_str.split(':')
    try:
        if len(parts) == 3:
            total_seconds = int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
        elif len(parts) == 2:
            total_seconds = int(parts[0]) * 60 + int(parts[1])
        else:
            return None
    except ValueError:
        return None
    return total_seconds / 86400


def _process_folder(folder, year, image_extensions, daily_data):
    """Process all images in a single folder, appending results to daily_data.

    year - int year to use for dates, or None to fall back to current year.
    """
    all_files = sorted(folder.iterdir())
    image_files = sorted([f for f in all_files if f.is_file() and f.suffix.lower() in image_extensions])
    if not image_files:
        print(f"  ⚠ No image files found in: {folder}")
        return

    non_image = [f.name for f in all_files if f.is_file() and f.suffix.lower() not in image_extensions]
    if non_image:
        print(f"  Files found but not matched as images: {non_image}")

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

            # Second fallback: the zone table sits directly on the coloured
            # zone-band graph background.  Binarising at a fixed threshold of
            # 160 turns the coloured bands (red ≈ 120, blue ≈ 130) black and
            # masks the overlaid text.  Re-OCR without binarisation so Tesseract
            # can apply its own adaptive (Otsu) threshold to handle the varying
            # background.  Also widen the crop to start at 4 % so the
            # "Zone"/"Time" column headers at ~7.5 % are not clipped.
            no_binary_text = None
            if screenshot_type == 'unknown':
                no_binary_text = _ocr_crop(img_path, 0.04, 0.32, binarize=False)
                screenshot_type = classify_screenshot(no_binary_text)
                if screenshot_type == 'unknown':
                    time_matches = re.findall(
                        r'\b\d{1,2}:\d{2}(?::\d{2})?\b', no_binary_text
                    )
                    if len(time_matches) >= 3:
                        screenshot_type = 'zones'

            # Third fallback: check the footer strip (88–100 %) for
            # "Set Reference" — a button unique to the zones page that sits
            # *below* the coloured zone bands and is reliably readable.
            if screenshot_type == 'unknown':
                footer_text = _ocr_crop(img_path, 0.88, 1.0)
                if re.search(r'Set\s+Reference', footer_text, re.IGNORECASE):
                    screenshot_type = 'zones'

            # Fourth fallback specifically for zones pages whose zone-row
            # statistics sit in the lower half of the screen (~45–85 % of
            # height), well below the narrow 8–32 % strip.
            zones_region_text = None
            if screenshot_type == 'unknown':
                zones_region_text = _ocr_crop(img_path, 0.45, 0.85)
                screenshot_type = classify_screenshot(zones_region_text)
                if screenshot_type == 'unknown':
                    time_matches = re.findall(
                        r'\b\d{1,2}:\d{2}(?::\d{2})?\b', zones_region_text
                    )
                    if len(time_matches) >= 3:
                        screenshot_type = 'zones'

            print(f"  Date: {date.strftime('%A, %B %d, %Y')} | Type: {screenshot_type}")

            if date_key not in daily_data:
                daily_data[date_key] = {
                    'date': date,
                    'zone_sessions': [],
                    'summary_sessions': [],
                }

            if screenshot_type == 'zones':
                # Try every text source we collected and keep the best result.
                # Order: non-binarized crop first (handles coloured zone-band
                # backgrounds), then lower-half crop, then binarized crop.
                # If nothing was collected yet, generate a fresh non-binarized
                # crop now.
                candidate_sources = [
                    s for s in [no_binary_text, zones_region_text, data_region_text]
                    if s is not None
                ]
                if not candidate_sources:
                    # Image was classified from the full frame without generating
                    # any crops.  Try both the top strip (zone table above the
                    # graph) and the lower half (zone table below the graph),
                    # since the layout varies across HeartGraph versions.
                    candidate_sources = [
                        _ocr_crop(img_path, 0.04, 0.32, binarize=False),
                        _ocr_crop(img_path, 0.45, 0.85, binarize=False),
                    ]
                zones = {}
                for src in candidate_sources:
                    candidate = extract_zones(src)
                    if len(candidate) > len(zones):
                        zones = candidate
                    if len(zones) == 5:
                        break
                if len(zones) < 5:
                    # Right-side-only crop: the ⓘ info icon sits at ~72–74 % of
                    # image width; the time column starts at ~76 %.  Cropping
                    # from 74 % excludes the icon while keeping all five zone
                    # times, preventing OCR from fusing ⓘ with the hour digit
                    # (e.g. ⓘ + "11:08:35" → "41:08:35").  Fall back to a
                    # wider 55 % crop if needed.  Try both vertical extents to
                    # cover zone-table-above and -below layouts.
                    for left_frac in [0.74, 0.55]:
                        for top, bot in [(0.04, 0.40), (0.40, 0.88)]:
                            right_text = _ocr_crop(img_path, top, bot, binarize=False, left_frac=left_frac)
                            candidate = extract_zones(right_text)
                            if len(candidate) > len(zones):
                                zones = candidate
                            if len(zones) == 5:
                                break
                        if len(zones) == 5:
                            break
                if len(zones) < 5:
                    # Last resort: full-image text.
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

                # If the narrow crop returned only one of {max_hr, mean_hr},
                # try progressively larger sources and *merge* any new values
                # rather than replacing.  The narrow crop is preferred because
                # the full image includes graph Y-axis numbers that can
                # interfere with extraction even after range-guarding.
                if len(summary) < 2:
                    # Step 1: slightly wider crop — catches stats that sit just
                    # outside the 8-32 % band in some UI layouts.
                    wider = _ocr_crop(img_path, 0.06, 0.38)
                    for k, v in extract_summary(wider).items():
                        if k not in summary:
                            summary[k] = v

                if len(summary) < 2:
                    # Step 1b: non-binarised tight top strip.  The fixed
                    # binarisation threshold (luminance > 160) sits right on
                    # the edge of the app's teal background (~161 luminance),
                    # so some pixels binarise incorrectly and OCR garbles the
                    # value (e.g. "57.5" → "57 5").  Re-cropping without
                    # binarisation lets Tesseract use its own adaptive Otsu
                    # threshold and typically recovers the decimal correctly.
                    # Capping at 0.22 excludes the heart-rate graph (which
                    # begins at ~22 %) so coloured zone bands don't interfere.
                    top_stats = _ocr_crop(img_path, 0.06, 0.22, binarize=False)
                    for k, v in extract_summary(top_stats).items():
                        if k not in summary:
                            summary[k] = v

                if len(summary) < 2:
                    # Step 2: bottom stats strip.  The heart-rate graph
                    # occupies roughly the upper 65 % of the summary screen;
                    # the session stats (Maximum / Mean heart rate) sit below
                    # it.  Targeting this band directly avoids the Y-axis
                    # labels (100, 75, 50, 25) and graph-area OCR noise that
                    # contaminate the full-image extraction.
                    bottom_stats = _ocr_crop(img_path, 0.65, 0.95)
                    for k, v in extract_summary(bottom_stats).items():
                        if k not in summary:
                            summary[k] = v

                if len(summary) < 2:
                    # Step 3: full-image text as last resort.
                    for k, v in extract_summary(text).items():
                        if k not in summary:
                            summary[k] = v

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
    if not sorted_dates:
        wb.save(output_path)
        return 0

    # Iterate over every calendar day from first to last so that days with no
    # screenshots get a date row with blank values instead of being omitted.
    first_date = datetime.strptime(sorted_dates[0], '%Y-%m-%d')
    last_date = datetime.strptime(sorted_dates[-1], '%Y-%m-%d')
    total_days = (last_date - first_date).days + 1

    for row_idx, day_offset in enumerate(range(total_days), 3):
        current_date = first_date + timedelta(days=day_offset)
        date_key = current_date.strftime('%Y-%m-%d')

        # Date column — written for every day, data or blank
        cell = ws.cell(row=row_idx, column=1, value=current_date.strftime('%A, %B %d, %Y'))
        cell.font = date_font
        cell.alignment = Alignment(horizontal='right')

        if date_key not in daily_data:
            continue  # leave B-H empty for missing days

        data = daily_data[date_key]
        zones, summary, _ = aggregate_day(data)

        # Zone times (B-F)
        zone_map = {'B': 'zone5', 'C': 'zone4', 'D': 'zone3', 'E': 'zone2', 'F': 'zone1'}
        for col_letter, zone_key in zone_map.items():
            if zone_key in zones:
                time_str = normalize_time(zones[zone_key])
                excel_val = _time_str_to_excel_serial(time_str)
                cell = ws.cell(row=row_idx, column=ord(col_letter) - 64,
                               value=excel_val if excel_val is not None else time_str)
                if excel_val is not None:
                    cell.number_format = 'h:mm:ss'
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
