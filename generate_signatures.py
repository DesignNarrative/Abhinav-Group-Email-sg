#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================
 Abhinav Group — Email Signature Generator
 Version: 2.0  |  Production Ready
=============================================================
 USAGE:
   Double-click  run_me.bat         (recommended for non-tech users)
   OR run:       python generate_signatures.py

 HOW UPDATES WORK:
   1. Open  employees/employees.xlsx
   2. Edit any row  (or add/delete rows)
   3. Save the Excel file
   4. Double-click  run_me.bat  again
   → All signatures are regenerated automatically

 ADDING A NEW EMPLOYEE:
   Open Excel  →  Add a new row  →  Run  run_me.bat

 REMOVING AN EMPLOYEE:
   Open Excel  →  Delete their row  →  Run  run_me.bat
   (also manually delete their .html file from output/ if needed)

 SECURITY NOTES:
   - All employee data is HTML-escaped before use (prevents injection)
   - No JavaScript in generated signatures
   - No external image links (all images Base64-embedded)
   - Excel file is opened in read-only mode
   - Output is isolated to output/ folder only
=============================================================
"""

import base64
import html as html_module
import io
import os
import re
import sys
from datetime import datetime
from pathlib import Path

# ── Dependency check ──────────────────────────────────────────────────────────
# Force UTF-8 output on Windows to handle any special chars in names
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

try:
    import openpyxl
except ImportError:
    print("\n[ERROR] Missing: openpyxl")
    print("  Fix:  pip install openpyxl")
    sys.exit(1)

try:
    from PIL import Image, ImageDraw
    PILLOW_AVAILABLE = True
except ImportError:
    PILLOW_AVAILABLE = False
    print("[WARNING] Pillow not found. Icons will not be resized to 28×28px.")
    print("          Install with:  pip install Pillow")

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "employees" / "employees.xlsx"
LOGO_PATH  = BASE_DIR / "images"    / "abhinav logo.png"
ICONS_DIR  = BASE_DIR / "icons"
OUTPUT_DIR = BASE_DIR / "output"

# ── Official social media links (update here if links ever change) ─────────────
SOCIAL_LINKS = {
    "website":   "https://abhinavgroup.co.in/",
    "facebook":  "https://www.facebook.com/AbhinavGroup/",
    "instagram": "https://www.instagram.com/abhinavgroup/?hl=en",
    "linkedin":  "https://www.linkedin.com/company/abhinavgroup/posts/?feedView=all",
    "youtube":   "https://www.youtube.com/@AbhinavGroupPune",
}

SOCIAL_ORDER = ["website", "facebook", "instagram", "linkedin", "youtube"]

SOCIAL_LABELS = {
    "website":   "Website",
    "facebook":  "Facebook",
    "instagram": "Instagram",
    "linkedin":  "LinkedIn",
    "youtube":   "YouTube",
}

# Possible filenames for each icon (tries each in order)
ICON_FILENAMES = {
    "website":   ["website.png",   "website.jpg",   "website.jpeg"],
    "facebook":  ["facebook.png",  "facebook.jpg",  "facebook.jpeg"],
    "instagram": ["instagram.png", "instagram.jpg", "instagram.jpeg"],
    "linkedin":  ["linkedin.png",  "linkedin.jpg",  "linkedin.jpeg"],
    "youtube":   ["youtube.png",   "youtube.jpg",   "youtube.jpeg"],
}

# ── Company address — same for every employee ─────────────────────────────────
ADDRESS = "FC Road, Shivajinagar, Pune"

# ── PNG contact icons — loaded from icons/ folder ──────────────────────────────
# Files: icons/phone.png  |  icons/email.png  |  icons/websiteglobe.png
# Each is a black-on-white PNG. We:
#   1. Make the white background transparent
#   2. Tint the dark pixels to Abhinav Group blue (#1b8ecb)
#   3. Resize to 15×15 px with LANCZOS (matches the HTML img size)
# PNG format — never causes Gmail attachment issues.

_BRAND_BLUE = (27, 142, 203, 255)   # #1b8ecb
_CONTACT_ICON_SIZE = 15              # matches width/height in HTML template


def _load_contact_icon(path: Path) -> str:
    """
    Load a black-on-white PNG icon, tint it to Abhinav blue,
    resize to _CONTACT_ICON_SIZE × _CONTACT_ICON_SIZE,
    return as Base64 PNG data URI.
    """
    from PIL import ImageOps

    # Open and convert to RGBA so we can manipulate every channel
    src = Image.open(path).convert("RGBA")

    # Work on a grayscale version to build the alpha mask:
    # dark pixels (the icon shape) → opaque  |  light pixels (background) → transparent
    gray = src.convert("L")
    alpha_mask = ImageOps.invert(gray)   # black icon → 255 (fully opaque) in mask

    # Create a flat image filled entirely with Abhinav blue
    blue = Image.new("RGBA", src.size, _BRAND_BLUE)

    # Apply the alpha mask — only the icon shape remains visible in blue
    blue.putalpha(alpha_mask)

    # Resize to final size with high-quality LANCZOS downsampling
    blue = blue.resize((_CONTACT_ICON_SIZE, _CONTACT_ICON_SIZE), Image.LANCZOS)

    buf = io.BytesIO()
    blue.save(buf, format="PNG", optimize=True)
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("ascii")


# Pre-compute contact icon data URIs once at module load
PHONE_ICON_URI = _load_contact_icon(ICONS_DIR / "phone.png")
EMAIL_ICON_URI = _load_contact_icon(ICONS_DIR / "email.png")
GLOBE_ICON_URI = _load_contact_icon(ICONS_DIR / "websiteglobe.png")


# ── Image encoder ─────────────────────────────────────────────────────────────
def encode_image_file(path: Path, resize: tuple = None) -> str:
    """
    Return a Base64 PNG data URI for an image file.
    If resize=(w, h) and Pillow is available, resize before encoding.
    """
    if PILLOW_AVAILABLE:
        img = Image.open(path).convert("RGBA")
        if resize:
            img = img.resize(resize, Image.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format="PNG", optimize=True)
        data = base64.b64encode(buf.getvalue()).decode("ascii")
        return f"data:image/png;base64,{data}"
    else:
        # Fallback: encode as-is without resizing
        suffix = path.suffix.lower()
        mime = {
            ".jpg":  "image/jpeg",
            ".jpeg": "image/jpeg",
            ".png":  "image/png",
            ".gif":  "image/gif",
            ".webp": "image/webp",
        }.get(suffix, "image/png")
        with open(path, "rb") as f:
            data = base64.b64encode(f.read()).decode("ascii")
        return f"data:{mime};base64,{data}"


def encode_logo_as_jpeg(path: Path) -> str:
    """
    Load logo image, resize to 128x128, flatten to a white background (since it's a JPEG),
    and return as a Base64 JPEG data URI. This keeps the file size around 10KB to prevent
    Gmail from treating it as an email attachment.
    """
    if PILLOW_AVAILABLE:
        # Load and convert to RGBA to process transparency
        img = Image.open(path).convert("RGBA")
        # Resize to 128x128
        img = img.resize((128, 128), Image.LANCZOS)
        # Create a white background canvas of 128x128
        bg = Image.new("RGB", (128, 128), (255, 255, 255))
        # Paste the logo over the white background (handles transparent PNG properly)
        bg.paste(img, (0, 0), img)
        
        buf = io.BytesIO()
        bg.save(buf, format="JPEG", quality=90)
        data = base64.b64encode(buf.getvalue()).decode("ascii")
        return f"data:image/jpeg;base64,{data}"
    else:
        # Fallback if Pillow not available
        with open(path, "rb") as f:
            data = base64.b64encode(f.read()).decode("ascii")
        return f"data:image/png;base64,{data}"


# ── Filename sanitizer ────────────────────────────────────────────────────────
def safe_filename(name: str) -> str:
    """'Monica  Mashale' → 'Monica_Mashale'  (safe for all OS)"""
    name = name.strip()
    # Remove characters not safe for file names
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", name)
    # Collapse whitespace to single underscore
    name = re.sub(r"\s+", "_", name)
    return name


# ── Phone formatter ───────────────────────────────────────────────────────────
def format_phone(value) -> str:
    """Keep phone exactly as stored in Excel. Return empty string if blank."""
    if value is None:
        return ""
    # openpyxl returns integers for numeric cells — convert to string
    return str(value).strip()


# ── HTML signature template ───────────────────────────────────────────────────
def build_signature_html(
    name: str,
    designation: str,
    phone: str,
    email: str,
    logo_b64: str,
    icons_b64: dict,
) -> str:
    """
    Build a complete, standalone HTML email signature.

    Rules followed for maximum email-client compatibility:
    - Table-based layout (no CSS grid/flexbox — Outlook doesn't support them)
    - All CSS is inline (Gmail strips <style> tags)
    - No JavaScript
    - All images embedded as Base64 (Outlook/Gmail block external images)
    - Outlook VML fallbacks included where needed
    - MSO-specific attributes added for Outlook
    """

    # Escape all user-supplied data — security: prevents HTML injection
    ename  = html_module.escape(name.strip())
    edesig = html_module.escape(designation.strip())
    ephone = html_module.escape(phone)
    eemail = html_module.escape(email.strip())

    # ── Social icons cells ────────────────────────────────────────────────────
    social_cells = ""
    for key in SOCIAL_ORDER:
        link  = SOCIAL_LINKS[key]
        label = SOCIAL_LABELS[key]
        img   = icons_b64.get(key, "")
        if not img:
            continue
        social_cells += (
            f'\n                <td style="padding:0 8px 0 0;'
            f'mso-table-lspace:0pt;mso-table-rspace:0pt;">'
            f'<a href="{link}" target="_blank" rel="noopener noreferrer"'
            f' style="text-decoration:none;display:inline-block;border:0;">'
            f'<img src="{img}" width="28" height="28" alt="{label}"'
            f' title="{label}"'
            f' style="display:block;border:0;border-radius:5px;'
            f'width:28px;height:28px;mso-border-alt:none;" />'
            f'</a></td>'
        )

    # ── Full HTML ─────────────────────────────────────────────────────────────
    return f"""<!DOCTYPE html>
<html lang="en"
      xmlns="http://www.w3.org/1999/xhtml"
      xmlns:v="urn:schemas-microsoft-com:vml"
      xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1.0" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <title>Email Signature &ndash; {ename} | Abhinav Group</title>
  <!--[if mso]>
  <xml>
    <o:OfficeDocumentSettings>
      <o:AllowPNG/>
      <o:PixelsPerInch>96</o:PixelsPerInch>
    </o:OfficeDocumentSettings>
  </xml>
  <![endif]-->
</head>
<body style="margin:0;padding:0;background-color:#ffffff;
             -webkit-text-size-adjust:100%;-ms-text-size-adjust:100%;">

<!-- ═══════════════════════════════════════════════════════════
     ABHINAV GROUP EMAIL SIGNATURE
     DO NOT EDIT THIS FILE DIRECTLY.
     Update employees/employees.xlsx and run generate_signatures.py
     ═══════════════════════════════════════════════════════════ -->

<table cellpadding="0" cellspacing="0" border="0" role="presentation"
       style="border-collapse:collapse;
              mso-table-lspace:0pt;mso-table-rspace:0pt;
              font-family:Arial,Helvetica,sans-serif;
              background-color:#ffffff;
              max-width:580px;width:580px;">

  <tr>

    <!-- ══ LOGO CELL ══════════════════════════════════════════ -->
    <td width="145" valign="middle"
        style="width:145px;vertical-align:middle;
               padding:12px 0 12px 0;
               mso-table-lspace:0pt;mso-table-rspace:0pt;">
      <img src="{logo_b64}"
           width="128" height="128"
           alt="Abhinav Group"
           title="Abhinav Group"
           style="display:block;border:0;
                  width:128px;height:128px;
                  mso-border-alt:none;" />
    </td>

    <!-- ══ VERTICAL DIVIDER ═══════════════════════════════════ -->
    <td width="2" valign="top" bgcolor="#1b8ecb"
        style="width:2px;background-color:#1b8ecb;padding:0;
               mso-table-lspace:0pt;mso-table-rspace:0pt;">
      <!--[if mso]>
      <v:rect style="width:2px;height:152px;" fillcolor="#1b8ecb" stroked="f">
        <v:fill type="solid" color="#1b8ecb"/>
      </v:rect>
      <![endif]-->
      <!--[if !mso]><!-->
      <div style="width:2px;min-height:152px;height:100%;
                  background-color:#1b8ecb;
                  font-size:1px;line-height:1px;">&nbsp;</div>
      <!--<![endif]-->
    </td>

    <!-- ══ HORIZONTAL SPACER ══════════════════════════════════ -->
    <td width="18"
        style="width:18px;mso-table-lspace:0pt;mso-table-rspace:0pt;">
      &nbsp;
    </td>

    <!-- ══ CONTENT CELL ═══════════════════════════════════════ -->
    <td valign="top"
        style="vertical-align:top;padding:10px 12px 10px 0;
               mso-table-lspace:0pt;mso-table-rspace:0pt;">

      <!-- NAME ─────────────────────────────────────────────── -->
      <div style="margin:0 0 2px 0;
                  font-size:22px;font-weight:700;
                  color:#1b8ecb;line-height:1.2;
                  font-family:Arial,Helvetica,sans-serif;
                  mso-line-height-rule:exactly;">
        {ename}
      </div>

      <!-- DESIGNATION ──────────────────────────────────────── -->
      <div style="margin:0 0 11px 0;
                  font-size:12.5px;color:#666666;font-weight:400;
                  line-height:1.4;
                  font-family:Arial,Helvetica,sans-serif;
                  mso-line-height-rule:exactly;">
        {edesig}
      </div>

      <!-- PHONE ────────────────────────────────────────────── -->
      <table cellpadding="0" cellspacing="0" border="0" role="presentation"
             style="border-collapse:collapse;margin:0 0 5px 0;
                    mso-table-lspace:0pt;mso-table-rspace:0pt;">
        <tr>
          <td width="22" valign="middle"
              style="width:22px;padding:0 7px 0 0;vertical-align:middle;
                     mso-table-lspace:0pt;mso-table-rspace:0pt;">
            <img src="{PHONE_ICON_URI}" width="15" height="15" alt="Phone"
                 style="display:block;border:0;width:15px;height:15px;" />
          </td>
          <td valign="middle"
              style="vertical-align:middle;
                     mso-table-lspace:0pt;mso-table-rspace:0pt;">
            <span style="font-size:12.5px;color:#333333;
                         font-family:Arial,Helvetica,sans-serif;">
              {ephone}
            </span>
          </td>
        </tr>
      </table>

      <!-- EMAIL ────────────────────────────────────────────── -->
      <table cellpadding="0" cellspacing="0" border="0" role="presentation"
             style="border-collapse:collapse;margin:0 0 5px 0;
                    mso-table-lspace:0pt;mso-table-rspace:0pt;">
        <tr>
          <td width="22" valign="middle"
              style="width:22px;padding:0 7px 0 0;vertical-align:middle;
                     mso-table-lspace:0pt;mso-table-rspace:0pt;">
            <img src="{EMAIL_ICON_URI}" width="15" height="15" alt="Email"
                 style="display:block;border:0;width:15px;height:15px;" />
          </td>
          <td valign="middle"
              style="vertical-align:middle;
                     mso-table-lspace:0pt;mso-table-rspace:0pt;">
            <a href="mailto:{eemail}"
               style="font-size:12.5px;color:#333333;text-decoration:none;
                      font-family:Arial,Helvetica,sans-serif;">
              {eemail}
            </a>
          </td>
        </tr>
      </table>

      <!-- WEBSITE ──────────────────────────────────────────── -->
      <table cellpadding="0" cellspacing="0" border="0" role="presentation"
             style="border-collapse:collapse;margin:0 0 12px 0;
                    mso-table-lspace:0pt;mso-table-rspace:0pt;">
        <tr>
          <td width="22" valign="middle"
              style="width:22px;padding:0 7px 0 0;vertical-align:middle;
                     mso-table-lspace:0pt;mso-table-rspace:0pt;">
            <img src="{GLOBE_ICON_URI}" width="15" height="15" alt="Website"
                 style="display:block;border:0;width:15px;height:15px;" />
          </td>
          <td valign="middle"
              style="vertical-align:middle;
                     mso-table-lspace:0pt;mso-table-rspace:0pt;">
            <a href="https://abhinavgroup.co.in/"
               target="_blank" rel="noopener noreferrer"
               style="font-size:12.5px;color:#333333;text-decoration:none;
                      font-family:Arial,Helvetica,sans-serif;">
              www.abhinavgroup.co.in
            </a>
          </td>
        </tr>
      </table>

      <!-- HORIZONTAL RULE ──────────────────────────────────── -->
      <table cellpadding="0" cellspacing="0" border="0" role="presentation"
             style="border-collapse:collapse;margin:0 0 8px 0;width:100%;
                    mso-table-lspace:0pt;mso-table-rspace:0pt;">
        <tr>
          <td height="1" bgcolor="#e0e0e0"
              style="height:1px;background-color:#e0e0e0;padding:0;
                     font-size:1px;line-height:1px;
                     mso-line-height-rule:exactly;">
            &nbsp;
          </td>
        </tr>
      </table>

      <!-- ADDRESS ──────────────────────────────────────────── -->
      <div style="font-size:12px;color:#555555;line-height:1.5;
                  font-weight:bold;
                  font-family:Arial,Helvetica,sans-serif;
                  margin:0 0 10px 0;
                  mso-line-height-rule:exactly;">
        {ADDRESS}
      </div>

      <!-- SOCIAL ICONS ─────────────────────────────────────── -->
      <table cellpadding="0" cellspacing="0" border="0" role="presentation"
             style="border-collapse:collapse;
                    mso-table-lspace:0pt;mso-table-rspace:0pt;">
        <tr>
          {social_cells}
        </tr>
      </table>

    </td><!-- end CONTENT CELL -->

  </tr>
</table>

</body>
</html>"""


# ── Preview dashboard ─────────────────────────────────────────────────────────
def build_preview_html(employees: list) -> str:
    """Build a browser preview page showing all signatures side-by-side."""
    cards = ""
    for idx, (name, designation, _phone, _email, html_filename) in enumerate(
        employees, start=1
    ):
        safe_n = html_module.escape(name)
        safe_d = html_module.escape(designation)
        cards += f"""
      <div class="card">
        <div class="card-header">
          <div>
            <span class="card-name">{safe_n}</span>
            <span class="card-role">{safe_d}</span>
          </div>
          <span class="card-number">#{idx}</span>
        </div>
        <div class="card-body">
          <iframe src="output/{html_filename}"
                  title="Signature for {safe_n}"
                  scrolling="no"
                  style="width:100%;border:none;display:block;height:185px;"
                  onload="autoResize(this)">
          </iframe>
        </div>
      </div>"""

    now = datetime.now().strftime("%d %B %Y, %I:%M %p")
    count = len(employees)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1.0"/>
  <title>Abhinav Group — Signature Preview</title>
  <style>
    *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

    body {{
      font-family: 'Segoe UI', Arial, sans-serif;
      background: #eef3f8;
      min-height: 100vh;
      padding: 0 0 60px 0;
    }}

    /* ── Top bar ── */
    .topbar {{
      background: linear-gradient(135deg, #1b8ecb 0%, #0d6ea3 100%);
      color: #fff;
      padding: 28px 32px 22px;
      box-shadow: 0 3px 12px rgba(27,142,203,.35);
    }}
    .topbar h1 {{
      font-size: 24px;
      font-weight: 700;
      letter-spacing: -.3px;
    }}
    .topbar p {{
      font-size: 13px;
      opacity: .8;
      margin-top: 5px;
    }}
    .badge {{
      display: inline-block;
      background: rgba(255,255,255,.2);
      border-radius: 20px;
      padding: 3px 12px;
      font-size: 12px;
      margin-top: 8px;
    }}

    /* ── Grid ── */
    .grid {{
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(590px, 1fr));
      gap: 22px;
      max-width: 1380px;
      margin: 30px auto 0;
      padding: 0 28px;
    }}

    /* ── Card ── */
    .card {{
      background: #fff;
      border-radius: 14px;
      box-shadow: 0 2px 16px rgba(0,0,0,.08);
      overflow: hidden;
      transition: box-shadow .2s;
    }}
    .card:hover {{
      box-shadow: 0 6px 28px rgba(27,142,203,.18);
    }}
    .card-header {{
      padding: 13px 18px;
      background: linear-gradient(90deg, #1b8ecb, #0d6ea3);
      display: flex;
      justify-content: space-between;
      align-items: center;
    }}
    .card-name {{
      display: block;
      color: #fff;
      font-size: 14.5px;
      font-weight: 600;
    }}
    .card-role {{
      display: block;
      color: rgba(255,255,255,.75);
      font-size: 11.5px;
      margin-top: 2px;
    }}
    .card-number {{
      color: rgba(255,255,255,.55);
      font-size: 18px;
      font-weight: 700;
    }}
    .card-body {{
      padding: 16px 20px 20px;
      background: #fff;
    }}

    /* ── Footer ── */
    .footer {{
      text-align: center;
      color: #aaa;
      font-size: 11.5px;
      margin-top: 36px;
    }}
  </style>
</head>
<body>

  <div class="topbar">
    <h1>Abhinav Group — Email Signatures</h1>
    <p>All employee signatures preview. To update: edit Excel &rarr; run <strong>run_me.bat</strong></p>
    <span class="badge">{count} signature(s) &nbsp;&bull;&nbsp; Generated {now}</span>
  </div>

  <div class="grid">
    {cards}
  </div>

  <div class="footer">
    <p>Abhinav Group Email Signature Generator &mdash; Version 2.0</p>
  </div>

  <script>
    function autoResize(iframe) {{
      try {{
        var h = iframe.contentWindow.document.body.scrollHeight;
        iframe.style.height = (h + 24) + 'px';
      }} catch (e) {{}}
    }}
  </script>
</body>
</html>"""


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    SEP = "=" * 62
    print(f"\n{SEP}")
    print("   Abhinav Group - Email Signature Generator  v2.0")
    print(SEP)

    # Create output directory if it doesn't exist
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # ── Validate required files ───────────────────────────────────────────────
    errors = []
    if not EXCEL_PATH.exists():
        errors.append(f"Excel file not found: {EXCEL_PATH}")
    if not LOGO_PATH.exists():
        errors.append(f"Logo not found: {LOGO_PATH}")
    if errors:
        print("\n[ERROR] Required files missing:")
        for e in errors:
            print(f"   • {e}")
        sys.exit(1)

    # ── Load logo ─────────────────────────────────────────────────────────────
    print("\n[STEP 1/4]  Loading company logo...")
    logo_b64 = encode_logo_as_jpeg(LOGO_PATH)  # optimized JPEG (128x128) under 10KB to prevent Gmail attachment
    print("            [OK] Logo loaded successfully")

    # ── Load and resize social icons ──────────────────────────────────────────
    print("\n[STEP 2/4]  Loading social media icons (resizing to 28×28px)...")
    icons_b64 = {}
    for key, candidates in ICON_FILENAMES.items():
        found = False
        for filename in candidates:
            path = ICONS_DIR / filename
            if path.exists():
                icons_b64[key] = encode_image_file(path, resize=(28, 28))
                print(f"            [OK] {key:10s} -> {filename}")
                found = True
                break
        if not found:
            print(f"            [!!] {key:10s} -> NOT FOUND (will be skipped)")

    # ── Read Excel ────────────────────────────────────────────────────────────
    print(f"\n[STEP 3/4]  Reading employee data from Excel...")
    print(f"            File: {EXCEL_PATH.name}")

    wb = openpyxl.load_workbook(str(EXCEL_PATH), read_only=True, data_only=True)
    ws = wb.active

    employees_done = []
    success_count  = 0
    skip_count     = 0

    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        # Pad to at least 4 columns
        padded = list(row) + [None, None, None, None]
        name, designation, phone, email = padded[:4]

        # Skip completely blank rows silently
        if not any([name, designation, phone, email]):
            skip_count += 1
            continue

        # Report missing required fields
        missing = []
        if not name:        missing.append("Name")
        if not designation: missing.append("Designation")
        if not email:       missing.append("Email")

        if missing:
            print(f"            [!!] Row {row_idx}: Skipping - missing: {', '.join(missing)}")
            skip_count += 1
            continue

        name_str  = str(name).strip()
        desig_str = str(designation).strip()
        phone_str = format_phone(phone)
        email_str = str(email).strip()

        # Generate HTML
        sig_html  = build_signature_html(
            name        = name_str,
            designation = desig_str,
            phone       = phone_str,
            email       = email_str,
            logo_b64    = logo_b64,
            icons_b64   = icons_b64,
        )

        # Write to output/
        fname    = safe_filename(name_str) + ".html"
        out_path = OUTPUT_DIR / fname
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(sig_html)

        employees_done.append((name_str, desig_str, phone_str, email_str, fname))
        success_count += 1
        print(f"            [OK] [{success_count:02d}] {name_str:<30s} -> output/{fname}")

    wb.close()

    # ── Generate preview dashboard ────────────────────────────────────────────
    print(f"\n[STEP 4/4]  Generating preview dashboard...")
    preview_html = build_preview_html(employees_done)
    preview_path = BASE_DIR / "preview_all.html"
    with open(preview_path, "w", encoding="utf-8") as f:
        f.write(preview_html)
    print(f"            [OK] preview_all.html written")

    # ── Summary ───────────────────────────────────────────────────────────────
    print(f"\n{SEP}")
    print(f"   [SUCCESS] {success_count} signature(s) generated")
    if skip_count:
        print(f"   [SKIPPED] {skip_count} empty / incomplete row(s) skipped")
    print(f"\n   Signature files : output\\")
    print(f"   Preview all     : preview_all.html")
    print(f"   Setup guide     : HOW_TO_SETUP.html")
    print(f"{SEP}\n")

    return success_count


if __name__ == "__main__":
    main()
