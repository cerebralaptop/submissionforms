#!/usr/bin/env python3
"""Parse the Green Star Buildings Excel and generate a complete submission website."""

import json
import html as html_mod
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# ── Parse Excel ──────────────────────────────────────────────────────────────
wb = load_workbook("Green_Star_Buildings_v1.1_Submission_Questions.xlsx")

# Category mapping for each sheet
CATEGORIES = {
    "Responsible": [
        "Industry Development", "Responsible Construction",
        "Verification and Handover", "Responsible Resource Mgmt",
        "Responsible Procurement", "Responsible Structure",
        "Responsible Envelope", "Responsible Systems",
        "Responsible Finishes", "Impacts Disclosure",
    ],
    "Healthy": [
        "Clean Air", "Light Quality", "Acoustic Comfort",
        "Exposure to Toxins", "Amenity and Comfort", "Connection to Nature",
    ],
    "Resilient": [
        "Climate Resilience", "Operations Resilience",
        "Community Resilience", "Heat Resilience", "Grid Resilience",
    ],
    "Positive": [
        "Energy Source", "Energy Use", "Upfront Carbon Reduction",
        "Upfront Carbon Compensation", "Refrigerant Systems Impacts",
        "Low-Emissions Transport", "Design for Circularity", "Water Use",
    ],
    "Places": [
        "Movement and Place", "Enjoyable Places",
        "Contribution to Place", "Culture Heritage Identity",
    ],
    "People": [
        "Inclusive Construction", "First Nations Inclusion",
        "Procurement Workforce Inclusion", "Design for Equity",
    ],
    "Nature": [
        "Impacts to Nature", "Biodiversity Enhancement",
        "Nature Connectivity", "Nature Stewardship", "Waterway Protection",
    ],
    "Leadership": [
        "Market Transformation", "Leadership Challenges",
    ],
}

# Parse all sheets
all_credits = []
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]
    credit_data = {
        "sheet_name": sheet_name,
        "sections": [],   # list of {type, title, questions}
        "questions": [],
    }

    current_section = None
    for row_idx in range(2, ws.max_row + 1):
        a = ws.cell(row=row_idx, column=1).value
        b = ws.cell(row=row_idx, column=2).value
        c = ws.cell(row=row_idx, column=3).value
        d = ws.cell(row=row_idx, column=4).value
        e = ws.cell(row=row_idx, column=5).value
        f = ws.cell(row=row_idx, column=6).value
        h = ws.cell(row=row_idx, column=8).value

        # Check if it's a merged header row
        fill_color = ws.cell(row=row_idx, column=1).fill.start_color
        font = ws.cell(row=row_idx, column=1).font

        if a and not b and not e:
            # This is a header row (credit header, level header, or criteria header)
            text = str(a)
            if font.color and hasattr(font.color, 'rgb') and font.color.rgb:
                rgb = str(font.color.rgb)
                if rgb == "00FFFFFF" and font.size == 12:
                    # Credit header (green bg, white text, size 12)
                    credit_data["title"] = text
                    continue
                elif rgb == "00FFFFFF" and font.size == 11:
                    # Level header
                    current_section = {"type": "level", "title": text, "criteria": []}
                    credit_data["sections"].append(current_section)
                    continue
                elif "1F4E28" in rgb:
                    # Criteria header (light green bg, dark text)
                    if current_section is not None:
                        current_section["criteria"].append({"name": text, "questions": []})
                    continue
            # Fallback detection by font size/bold
            if font.bold and font.size and font.size >= 12:
                credit_data["title"] = text
                continue
            elif font.bold and font.size and font.size >= 11:
                if font.color and hasattr(font.color, 'rgb') and "1F4E28" in str(font.color.rgb):
                    if current_section is not None:
                        current_section["criteria"].append({"name": text, "questions": []})
                    continue
                current_section = {"type": "level", "title": text, "criteria": []}
                credit_data["sections"].append(current_section)
                continue

        # It's a question row if columns E and F have content
        if e and f:
            q = {
                "ref": str(a) if a else "",
                "credit": str(b) if b else "",
                "level": str(c) if c else "",
                "criteria": str(d) if d else "",
                "type": str(e) if e else "",
                "question": str(f) if f else "",
                "data_note": str(h) if h else "",
            }
            credit_data["questions"].append(q)
            # Add to current criteria; create a "General" one if none exists yet
            if current_section:
                if not current_section["criteria"]:
                    current_section["criteria"].append({"name": "General", "questions": []})
                current_section["criteria"][-1]["questions"].append(q)

    all_credits.append(credit_data)

# Map credits to categories
def find_category(sheet_name):
    for cat, sheets in CATEGORIES.items():
        for s in sheets:
            if s.lower().replace(" ", "") in sheet_name.lower().replace(" ", ""):
                return cat
    return "Other"

for c in all_credits:
    c["category"] = find_category(c["sheet_name"])

# ── Generate HTML ────────────────────────────────────────────────────────────
def esc(s):
    return html_mod.escape(str(s)) if s else ""

category_colors = {
    "Responsible": {"bg": "#1F4E28", "light": "#E8F5E9", "mid": "#A5D6A7"},
    "Healthy": {"bg": "#1565C0", "light": "#E3F2FD", "mid": "#90CAF9"},
    "Resilient": {"bg": "#E65100", "light": "#FFF3E0", "mid": "#FFCC80"},
    "Positive": {"bg": "#2E7D32", "light": "#F1F8E9", "mid": "#C5E1A5"},
    "Places": {"bg": "#6A1B9A", "light": "#F3E5F5", "mid": "#CE93D8"},
    "People": {"bg": "#C62828", "light": "#FFEBEE", "mid": "#EF9A9A"},
    "Nature": {"bg": "#00695C", "light": "#E0F2F1", "mid": "#80CBC4"},
    "Leadership": {"bg": "#F57F17", "light": "#FFFDE7", "mid": "#FFF176"},
}

category_icons = {
    "Responsible": "&#9878;",   # recycling
    "Healthy": "&#9829;",      # heart
    "Resilient": "&#9730;",    # umbrella
    "Positive": "&#9889;",     # lightning
    "Places": "&#9962;",       # building
    "People": "&#9823;",       # person
    "Nature": "&#9752;",       # leaf
    "Leadership": "&#9733;",   # star
}

# Build sidebar and pages
sidebar_html = ""
pages_html = ""
credit_index = 0

for cat_name, cat_sheets in CATEGORIES.items():
    colors = category_colors[cat_name]
    icon = category_icons[cat_name]
    cat_credits = [c for c in all_credits if c["category"] == cat_name]

    sidebar_html += f'''
    <div class="sidebar-category">
      <div class="sidebar-category-header" style="background:{colors['bg']}" onclick="toggleCategory(this)">
        <span>{icon} {esc(cat_name)}</span>
        <span class="arrow">&#9662;</span>
      </div>
      <div class="sidebar-category-items">'''

    for credit in cat_credits:
        credit_id = f"credit-{credit_index}"
        q_count = len(credit["questions"])
        sidebar_html += f'''
        <div class="sidebar-item" data-credit="{credit_id}" onclick="showCredit('{credit_id}')">
          <span class="sidebar-item-name">{esc(credit["sheet_name"])}</span>
          <span class="sidebar-badge">{q_count}</span>
        </div>'''

        # Build credit page
        title = credit.get("title", credit["sheet_name"])
        pages_html += f'''
    <div class="credit-page" id="{credit_id}" style="display:none">
      <div class="credit-header" style="background:{colors['bg']}">
        <div class="credit-header-top">
          <span class="credit-category-tag" style="background:{colors['mid']};color:{colors['bg']}">{esc(cat_name)}</span>
          <span class="credit-count">{q_count} questions</span>
        </div>
        <h2>{esc(title)}</h2>
      </div>
      <div class="credit-progress-bar">
        <div class="credit-progress-fill" id="{credit_id}-progress" style="background:{colors['bg']}"></div>
      </div>
      <div class="credit-progress-text">
        <span id="{credit_id}-progress-text">0 of {q_count} answered</span>
      </div>
      <div class="credit-body">'''

        for section in credit["sections"]:
            pages_html += f'''
        <div class="level-header" style="background:{colors['bg']}">{esc(section["title"])}</div>'''

            for crit in section["criteria"]:
                pages_html += f'''
        <div class="criteria-header" style="border-left-color:{colors['bg']};background:{colors['light']}">{esc(crit["name"])}</div>'''

                for q in crit["questions"]:
                    q_id = f"{credit_id}-{q['ref'].replace('.', '-')}"
                    type_class = ""
                    input_html = ""

                    if q["type"] == "Condition (Y/N)":
                        type_class = "q-condition"
                        input_html = f'''
              <div class="response-field">
                <select id="{q_id}" class="yn-select" onchange="updateProgress('{credit_id}', {q_count})" data-credit="{credit_id}">
                  <option value="">-- Select --</option>
                  <option value="Yes">Yes</option>
                  <option value="No">No</option>
                </select>
              </div>'''
                    elif q["type"] == "Data":
                        type_class = "q-data"
                        input_html = f'''
              <div class="response-field">
                <textarea id="{q_id}" class="data-input" rows="3" placeholder="Enter data..." oninput="updateProgress('{credit_id}', {q_count})" data-credit="{credit_id}"></textarea>
              </div>'''
                    else:
                        type_class = "q-descriptive"
                        input_html = f'''
              <div class="response-field">
                <textarea id="{q_id}" class="desc-input" rows="4" placeholder="Describe..." oninput="updateProgress('{credit_id}', {q_count})" data-credit="{credit_id}"></textarea>
              </div>'''

                    data_note_html = ""
                    if q["data_note"]:
                        data_note_html = f'''
              <div class="data-note">
                <span class="data-note-icon">&#128269;</span>
                <span class="data-note-label">Research Note:</span> {esc(q["data_note"])}
              </div>'''

                    type_badge = q["type"]
                    pages_html += f'''
        <div class="question-card {type_class}">
          <div class="question-header">
            <span class="question-ref">{esc(q["ref"])}</span>
            <span class="question-type-badge {type_class}-badge">{esc(type_badge)}</span>
          </div>
          <div class="question-text">{esc(q["question"])}</div>
          {input_html}
          {data_note_html}
        </div>'''

        pages_html += '''
      </div>
    </div>'''
        credit_index += 1

    sidebar_html += '''
      </div>
    </div>'''

# Build dashboard summary
total_questions = sum(len(c["questions"]) for c in all_credits)
total_credits = len(all_credits)

dashboard_cards = ""
for cat_name, cat_sheets in CATEGORIES.items():
    colors = category_colors[cat_name]
    icon = category_icons[cat_name]
    cat_credits = [c for c in all_credits if c["category"] == cat_name]
    cat_q_count = sum(len(c["questions"]) for c in cat_credits)

    dashboard_cards += f'''
      <div class="dash-card" style="border-top:4px solid {colors['bg']}">
        <div class="dash-card-icon" style="color:{colors['bg']}">{icon}</div>
        <div class="dash-card-title">{esc(cat_name)}</div>
        <div class="dash-card-stats">
          <span>{len(cat_credits)} credits</span>
          <span>{cat_q_count} questions</span>
        </div>
        <div class="dash-card-bar">
          <div class="dash-card-bar-fill" id="dash-{cat_name.lower()}-bar" style="background:{colors['bg']}"></div>
        </div>
        <div class="dash-card-pct" id="dash-{cat_name.lower()}-pct">0% complete</div>
      </div>'''

# ── Full HTML ────────────────────────────────────────────────────────────────
html = f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Green Star Buildings v1.1 — Submission Forms</title>
<style>
*, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}

:root {{
  --sidebar-width: 300px;
  --header-height: 60px;
  --bg: #f5f7f5;
  --text: #1a1a1a;
  --text-light: #555;
  --border: #e0e0e0;
  --white: #ffffff;
  --green-dark: #0D3318;
  --green-primary: #1F4E28;
  --green-mid: #2E7D32;
  --green-light: #E8F5E9;
}}

html {{ height: 100%; }}
body {{
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
  background: var(--bg);
  color: var(--text);
  height: 100%;
  overflow: hidden;
}}

/* ── Header ── */
.app-header {{
  position: fixed; top: 0; left: 0; right: 0;
  height: var(--header-height);
  background: var(--green-dark);
  color: white;
  display: flex;
  align-items: center;
  padding: 0 24px;
  z-index: 100;
  box-shadow: 0 2px 8px rgba(0,0,0,0.15);
}}
.app-header h1 {{
  font-size: 18px;
  font-weight: 600;
  letter-spacing: -0.3px;
}}
.app-header .subtitle {{
  margin-left: 12px;
  font-size: 13px;
  opacity: 0.7;
  font-weight: 400;
}}
.header-actions {{
  margin-left: auto;
  display: flex;
  gap: 10px;
}}
.header-btn {{
  background: rgba(255,255,255,0.15);
  color: white;
  border: 1px solid rgba(255,255,255,0.2);
  padding: 6px 14px;
  border-radius: 6px;
  cursor: pointer;
  font-size: 13px;
  transition: background 0.2s;
}}
.header-btn:hover {{ background: rgba(255,255,255,0.25); }}
.header-btn.primary {{ background: var(--green-mid); border-color: var(--green-mid); }}
.header-btn.primary:hover {{ background: #388E3C; }}
.menu-toggle {{
  display: none;
  background: none;
  border: none;
  color: white;
  font-size: 22px;
  cursor: pointer;
  margin-right: 12px;
  padding: 4px;
}}

/* ── Sidebar ── */
.sidebar {{
  position: fixed;
  top: var(--header-height);
  left: 0;
  bottom: 0;
  width: var(--sidebar-width);
  background: var(--white);
  border-right: 1px solid var(--border);
  overflow-y: auto;
  z-index: 50;
  transition: transform 0.3s;
}}
.sidebar-category {{ border-bottom: 1px solid var(--border); }}
.sidebar-category-header {{
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 12px 16px;
  color: white;
  font-weight: 600;
  font-size: 13px;
  cursor: pointer;
  user-select: none;
  text-transform: uppercase;
  letter-spacing: 0.5px;
}}
.sidebar-category-header .arrow {{
  transition: transform 0.2s;
  font-size: 10px;
}}
.sidebar-category-header.collapsed .arrow {{
  transform: rotate(-90deg);
}}
.sidebar-category-items {{
  max-height: 800px;
  overflow: hidden;
  transition: max-height 0.3s ease;
}}
.sidebar-category-header.collapsed + .sidebar-category-items {{
  max-height: 0;
}}
.sidebar-item {{
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 10px 16px 10px 24px;
  cursor: pointer;
  font-size: 13px;
  color: var(--text);
  transition: background 0.15s;
  border-left: 3px solid transparent;
}}
.sidebar-item:hover {{ background: var(--green-light); }}
.sidebar-item.active {{
  background: var(--green-light);
  border-left-color: var(--green-primary);
  font-weight: 600;
}}
.sidebar-item-name {{
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  flex: 1;
}}
.sidebar-badge {{
  background: var(--border);
  color: var(--text-light);
  font-size: 11px;
  padding: 2px 7px;
  border-radius: 10px;
  margin-left: 8px;
  flex-shrink: 0;
}}

/* ── Main Content ── */
.main-content {{
  margin-left: var(--sidebar-width);
  margin-top: var(--header-height);
  height: calc(100vh - var(--header-height));
  overflow-y: auto;
  padding: 0;
}}

/* ── Dashboard ── */
.dashboard {{
  padding: 32px;
}}
.dash-title {{
  font-size: 28px;
  font-weight: 700;
  color: var(--green-dark);
  margin-bottom: 4px;
}}
.dash-subtitle {{
  font-size: 14px;
  color: var(--text-light);
  margin-bottom: 24px;
}}
.dash-overview {{
  display: flex;
  gap: 16px;
  margin-bottom: 32px;
  flex-wrap: wrap;
}}
.dash-stat {{
  background: var(--white);
  border-radius: 10px;
  padding: 20px 28px;
  box-shadow: 0 1px 4px rgba(0,0,0,0.06);
  text-align: center;
  flex: 1;
  min-width: 140px;
}}
.dash-stat-number {{
  font-size: 36px;
  font-weight: 700;
  color: var(--green-primary);
}}
.dash-stat-label {{
  font-size: 12px;
  color: var(--text-light);
  text-transform: uppercase;
  letter-spacing: 0.5px;
  margin-top: 4px;
}}
.dash-grid {{
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(240px, 1fr));
  gap: 16px;
}}
.dash-card {{
  background: var(--white);
  border-radius: 10px;
  padding: 20px;
  box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}}
.dash-card-icon {{ font-size: 28px; margin-bottom: 8px; }}
.dash-card-title {{
  font-size: 16px;
  font-weight: 700;
  margin-bottom: 8px;
}}
.dash-card-stats {{
  display: flex;
  gap: 12px;
  font-size: 12px;
  color: var(--text-light);
  margin-bottom: 12px;
}}
.dash-card-bar {{
  height: 6px;
  background: #eee;
  border-radius: 3px;
  overflow: hidden;
  margin-bottom: 6px;
}}
.dash-card-bar-fill {{
  height: 100%;
  width: 0%;
  border-radius: 3px;
  transition: width 0.4s;
}}
.dash-card-pct {{
  font-size: 12px;
  color: var(--text-light);
}}

/* ── Credit Pages ── */
.credit-page {{ padding: 0; }}
.credit-header {{
  padding: 28px 32px;
  color: white;
}}
.credit-header-top {{
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 10px;
}}
.credit-category-tag {{
  font-size: 11px;
  font-weight: 700;
  padding: 3px 10px;
  border-radius: 4px;
  text-transform: uppercase;
  letter-spacing: 0.5px;
}}
.credit-count {{
  font-size: 13px;
  opacity: 0.8;
}}
.credit-header h2 {{
  font-size: 24px;
  font-weight: 700;
  line-height: 1.3;
}}
.credit-progress-bar {{
  height: 4px;
  background: #eee;
}}
.credit-progress-fill {{
  height: 100%;
  width: 0%;
  transition: width 0.4s;
}}
.credit-progress-text {{
  padding: 8px 32px;
  font-size: 12px;
  color: var(--text-light);
  background: var(--white);
  border-bottom: 1px solid var(--border);
}}
.credit-body {{ padding: 24px 32px 48px; }}

.level-header {{
  color: white;
  padding: 10px 16px;
  border-radius: 6px;
  font-weight: 600;
  font-size: 14px;
  margin: 24px 0 12px 0;
}}
.level-header:first-child {{ margin-top: 0; }}

.criteria-header {{
  padding: 10px 16px;
  border-left: 4px solid;
  border-radius: 0 6px 6px 0;
  font-weight: 600;
  font-size: 13px;
  margin: 16px 0 10px 0;
  color: var(--text);
}}

/* ── Question Cards ── */
.question-card {{
  background: var(--white);
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 16px 20px;
  margin-bottom: 10px;
  transition: box-shadow 0.2s;
}}
.question-card:hover {{ box-shadow: 0 2px 8px rgba(0,0,0,0.08); }}

.question-header {{
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 8px;
}}
.question-ref {{
  font-size: 12px;
  font-weight: 700;
  color: var(--text-light);
  font-family: monospace;
}}
.question-type-badge {{
  font-size: 10px;
  padding: 2px 8px;
  border-radius: 4px;
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.3px;
}}
.q-descriptive-badge {{ background: #F1F8E9; color: #33691E; }}
.q-data-badge {{ background: #E3F2FD; color: #1565C0; }}
.q-condition-badge {{ background: #EDE7F6; color: #4A148C; }}

.question-text {{
  font-size: 14px;
  line-height: 1.5;
  color: var(--text);
  margin-bottom: 12px;
}}

.q-condition {{ border-left: 3px solid #7E57C2; }}
.q-data {{ border-left: 3px solid #42A5F5; }}
.q-descriptive {{ border-left: 3px solid #66BB6A; }}

.response-field {{ margin-bottom: 8px; }}

.yn-select {{
  width: 180px;
  padding: 8px 12px;
  border: 1px solid var(--border);
  border-radius: 6px;
  font-size: 14px;
  background: var(--white);
  cursor: pointer;
  color: var(--text);
}}
.yn-select:focus {{ outline: none; border-color: var(--green-primary); box-shadow: 0 0 0 2px rgba(31,78,40,0.15); }}

.desc-input, .data-input {{
  width: 100%;
  padding: 10px 12px;
  border: 1px solid var(--border);
  border-radius: 6px;
  font-size: 13px;
  font-family: inherit;
  resize: vertical;
  color: var(--text);
  line-height: 1.5;
}}
.desc-input:focus, .data-input:focus {{
  outline: none;
  border-color: var(--green-primary);
  box-shadow: 0 0 0 2px rgba(31,78,40,0.15);
}}
.data-input {{ background: #FAFBFF; }}

.data-note {{
  font-size: 12px;
  color: #666;
  background: #FFFDE7;
  padding: 8px 12px;
  border-radius: 4px;
  line-height: 1.4;
  border: 1px solid #FFF9C4;
}}
.data-note-icon {{ margin-right: 4px; }}
.data-note-label {{ font-weight: 600; color: #F57F17; }}

/* ── Export Modal ── */
.modal-overlay {{
  display: none;
  position: fixed;
  top: 0; left: 0; right: 0; bottom: 0;
  background: rgba(0,0,0,0.4);
  z-index: 200;
  justify-content: center;
  align-items: center;
}}
.modal-overlay.active {{ display: flex; }}
.modal {{
  background: white;
  border-radius: 12px;
  padding: 32px;
  width: 500px;
  max-width: 90%;
  box-shadow: 0 20px 60px rgba(0,0,0,0.2);
}}
.modal h3 {{ font-size: 18px; margin-bottom: 16px; }}
.modal p {{ font-size: 14px; color: var(--text-light); margin-bottom: 16px; line-height: 1.5; }}
.modal-actions {{
  display: flex;
  gap: 10px;
  justify-content: flex-end;
}}
.modal-btn {{
  padding: 8px 20px;
  border-radius: 6px;
  font-size: 14px;
  cursor: pointer;
  border: 1px solid var(--border);
  background: white;
  color: var(--text);
}}
.modal-btn:hover {{ background: #f5f5f5; }}
.modal-btn.primary {{
  background: var(--green-primary);
  color: white;
  border-color: var(--green-primary);
}}
.modal-btn.primary:hover {{ background: var(--green-mid); }}

/* ── Responsive ── */
@media (max-width: 768px) {{
  .menu-toggle {{ display: block; }}
  .sidebar {{
    transform: translateX(-100%);
    width: 280px;
    box-shadow: 4px 0 20px rgba(0,0,0,0.15);
  }}
  .sidebar.open {{ transform: translateX(0); }}
  .main-content {{ margin-left: 0; }}
  .credit-header {{ padding: 20px; }}
  .credit-body {{ padding: 16px; }}
  .dashboard {{ padding: 20px; }}
  .dash-overview {{ flex-direction: column; }}
  .app-header .subtitle {{ display: none; }}
}}

/* ── Scrollbar ── */
::-webkit-scrollbar {{ width: 8px; }}
::-webkit-scrollbar-track {{ background: transparent; }}
::-webkit-scrollbar-thumb {{ background: #ccc; border-radius: 4px; }}
::-webkit-scrollbar-thumb:hover {{ background: #aaa; }}

/* Save notification */
.save-toast {{
  position: fixed;
  bottom: 24px;
  right: 24px;
  background: var(--green-primary);
  color: white;
  padding: 12px 20px;
  border-radius: 8px;
  font-size: 13px;
  box-shadow: 0 4px 12px rgba(0,0,0,0.15);
  transform: translateY(80px);
  opacity: 0;
  transition: all 0.3s;
  z-index: 300;
}}
.save-toast.show {{ transform: translateY(0); opacity: 1; }}
</style>
</head>
<body>

<!-- Header -->
<header class="app-header">
  <button class="menu-toggle" onclick="toggleSidebar()">&#9776;</button>
  <h1>Green Star Buildings v1.1</h1>
  <span class="subtitle">Submission Forms</span>
  <div class="header-actions">
    <button class="header-btn" onclick="showDashboard()">Dashboard</button>
    <button class="header-btn" onclick="saveAllResponses()">Save</button>
    <button class="header-btn primary" onclick="showExportModal()">Export</button>
  </div>
</header>

<!-- Sidebar -->
<nav class="sidebar" id="sidebar">
  <div class="sidebar-item active" data-credit="dashboard" onclick="showDashboard()" style="padding:14px 16px;font-weight:600;border-bottom:1px solid var(--border);">
    <span class="sidebar-item-name">&#9638; Dashboard</span>
  </div>
  {sidebar_html}
</nav>

<!-- Main Content -->
<main class="main-content" id="main-content">

  <!-- Dashboard -->
  <div class="dashboard" id="dashboard">
    <div class="dash-title">Submission Dashboard</div>
    <div class="dash-subtitle">Green Star Buildings v1.1 — Track your progress across all {total_credits} credits</div>

    <div class="dash-overview">
      <div class="dash-stat">
        <div class="dash-stat-number">{total_credits}</div>
        <div class="dash-stat-label">Credits</div>
      </div>
      <div class="dash-stat">
        <div class="dash-stat-number">{total_questions}</div>
        <div class="dash-stat-label">Questions</div>
      </div>
      <div class="dash-stat">
        <div class="dash-stat-number" id="dash-answered">0</div>
        <div class="dash-stat-label">Answered</div>
      </div>
      <div class="dash-stat">
        <div class="dash-stat-number" id="dash-pct">0%</div>
        <div class="dash-stat-label">Complete</div>
      </div>
    </div>

    <div class="dash-grid">
      {dashboard_cards}
    </div>
  </div>

  <!-- Credit Pages -->
  {pages_html}
</main>

<!-- Export Modal -->
<div class="modal-overlay" id="export-modal">
  <div class="modal">
    <h3>Export Responses</h3>
    <p>Download all your responses as a JSON file. You can later import this file to restore your progress.</p>
    <div class="modal-actions">
      <button class="modal-btn" onclick="importResponses()">Import</button>
      <button class="modal-btn primary" onclick="exportResponses()">Export JSON</button>
      <button class="modal-btn" onclick="closeExportModal()">Cancel</button>
    </div>
  </div>
</div>

<!-- Save Toast -->
<div class="save-toast" id="save-toast">Responses saved to browser storage.</div>

<!-- Hidden file input for import -->
<input type="file" id="import-file" accept=".json" style="display:none" onchange="handleImport(event)">

<script>
// ── Navigation ──
function showCredit(id) {{
  document.querySelectorAll('.credit-page').forEach(p => p.style.display = 'none');
  document.getElementById('dashboard').style.display = 'none';
  const page = document.getElementById(id);
  if (page) {{
    page.style.display = 'block';
    document.getElementById('main-content').scrollTop = 0;
  }}
  document.querySelectorAll('.sidebar-item').forEach(i => i.classList.remove('active'));
  const item = document.querySelector(`.sidebar-item[data-credit="${{id}}"]`);
  if (item) item.classList.add('active');
  // Close sidebar on mobile
  document.getElementById('sidebar').classList.remove('open');
}}

function showDashboard() {{
  document.querySelectorAll('.credit-page').forEach(p => p.style.display = 'none');
  document.getElementById('dashboard').style.display = 'block';
  document.getElementById('main-content').scrollTop = 0;
  document.querySelectorAll('.sidebar-item').forEach(i => i.classList.remove('active'));
  document.querySelector('.sidebar-item[data-credit="dashboard"]').classList.add('active');
  updateDashboard();
  document.getElementById('sidebar').classList.remove('open');
}}

function toggleCategory(el) {{
  el.classList.toggle('collapsed');
}}

function toggleSidebar() {{
  document.getElementById('sidebar').classList.toggle('open');
}}

// ── Progress tracking ──
function updateProgress(creditId, total) {{
  const inputs = document.querySelectorAll(`[data-credit="${{creditId}}"]`);
  let answered = 0;
  inputs.forEach(el => {{
    if (el.tagName === 'SELECT' && el.value) answered++;
    if (el.tagName === 'TEXTAREA' && el.value.trim()) answered++;
  }});
  const pct = total > 0 ? (answered / total) * 100 : 0;
  const bar = document.getElementById(`${{creditId}}-progress`);
  const text = document.getElementById(`${{creditId}}-progress-text`);
  if (bar) bar.style.width = pct + '%';
  if (text) text.textContent = `${{answered}} of ${{total}} answered`;

  // Auto-save
  clearTimeout(window._saveTimer);
  window._saveTimer = setTimeout(saveAllResponses, 2000);
}}

function updateDashboard() {{
  const allInputs = document.querySelectorAll('[data-credit]');
  let totalAnswered = 0;
  const categoryStats = {{}};

  allInputs.forEach(el => {{
    const val = el.tagName === 'SELECT' ? el.value : el.value.trim();
    if (val) totalAnswered++;
  }});

  document.getElementById('dash-answered').textContent = totalAnswered;
  document.getElementById('dash-pct').textContent = Math.round((totalAnswered / {total_questions}) * 100) + '%';

  // Per-category
  const catMap = {json.dumps({cat: [c["sheet_name"] for c in all_credits if c["category"] == cat] for cat in CATEGORIES})};
  const creditMap = {json.dumps({c["sheet_name"]: {"id": f"credit-{i}", "count": len(c["questions"])} for i, c in enumerate(all_credits)})};

  for (const [cat, sheets] of Object.entries(catMap)) {{
    let catTotal = 0, catAnswered = 0;
    sheets.forEach(s => {{
      const info = creditMap[s];
      if (!info) return;
      catTotal += info.count;
      document.querySelectorAll(`[data-credit="${{info.id}}"]`).forEach(el => {{
        const val = el.tagName === 'SELECT' ? el.value : el.value.trim();
        if (val) catAnswered++;
      }});
    }});
    const pct = catTotal > 0 ? Math.round((catAnswered / catTotal) * 100) : 0;
    const bar = document.getElementById(`dash-${{cat.toLowerCase()}}-bar`);
    const pctEl = document.getElementById(`dash-${{cat.toLowerCase()}}-pct`);
    if (bar) bar.style.width = pct + '%';
    if (pctEl) pctEl.textContent = pct + '% complete';
  }}
}}

// ── Save / Load ──
function saveAllResponses() {{
  const data = {{}};
  document.querySelectorAll('[data-credit]').forEach(el => {{
    if (el.id) {{
      data[el.id] = el.tagName === 'SELECT' ? el.value : el.value;
    }}
  }});
  try {{
    localStorage.setItem('greenstar_responses', JSON.stringify(data));
    showToast();
  }} catch(e) {{
    console.error('Save failed', e);
  }}
}}

function loadResponses() {{
  try {{
    const raw = localStorage.getItem('greenstar_responses');
    if (!raw) return;
    const data = JSON.parse(raw);
    for (const [id, val] of Object.entries(data)) {{
      const el = document.getElementById(id);
      if (el) el.value = val;
    }}
    // Update all progress bars
    document.querySelectorAll('.credit-page').forEach(page => {{
      const creditId = page.id;
      const inputs = page.querySelectorAll('[data-credit]');
      const total = inputs.length;
      let answered = 0;
      inputs.forEach(el => {{
        const v = el.tagName === 'SELECT' ? el.value : el.value.trim();
        if (v) answered++;
      }});
      const pct = total > 0 ? (answered / total) * 100 : 0;
      const bar = document.getElementById(`${{creditId}}-progress`);
      const text = document.getElementById(`${{creditId}}-progress-text`);
      if (bar) bar.style.width = pct + '%';
      if (text) text.textContent = `${{answered}} of ${{total}} answered`;
    }});
    updateDashboard();
  }} catch(e) {{
    console.error('Load failed', e);
  }}
}}

function showToast() {{
  const t = document.getElementById('save-toast');
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 2000);
}}

// ── Export / Import ──
function showExportModal() {{
  document.getElementById('export-modal').classList.add('active');
}}
function closeExportModal() {{
  document.getElementById('export-modal').classList.remove('active');
}}

function exportResponses() {{
  const data = {{}};
  document.querySelectorAll('[data-credit]').forEach(el => {{
    if (el.id) data[el.id] = el.tagName === 'SELECT' ? el.value : el.value;
  }});
  const blob = new Blob([JSON.stringify(data, null, 2)], {{type: 'application/json'}});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'greenstar_v1.1_responses.json';
  a.click();
  closeExportModal();
}}

function importResponses() {{
  document.getElementById('import-file').click();
}}

function handleImport(event) {{
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {{
    try {{
      const data = JSON.parse(e.target.result);
      for (const [id, val] of Object.entries(data)) {{
        const el = document.getElementById(id);
        if (el) el.value = val;
      }}
      localStorage.setItem('greenstar_responses', JSON.stringify(data));
      loadResponses();
      closeExportModal();
      showToast();
    }} catch(err) {{
      alert('Invalid file format.');
    }}
  }};
  reader.readAsText(file);
  event.target.value = '';
}}

// ── Keyboard shortcut ──
document.addEventListener('keydown', function(e) {{
  if ((e.ctrlKey || e.metaKey) && e.key === 's') {{
    e.preventDefault();
    saveAllResponses();
  }}
}});

// ── Init ──
window.addEventListener('DOMContentLoaded', function() {{
  loadResponses();
  updateDashboard();
}});
</script>
</body>
</html>'''

with open("index.html", "w") as f:
    f.write(html)

print(f"Generated index.html")
print(f"  Credits: {total_credits}")
print(f"  Questions: {total_questions}")
print(f"  Categories: {len(CATEGORIES)}")
print(f"  File size: {len(html):,} bytes")
