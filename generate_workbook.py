#!/usr/bin/env python3
"""
Generate a macro-enabled Excel workbook (.xlsm) from the Green Star Buildings
submission data, with full interactivity via VBA:

  - Dashboard with progress tracking
  - Per-credit sheets with formatted questions, guidance, and input cells
  - Sidebar-style navigation (Dashboard index with hyperlinks)
  - Conditional row visibility (Y/N gateway questions)
  - N/A toggle per credit
  - Review mode (highlight unanswered)
  - Search across all questions
  - Version history on a hidden sheet
  - Dark mode toggle
  - Sheet protection (users can only edit response cells)
"""

import json
import struct
import zlib
import io
import zipfile
import copy
import re
import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, Protection, numbers
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from docx import Document as DocxDocument

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 1: Parse Excel questions
# ═══════════════════════════════════════════════════════════════════════════════
print("Parsing Excel questions...")
wb_src = load_workbook("Green_Star_Buildings_v1.1_Submission_Questions.xlsx")

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

CATEGORY_COLORS = {
    "Responsible": "1F4E28",
    "Healthy": "1565C0",
    "Resilient": "E65100",
    "Positive": "2E7D32",
    "Places": "6A1B9A",
    "People": "C62828",
    "Nature": "00695C",
    "Leadership": "F57F17",
}

all_credits = []
for sheet_name in wb_src.sheetnames:
    ws = wb_src[sheet_name]
    credit = {
        "sheet_name": sheet_name,
        "title": sheet_name,
        "sections": [],
        "questions": [],
    }
    current_section = None
    for row_idx in range(2, ws.max_row + 1):
        a = ws.cell(row=row_idx, column=1).value
        b = ws.cell(row=row_idx, column=2).value
        e = ws.cell(row=row_idx, column=5).value
        f_val = ws.cell(row=row_idx, column=6).value
        h = ws.cell(row=row_idx, column=8).value
        font = ws.cell(row=row_idx, column=1).font

        if a and not b and not e:
            text = str(a)
            if font.color and hasattr(font.color, 'rgb') and font.color.rgb:
                rgb = str(font.color.rgb)
                if rgb == "00FFFFFF" and font.size == 12:
                    credit["title"] = text
                    continue
                elif rgb == "00FFFFFF" and font.size == 11:
                    current_section = {"type": "level", "title": text, "criteria": []}
                    credit["sections"].append(current_section)
                    continue
                elif "1F4E28" in rgb:
                    if current_section is not None:
                        current_section["criteria"].append({"name": text, "questions": []})
                    continue
            if font.bold and font.size and font.size >= 12:
                credit["title"] = text
                continue
            elif font.bold and font.size and font.size >= 11:
                if font.color and hasattr(font.color, 'rgb') and "1F4E28" in str(font.color.rgb):
                    if current_section is not None:
                        current_section["criteria"].append({"name": text, "questions": []})
                    continue
                current_section = {"type": "level", "title": text, "criteria": []}
                credit["sections"].append(current_section)
                continue

        if e and f_val:
            c_val = ws.cell(row=row_idx, column=3).value
            d_val = ws.cell(row=row_idx, column=4).value
            q = {
                "ref": str(a) if a else "",
                "credit": str(b) if b else "",
                "level": str(c_val) if c_val else "",
                "criteria": str(d_val) if d_val else "",
                "type": str(e) if e else "",
                "question": str(f_val) if f_val else "",
                "data_note": str(h) if h else "",
            }
            credit["questions"].append(q)
            if current_section:
                if not current_section["criteria"]:
                    current_section["criteria"].append({"name": "General", "questions": []})
                current_section["criteria"][-1]["questions"].append(q)

    all_credits.append(credit)

def find_category(sheet_name):
    for cat, sheets in CATEGORIES.items():
        for s in sheets:
            if s.lower().replace(" ", "") in sheet_name.lower().replace(" ", ""):
                return cat
    return "Other"

for c in all_credits:
    c["category"] = find_category(c["sheet_name"])

total_credits = len(all_credits)
total_questions = sum(len(c["questions"]) for c in all_credits)
print(f"  {total_credits} credits, {total_questions} questions")

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 2: Parse DOCX guidance
# ═══════════════════════════════════════════════════════════════════════════════
print("Parsing DOCX guidelines...")
docx_doc = DocxDocument("Green Star Buildings v1.1_Submission Guidelines_RevA.docx")

_skip_h1 = {"Version control", "Table of contents", "Introduction",
            "Responsible", "Healthy", "Resilient", "Positive",
            "Places", "People", "Nature", "Leadership"}

docx_guidance = {}
_cc = _h2 = _h3 = _h4 = _h6 = None

for para in docx_doc.paragraphs:
    sn = para.style.name if para.style else ""
    txt = para.text.strip()
    if not txt:
        continue
    if "Heading 1" in sn:
        if txt in _skip_h1 or txt.startswith("Appendix"):
            _cc = None
            continue
        _cc = txt
        _h2 = _h3 = _h4 = _h6 = None
        docx_guidance[_cc] = {"outcome": "", "requirements": {}, "guidance": {}, "evidence": {}, "definitions": []}
    elif _cc and _cc in docx_guidance:
        g = docx_guidance[_cc]
        if "Heading 2" in sn:
            _h2 = txt; _h3 = _h4 = _h6 = None
        elif "Heading 3" in sn:
            _h3 = txt; _h4 = _h6 = None
            if _h2 == "Requirements" and _h3 not in g["requirements"]:
                g["requirements"][_h3] = {}
        elif "Heading 4" in sn:
            _h4 = txt; _h6 = None
            if _h2 == "Requirements" and _h3 and _h3 in g["requirements"]:
                g["requirements"][_h3][_h4] = ""
        elif any(f"Heading {n}" in sn for n in [5, 6, 7]):
            _h6 = txt
            if _h2 == "Guidance":
                g["guidance"][_h6] = ""
            elif _h2 == "Submission content":
                g["evidence"][_h6] = []
        else:
            if _h2 == "Outcome":
                g["outcome"] += txt + " "
            elif _h2 == "Requirements" and _h3 and _h4:
                if _h3 in g["requirements"] and _h4 in g["requirements"][_h3]:
                    g["requirements"][_h3][_h4] += txt + " "
            elif _h2 == "Guidance":
                if _h6 and _h6 in g["guidance"]:
                    g["guidance"][_h6] += txt + " "
                else:
                    g["guidance"]["_general"] = g["guidance"].get("_general", "") + txt + " "
            elif _h2 == "Submission content":
                if _h6 and _h6 in g["evidence"]:
                    g["evidence"][_h6].append(txt)
            elif _h2 == "Definitions":
                g["definitions"].append(txt)

print(f"  {len(docx_guidance)} credits from DOCX")


def _find_docx(sheet_name):
    sn = sheet_name.lower().replace(" ", "")
    for dn, data in docx_guidance.items():
        if sn in dn.lower().replace(" ", "") or dn.lower().replace(" ", "") in sn:
            return data
    return None


def _match_criteria(crit_name, req_dict, guide_dict, ev_dict):
    cn = crit_name.lower().replace(" ", "").replace("-", "").replace("–", "")
    if not cn or cn == "general":
        return None, None, None
    req_match = guide_match = ev_match = None
    for level, crits in req_dict.items():
        for cname, ctext in crits.items():
            if cn in cname.lower().replace(" ", "").replace("-", "") or \
               cname.lower().replace(" ", "").replace("-", "") in cn:
                req_match = (level, cname, ctext.strip())
                break
    for topic, gtext in guide_dict.items():
        if topic == "_general":
            continue
        tn = topic.lower().replace(" ", "").replace("-", "")
        if cn in tn or tn in cn:
            guide_match = gtext.strip()
            break
    for topic, items in ev_dict.items():
        tn = topic.lower().replace(" ", "").replace("-", "")
        if cn in tn or tn in cn:
            ev_match = items
            break
    return req_match, guide_match, ev_match


def get_guidance_text(sheet_name, crit_name, q_type, data_note):
    """Build a plain-text guidance string for a question."""
    g = _find_docx(sheet_name)
    parts = []
    if g:
        if g["outcome"]:
            parts.append(f"OUTCOME: {g['outcome'].strip()[:250]}")
        req_m, guide_m, ev_m = _match_criteria(crit_name or "", g["requirements"], g["guidance"], g["evidence"])
        if req_m:
            parts.append(f"REQUIREMENT ({req_m[0]} — {req_m[1]}): {req_m[2][:350]}")
        if guide_m:
            parts.append(f"WATCH OUT: {guide_m[:350]}")
        if ev_m:
            parts.append("EVIDENCE NEEDED: " + "; ".join(e[:100] for e in ev_m[:4]))
        if g["definitions"]:
            parts.append("DEFINITIONS: " + "; ".join(d[:100] for d in g["definitions"][:2]))
    if data_note:
        parts.append(f"NOTE: {data_note}")
    return "\n\n".join(parts) if parts else data_note or ""


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 3: Build conditional rules (same as website)
# ═══════════════════════════════════════════════════════════════════════════════
print("Building conditional rules...")

conditional_rules = {}  # (credit_idx, follower_ref) -> (credit_idx, gateway_ref, show_when)

for ci, c in enumerate(all_credits):
    sheet = c["sheet_name"]
    qs = c["questions"]
    refs = {q["ref"]: q for q in qs}
    ref_list = [q["ref"] for q in qs]
    ref_idx = {r: i for i, r in enumerate(ref_list)}
    rules = []

    def add_rules(gw, followers, val="Yes"):
        for f in followers:
            if gw in ref_idx and f in ref_idx:
                rules.append((f, gw, val))

    sn = sheet.lower().replace(" ", "")
    if "industrydevelopment" in sn:
        add_rules("ID.5", ["ID.6"], "Yes")
    elif "responsibleconstruction" in sn:
        add_rules("RC.1", ["RC.3"], "Yes")
        add_rules("RC.1", ["RC.2"], "No")
        add_rules("RC.4", ["RC.5"], "Yes")
    elif "verificationandhandover" in sn:
        add_rules("VH.5", ["VH.6"], "Yes")
        add_rules("VH.7", ["VH.8"], "Yes")
        add_rules("VH.26", ["VH.27"], "Yes")
    elif "responsibleresource" in sn:
        add_rules("RRM.2", ["RRM.3"], "Yes")
        add_rules("RRM.5", ["RRM.6"], "Yes")
        add_rules("RRM.7", ["RRM.8"], "Yes")
        add_rules("RRM.12", ["RRM.13"], "Yes")
    elif "responsibleprocurement" in sn:
        add_rules("RP.12", ["RP.13"], "Yes")
    elif "cleanair" in sn:
        add_rules("CA.9", ["CA.10"], "Yes")
    elif "lightquality" in sn:
        add_rules("LQ.7", ["LQ.8"], "Yes")
        add_rules("LQ.7", ["LQ.9", "LQ.10", "LQ.11"], "No")
    elif "exposuretotoxins" in sn:
        add_rules("ET.1", ["ET.2", "ET.3"], "Yes")
    elif "amenityandcomfort" in sn or "amenity" in sn:
        add_rules("AmC.3", ["AmC.4"], "Yes")
    elif "connectiontonature" in sn:
        add_rules("CN.4", ["CN.5", "CN.6"], "Yes")
    elif "climateresilience" in sn:
        add_rules("CR.1", ["CR.2", "CR.3", "CR.4"], "Yes")
    elif "operationsresilience" in sn:
        add_rules("OR.4", ["OR.5"], "Yes")
        add_rules("OR.6", ["OR.7"], "Yes")
    elif "communityresilience" in sn:
        add_rules("CoR.1", ["CoR.2", "CoR.3", "CoR.4"], "Yes")
    elif "gridresilience" in sn:
        add_rules("GR.1", ["GR.2", "GR.3"], "Yes")
        add_rules("GR.4", ["GR.5", "GR.6"], "Yes")
        add_rules("GR.7", ["GR.8"], "Yes")
    elif "energysource" in sn:
        add_rules("ES.5", ["ES.6"], "Yes")
    elif "upfrontcarbonreduction" in sn:
        add_rules("UCR.2", ["UCR.3", "UCR.4", "UCR.5"], "Yes")
    elif "wateruse" in sn:
        add_rules("WU.3", ["WU.4"], "Yes")
        add_rules("WU.5", ["WU.6"], "Yes")
    elif "contributiontoplace" in sn:
        add_rules("CP.1", ["CP.2"], "Yes")
    elif "cultureheritage" in sn:
        add_rules("CHI.1", ["CHI.2"], "Yes")
    elif "firstnations" in sn:
        add_rules("FNI.1", ["FNI.2", "FNI.3"], "Yes")
    elif "designforequity" in sn:
        add_rules("DE.4", ["DE.5"], "Yes")
    elif "impactstonature" in sn:
        add_rules("IN.1", ["IN.2", "IN.3"], "Yes")
        add_rules("IN.7", ["IN.8"], "Yes")
    elif "natureconnectivity" in sn:
        add_rules("NC.1", ["NC.2"], "Yes")
        add_rules("NC.5", ["NC.6"], "Yes")
    elif "naturestewardship" in sn:
        add_rules("NS.1", ["NS.2", "NS.3"], "Yes")
    elif "markettransformation" in sn:
        add_rules("MT.4", ["MT.5"], "Yes")
    elif "waterwayprotection" in sn:
        add_rules("WP.5", ["WP.6"], "Yes")
    elif "impactsdisclosure" in sn:
        add_rules("ID2.5", ["ID2.6"], "Yes")

    for follower_ref, gateway_ref, show_val in rules:
        conditional_rules[(ci, follower_ref)] = (ci, gateway_ref, show_val)

print(f"  {len(conditional_rules)} conditional rules")

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 4: Create the VBA project binary
# ═══════════════════════════════════════════════════════════════════════════════
print("Building VBA project...")


def compress_vba(source_bytes):
    """Compress VBA source using MS-OVBA compression (RLE variant)."""
    compressed = bytearray(b'\x01')  # Signature byte
    src_pos = 0
    src_len = len(source_bytes)

    while src_pos < src_len:
        # Each chunk: up to 4096 bytes of uncompressed data
        chunk_start = src_pos
        chunk_end = min(src_pos + 4096, src_len)
        chunk_data = source_bytes[chunk_start:chunk_end]

        # Compress the chunk
        compressed_chunk = bytearray()
        token_buf = bytearray()
        flags = 0
        flag_count = 0

        i = 0
        while i < len(chunk_data):
            if flag_count == 8:
                compressed_chunk.append(flags)
                compressed_chunk.extend(token_buf)
                flags = 0
                flag_count = 0
                token_buf = bytearray()

            # Try to find a match in the already-processed data
            best_len = 0
            best_off = 0
            decompressed_current = chunk_start + i
            decompressed_chunk_start = chunk_start

            if i > 0:
                # Calculate bit sizes for offset/length encoding
                decompressed_so_far = i
                bit_count = max(4, decompressed_so_far.bit_length())
                max_len_bits = 16 - bit_count
                max_length = (1 << max_len_bits) + 2

                search_start = max(0, i - (1 << bit_count) + 1)
                for j in range(search_start, i):
                    match_len = 0
                    while (i + match_len < len(chunk_data) and
                           match_len < max_length and
                           chunk_data[j + match_len] == chunk_data[i + match_len]):
                        match_len += 1
                        if j + match_len >= i:
                            break
                    if match_len >= 3 and match_len > best_len:
                        best_len = match_len
                        best_off = i - j

            if best_len >= 3:
                # Emit a copy token
                decompressed_so_far = i
                bit_count = max(4, decompressed_so_far.bit_length())
                length_bits = 16 - bit_count

                offset_encoded = best_off - 1
                length_encoded = best_len - 3
                token = (offset_encoded << length_bits) | length_encoded
                token_buf.append(token & 0xFF)
                token_buf.append((token >> 8) & 0xFF)
                flags |= (1 << flag_count)
                i += best_len
            else:
                # Emit a literal byte
                token_buf.append(chunk_data[i])
                i += 1

            flag_count += 1

        # Flush remaining tokens
        if flag_count > 0:
            compressed_chunk.append(flags)
            compressed_chunk.extend(token_buf)

        # Build chunk header
        chunk_size = len(compressed_chunk)
        is_compressed = 1
        if chunk_size >= len(chunk_data):
            # Compression didn't help, store raw
            compressed_chunk = bytearray(chunk_data)
            chunk_size = len(compressed_chunk)
            is_compressed = 0

        # Chunk header: 2 bytes
        # Bits 0-11: chunk size - 3
        # Bit 12-14: signature (0b011)
        # Bit 15: is_compressed
        header = ((chunk_size - 3) & 0x0FFF) | 0x3000
        if is_compressed:
            header |= 0x8000
        compressed.append(header & 0xFF)
        compressed.append((header >> 8) & 0xFF)
        compressed.extend(compressed_chunk)

        src_pos = chunk_end

    return bytes(compressed)


def build_cfb(streams):
    """
    Build a minimal Compound File Binary (CFB / OLE) container.
    streams: list of (name, data_bytes) — paths like "VBA/dir"
    Returns the complete binary content.
    """
    SECTOR_SIZE = 512
    MINI_SECTOR_SIZE = 64
    MINI_STREAM_CUTOFF = 0x1000

    # Build directory tree
    # Root entry is always first
    entries = [{"name": "Root Entry", "type": 5, "data": b"", "children": [], "child_idx": -1}]
    # Group streams by storage
    storages = {}
    for path, data in streams:
        parts = path.split("/")
        if len(parts) == 1:
            entries.append({"name": parts[0], "type": 2, "data": data, "children": [], "child_idx": -1})
            entries[0]["children"].append(len(entries) - 1)
        else:
            storage_name = parts[0]
            stream_name = parts[1]
            if storage_name not in storages:
                storages[storage_name] = len(entries)
                entries.append({"name": storage_name, "type": 1, "data": b"", "children": [], "child_idx": -1})
                entries[0]["children"].append(len(entries) - 1)
            entries.append({"name": stream_name, "type": 2, "data": data, "children": [], "child_idx": -1})
            entries[storages[storage_name]]["children"].append(len(entries) - 1)

    # Set child pointers (simple: first child is the root of a binary tree)
    for entry in entries:
        if entry["children"]:
            entry["child_idx"] = entry["children"][0]

    # Collect all data that goes into regular sectors (>= MINI_STREAM_CUTOFF or storage)
    # and mini-stream data
    mini_stream_data = bytearray()
    for entry in entries:
        if entry["type"] == 2:  # stream
            if len(entry["data"]) < MINI_STREAM_CUTOFF:
                entry["mini_start"] = len(mini_stream_data)
                mini_stream_data.extend(entry["data"])
                # Pad to mini sector boundary
                pad = (MINI_SECTOR_SIZE - len(mini_stream_data) % MINI_SECTOR_SIZE) % MINI_SECTOR_SIZE
                mini_stream_data.extend(b'\x00' * pad)
                entry["start_sector"] = -1  # will be set to mini stream
            else:
                entry["mini_start"] = -1
                entry["start_sector"] = -1  # will assign later
        else:
            entry["mini_start"] = -1
            entry["start_sector"] = -1

    # Build sectors: directory, mini-stream, mini-FAT, then regular stream data
    sectors = []

    def alloc_sector(data_chunk):
        idx = len(sectors)
        sectors.append(bytearray(data_chunk) + b'\x00' * (SECTOR_SIZE - len(data_chunk)))
        return idx

    # Build directory entries (128 bytes each, 4 per sector)
    dir_entries_bin = bytearray()
    for i, entry in enumerate(entries):
        name_utf16 = entry["name"].encode("utf-16-le")
        name_bytes = name_utf16[:62]  # max 31 chars + null = 64 bytes
        name_padded = name_bytes + b'\x00' * (64 - len(name_bytes))
        name_size = len(name_bytes) + 2  # include null terminator

        # Links: for simplicity, use a flat structure
        # Left sibling, right sibling
        left_sib = 0xFFFFFFFF
        right_sib = 0xFFFFFFFF

        # For children of the same parent, chain them as right siblings
        # First child has no left; rest chain right
        parent = None
        for pe in entries:
            if i in pe["children"]:
                parent = pe
                break

        if parent:
            idx_in_parent = parent["children"].index(i)
            if idx_in_parent > 0:
                left_sib = parent["children"][idx_in_parent - 1]
            if idx_in_parent < len(parent["children"]) - 1:
                right_sib = parent["children"][idx_in_parent + 1]

        child_id = entry["child_idx"] if entry["child_idx"] >= 0 else 0xFFFFFFFF

        # Start sector
        if entry["type"] == 5:  # root -> mini stream start
            start_sect = 0xFFFFFFFE  # will fix later
        elif entry["type"] == 2 and entry["mini_start"] >= 0:
            start_sect = entry["mini_start"] // MINI_SECTOR_SIZE
        else:
            start_sect = 0xFFFFFFFE

        data_size = len(entry["data"]) if entry["type"] == 2 else len(mini_stream_data) if entry["type"] == 5 else 0

        # Color: always black (1) for simplicity
        color = 1

        # Build the 128-byte directory entry
        de = bytearray(128)
        de[0:64] = name_padded
        struct.pack_into("<H", de, 64, name_size)
        de[66] = entry["type"]
        de[67] = color
        struct.pack_into("<I", de, 68, left_sib)
        struct.pack_into("<I", de, 72, right_sib)
        struct.pack_into("<I", de, 76, child_id)
        # CLSID (16 bytes at offset 80) - zero
        struct.pack_into("<I", de, 116, start_sect)
        struct.pack_into("<I", de, 120, data_size)

        dir_entries_bin.extend(de)

    # Allocate sectors for mini-stream data (stored as root entry's data)
    mini_stream_sectors = []
    for off in range(0, len(mini_stream_data), SECTOR_SIZE):
        chunk = mini_stream_data[off:off + SECTOR_SIZE]
        idx = alloc_sector(chunk)
        mini_stream_sectors.append(idx)

    # Fix root entry start sector
    if mini_stream_sectors:
        struct.pack_into("<I", dir_entries_bin, 116, mini_stream_sectors[0])

    # Allocate sectors for regular (large) streams
    for entry in entries:
        if entry["type"] == 2 and entry["mini_start"] < 0:
            data = entry["data"]
            stream_sectors = []
            for off in range(0, len(data), SECTOR_SIZE):
                chunk = data[off:off + SECTOR_SIZE]
                idx = alloc_sector(chunk)
                stream_sectors.append(idx)
            if stream_sectors:
                entry["start_sector"] = stream_sectors[0]
                entry["_sectors"] = stream_sectors

    # Allocate directory sectors
    dir_sectors = []
    for off in range(0, len(dir_entries_bin), SECTOR_SIZE):
        chunk = dir_entries_bin[off:off + SECTOR_SIZE]
        idx = alloc_sector(chunk)
        dir_sectors.append(idx)

    # Build mini-FAT
    mini_fat_entries = []
    total_mini_sectors = len(mini_stream_data) // MINI_SECTOR_SIZE

    for entry in entries:
        if entry["type"] == 2 and entry["mini_start"] >= 0:
            start_ms = entry["mini_start"] // MINI_SECTOR_SIZE
            num_ms = (len(entry["data"]) + MINI_SECTOR_SIZE - 1) // MINI_SECTOR_SIZE
            for j in range(num_ms):
                ms_idx = start_ms + j
                while len(mini_fat_entries) <= ms_idx:
                    mini_fat_entries.append(0xFFFFFFFF)
                if j < num_ms - 1:
                    mini_fat_entries[ms_idx] = ms_idx + 1
                else:
                    mini_fat_entries[ms_idx] = 0xFFFFFFFE  # end of chain

    # Pad mini FAT
    while len(mini_fat_entries) < total_mini_sectors:
        mini_fat_entries.append(0xFFFFFFFF)

    mini_fat_bin = b''.join(struct.pack("<I", e) for e in mini_fat_entries)
    mini_fat_sectors = []
    for off in range(0, max(len(mini_fat_bin), 1), SECTOR_SIZE):
        chunk = mini_fat_bin[off:off + SECTOR_SIZE] if off < len(mini_fat_bin) else b''
        idx = alloc_sector(chunk)
        mini_fat_sectors.append(idx)

    # Build FAT
    total_sectors = len(sectors)
    # FAT sectors come after all data sectors
    fat_entries_needed = total_sectors + 1  # +1 for FAT sector itself (at least)
    # How many FAT sectors? Each FAT sector holds 128 entries
    num_fat_sectors = (fat_entries_needed + 127) // 128

    fat = [0xFFFFFFFF] * ((num_fat_sectors) * 128)
    fat_sector_ids = []

    for _ in range(num_fat_sectors):
        idx = len(sectors)
        sectors.append(bytearray(SECTOR_SIZE))
        fat_sector_ids.append(idx)

    # Now fill in the FAT chain entries
    # Mini stream sectors chain
    for i, ms in enumerate(mini_stream_sectors):
        if i < len(mini_stream_sectors) - 1:
            fat[ms] = mini_stream_sectors[i + 1]
        else:
            fat[ms] = 0xFFFFFFFE

    # Regular stream sectors chain
    for entry in entries:
        if hasattr(entry, '_sectors') or (entry["type"] == 2 and entry.get("_sectors")):
            pass
    for entry in entries:
        ss = entry.get("_sectors", [])
        for i, s in enumerate(ss):
            if i < len(ss) - 1:
                fat[s] = ss[i + 1]
            else:
                fat[s] = 0xFFFFFFFE

    # Directory sectors chain
    for i, ds in enumerate(dir_sectors):
        if i < len(dir_sectors) - 1:
            fat[ds] = dir_sectors[i + 1]
        else:
            fat[ds] = 0xFFFFFFFE

    # Mini FAT sectors chain
    for i, mfs in enumerate(mini_fat_sectors):
        if i < len(mini_fat_sectors) - 1:
            fat[mfs] = mini_fat_sectors[i + 1]
        else:
            fat[mfs] = 0xFFFFFFFE

    # FAT sectors are marked as FAT sectors (0xFFFFFFFD)
    for fs in fat_sector_ids:
        fat[fs] = 0xFFFFFFFD

    # Write FAT data into FAT sectors
    fat_bin = b''.join(struct.pack("<I", e) for e in fat[:num_fat_sectors * 128])
    for i, fs in enumerate(fat_sector_ids):
        sectors[fs] = bytearray(fat_bin[i * SECTOR_SIZE:(i + 1) * SECTOR_SIZE])
        if len(sectors[fs]) < SECTOR_SIZE:
            sectors[fs].extend(b'\xff' * (SECTOR_SIZE - len(sectors[fs])))

    # Update directory entries with correct start sectors
    # Re-serialize directory entries with updated start sectors
    dir_entries_bin2 = bytearray()
    for i, entry in enumerate(entries):
        # Copy the old entry
        de = bytearray(dir_entries_bin[i * 128:(i + 1) * 128])

        if entry["type"] == 5:  # root
            if mini_stream_sectors:
                struct.pack_into("<I", de, 116, mini_stream_sectors[0])
            else:
                struct.pack_into("<I", de, 116, 0xFFFFFFFE)
            struct.pack_into("<I", de, 120, len(mini_stream_data))
        elif entry["type"] == 2:
            ss = entry.get("_sectors", [])
            if ss:
                struct.pack_into("<I", de, 116, ss[0])
            elif entry["mini_start"] >= 0:
                struct.pack_into("<I", de, 116, entry["mini_start"] // MINI_SECTOR_SIZE)
            struct.pack_into("<I", de, 120, len(entry["data"]))

        dir_entries_bin2.extend(de)

    # Re-write directory sectors
    for i, ds in enumerate(dir_sectors):
        off = i * SECTOR_SIZE
        chunk = dir_entries_bin2[off:off + SECTOR_SIZE]
        sectors[ds] = bytearray(chunk) + b'\x00' * (SECTOR_SIZE - len(chunk))

    # Build header (512 bytes)
    header = bytearray(SECTOR_SIZE)
    # Magic
    header[0:8] = b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1'
    # Minor version
    struct.pack_into("<H", header, 24, 0x003E)
    # Major version (3 = v3)
    struct.pack_into("<H", header, 26, 0x0003)
    # Byte order (little-endian)
    struct.pack_into("<H", header, 28, 0xFFFE)
    # Sector size power (9 = 512)
    struct.pack_into("<H", header, 30, 9)
    # Mini sector size power (6 = 64)
    struct.pack_into("<H", header, 32, 6)
    # Total sectors in directory (0 for v3)
    struct.pack_into("<I", header, 40, 0)
    # Total FAT sectors
    struct.pack_into("<I", header, 44, num_fat_sectors)
    # First directory sector SECID
    struct.pack_into("<I", header, 48, dir_sectors[0])
    # Mini stream cutoff
    struct.pack_into("<I", header, 56, MINI_STREAM_CUTOFF)
    # First mini FAT sector
    struct.pack_into("<I", header, 60, mini_fat_sectors[0] if mini_fat_sectors else 0xFFFFFFFE)
    # Total mini FAT sectors
    struct.pack_into("<I", header, 64, len(mini_fat_sectors))
    # First DIFAT sector (none needed for small files)
    struct.pack_into("<I", header, 68, 0xFFFFFFFE)
    # Total DIFAT sectors
    struct.pack_into("<I", header, 72, 0)
    # DIFAT array (109 entries starting at offset 76)
    for i in range(109):
        if i < len(fat_sector_ids):
            struct.pack_into("<I", header, 76 + i * 4, fat_sector_ids[i])
        else:
            struct.pack_into("<I", header, 76 + i * 4, 0xFFFFFFFF)

    # Assemble
    result = bytearray(header)
    for sector in sectors:
        result.extend(sector)

    return bytes(result)


def build_dir_stream(modules):
    """Build the VBA 'dir' stream (compressed).
    modules: list of (name, is_document, code_str)
    """
    buf = bytearray()

    def add_record(rec_id, data):
        buf.extend(struct.pack("<HI", rec_id, len(data)))
        buf.extend(data)

    # PROJECTSYSKIND
    add_record(0x0001, struct.pack("<I", 1))  # Win32
    # PROJECTLCID
    add_record(0x0002, struct.pack("<I", 0x0409))
    # PROJECTLCIDINVOKE
    add_record(0x0014, struct.pack("<I", 0x0409))
    # PROJECTCODEPAGE
    add_record(0x0003, struct.pack("<H", 1252))
    # PROJECTNAME
    add_record(0x0004, b"VBAProject")
    # PROJECTDOCSTRING
    add_record(0x0005, b"")
    add_record(0x0040, b"")  # Unicode version
    # PROJECTHELPFILEPATH
    add_record(0x0006, b"")
    add_record(0x003D, b"")
    # PROJECTHELPCONTEXT
    add_record(0x0007, struct.pack("<I", 0))
    # PROJECTLIBFLAGS
    add_record(0x0008, struct.pack("<I", 0))
    # PROJECTVERSION
    buf.extend(struct.pack("<HI", 0x0009, 4))
    buf.extend(struct.pack("<IH", 1, 1))  # major.minor
    # PROJECTCONSTANTS
    add_record(0x000C, b"")
    add_record(0x003C, b"")

    # Reference to stdole
    add_record(0x000D, b"*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\SysWOW64\\stdole2.tlb#OLE Automation")
    add_record(0x000E, b"stdole")

    # Reference to Office
    add_record(0x000D, b"*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16\\MSO.DLL#Microsoft Office 16.0 Object Library")
    add_record(0x000E, b"Office")

    # PROJECTMODULES
    buf.extend(struct.pack("<HI", 0x000F, 2))
    buf.extend(struct.pack("<H", len(modules)))
    # PROJECTCOOKIE
    add_record(0x0013, struct.pack("<H", 0xFFFF))

    for mod_name, is_document, code_str in modules:
        name_bytes = mod_name.encode("ascii")
        # MODULENAME
        add_record(0x0019, name_bytes)
        # MODULENAMEUNICODE
        add_record(0x0047, mod_name.encode("utf-16-le"))
        # MODULESTREAMNAME
        add_record(0x001A, name_bytes)
        add_record(0x0032, mod_name.encode("utf-16-le"))
        # MODULEDOCSTRING
        add_record(0x001C, b"")
        add_record(0x0048, b"")
        # MODULEOFFSET
        add_record(0x0031, struct.pack("<I", 0))
        # MODULEHELPCONTEXT
        add_record(0x001E, struct.pack("<I", 0))
        # MODULECOOKIE
        add_record(0x002C, struct.pack("<H", 0xFFFF))
        # MODULETYPE
        if is_document:
            add_record(0x0022, b"")  # Document module
        else:
            add_record(0x0021, b"")  # Procedural module
        # MODULEREADONLY (not set)
        # MODULEPRIVATE (not set)
        # MODULEEND
        buf.extend(struct.pack("<HI", 0x002B, 0))

    # PROJECTEND
    buf.extend(struct.pack("<HI", 0x0010, 0))

    return compress_vba(buf)


def build_module_stream(source_code):
    """Build a VBA module stream: performance cache (dummy) + compressed source."""
    # The module stream has two parts:
    # 1. Performance cache (we write a minimal one)
    # 2. Compressed source code starting at the offset recorded in dir stream
    # For simplicity, we put offset=0 and just have the compressed source
    return compress_vba(source_code.encode("latin-1", errors="replace"))


def build_vba_project_bin(modules):
    """Build a complete vbaProject.bin OLE file with the given VBA modules.
    modules: list of (name, is_document, source_code_str)
    """
    streams = []

    # _VBA_PROJECT stream (minimal, required)
    vba_project_stream = struct.pack("<HH", 0x61CC, 0x0000) + b'\x00' * 3
    streams.append(("VBA/_VBA_PROJECT", vba_project_stream))

    # dir stream
    dir_data = build_dir_stream(modules)
    streams.append(("VBA/dir", dir_data))

    # Module streams
    for name, is_doc, source in modules:
        mod_data = build_module_stream(source)
        streams.append((f"VBA/{name}", mod_data))

    # PROJECT stream (text)
    project_lines = [
        'ID="{00000000-0000-0000-0000-000000000001}"',
        "Module=GreenStarMacros",
        "Document=ThisWorkbook/&H00000000",
        "Name=\"VBAProject\"",
        "HelpContextID=\"0\"",
        "VersionCompatible32=\"393222000\"",
        "CMG=\"0000\"",
        "DPB=\"0000\"",
        "GC=\"0000\"",
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
        "ThisWorkbook=0, 0, 0, 0, C",
        "GreenStarMacros=0, 0, 0, 0, C",
    ]
    streams.append(("PROJECT", "\r\n".join(project_lines).encode("latin-1")))

    # PROJECTwm stream (module name Unicode map)
    wm_data = bytearray()
    for name, _, _ in modules:
        wm_data.extend(name.encode("ascii") + b'\x00')
        wm_data.extend(name.encode("utf-16-le") + b'\x00\x00')
    wm_data.extend(b'\x00')
    streams.append(("PROJECTwm", bytes(wm_data)))

    return build_cfb(streams)


# ═══════════════════════════════════════════════════════════════════════════════
# STEP 5: Build VBA source code
# ═══════════════════════════════════════════════════════════════════════════════

# Build a lookup table for conditional rules as VBA readable format
# Format: "SheetName|FollowerRef|GatewayRef|ShowWhen"
cond_rules_vba = []
for (ci, fref), (_, gref, sval) in conditional_rules.items():
    sname = all_credits[ci]["sheet_name"][:31]
    cond_rules_vba.append(f"{sname}|{fref}|{gref}|{sval}")

# Build the credit metadata for VBA
credit_meta_vba = []
for ci, c in enumerate(all_credits):
    cat = c["category"]
    color = CATEGORY_COLORS.get(cat, "333333")
    qcount = len(c["questions"])
    credit_meta_vba.append(f"{c['sheet_name'][:31]}|{cat}|{color}|{qcount}")

# Build question row map: SheetName -> list of (ref, row_number, q_type)
# We'll populate row numbers during workbook creation (Step 6)
# For now, store refs and types per credit
question_refs_per_credit = {}
for ci, c in enumerate(all_credits):
    sname = c["sheet_name"][:31]
    question_refs_per_credit[sname] = [(q["ref"], q["type"]) for q in c["questions"]]


# ── VBA ThisWorkbook module ──
THISWORKBOOK_CODE = '''Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    If Target.Column = 7 And Target.Count = 1 Then
        Application.EnableEvents = False
        GreenStarMacros.HandleChange Sh, Target
        Application.EnableEvents = True
    End If
    On Error GoTo 0
End Sub

Private Sub Workbook_Open()
    On Error Resume Next
    GreenStarMacros.InitWorkbook
    On Error GoTo 0
End Sub
'''

# Build VBA rules data as a const string (newline-separated entries)
rules_str = chr(10).join(cond_rules_vba)
meta_str = chr(10).join(credit_meta_vba)

# ── VBA GreenStarMacros module ──
MACROS_CODE = '''Attribute VB_Name = "GreenStarMacros"
Option Explicit

' ============================================================
' GREEN STAR BUILDINGS v1.1 - INTERACTIVE WORKBOOK MACROS
' ============================================================

' ── Data ──
Private Const RULES_DATA As String = "''' + rules_str.replace('"', '""') + '''"
Private Const META_DATA As String = "''' + meta_str.replace('"', '""') + '''"

' ── Colour Scheme ──
Private Const CLR_DARK_BG As Long = &H21201A
Private Const CLR_DARK_CARD As Long = &H342C28
Private Const CLR_DARK_TEXT As Long = &HE0E0E0
Private Const CLR_DARK_INPUT As Long = &H3A3130

Private gDarkMode As Boolean
Private gSearchSheet As String
Private gSearchRow As Long

' ============================================================
' INITIALISATION
' ============================================================
Public Sub InitWorkbook()
    SetupDashboard
End Sub

' ============================================================
' HANDLE CELL CHANGES (Response column = G)
' ============================================================
Public Sub HandleChange(Sh As Object, Target As Range)
    ' Apply conditional visibility rules
    ApplyConditionalRules Sh
    ' Update progress on Dashboard
    UpdateDashboardProgress
    ' Log to history
    LogChange Sh.Name, Target.Row, Target.Value
End Sub

' ============================================================
' CONDITIONAL VISIBILITY
' ============================================================
Public Sub ApplyConditionalRules(Sh As Object)
    Dim rules() As String
    Dim parts() As String
    Dim i As Long
    Dim sName As String

    If Len(RULES_DATA) = 0 Then Exit Sub
    rules = Split(RULES_DATA, vbLf)
    sName = Sh.Name

    For i = LBound(rules) To UBound(rules)
        If Len(rules(i)) = 0 Then GoTo NextRule
        parts = Split(rules(i), "|")
        If UBound(parts) < 3 Then GoTo NextRule
        If parts(0) <> sName Then GoTo NextRule

        Dim followerRef As String, gatewayRef As String, showWhen As String
        followerRef = parts(1)
        gatewayRef = parts(2)
        showWhen = parts(3)

        ' Find gateway row
        Dim gwRow As Long, fRow As Long
        gwRow = FindRefRow(Sh, gatewayRef)
        fRow = FindRefRow(Sh, followerRef)
        If gwRow = 0 Or fRow = 0 Then GoTo NextRule

        Dim gwVal As String
        gwVal = CStr(Sh.Cells(gwRow, 7).Value)

        If gwVal = showWhen Then
            If Sh.Rows(fRow).Hidden Then
                Sh.Rows(fRow).Hidden = False
            End If
        Else
            If Not Sh.Rows(fRow).Hidden Then
                Sh.Rows(fRow).Hidden = True
            End If
        End If
NextRule:
    Next i
End Sub

Private Function FindRefRow(Sh As Object, ref As String) As Long
    Dim lastRow As Long, r As Long
    lastRow = Sh.Cells(Sh.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If CStr(Sh.Cells(r, 1).Value) = ref Then
            FindRefRow = r
            Exit Function
        End If
    Next r
    FindRefRow = 0
End Function

Public Sub ApplyAllConditionalRules()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Dashboard" And ws.Name <> "History" And ws.Name <> "SearchResults" Then
            ApplyConditionalRules ws
        End If
    Next ws
End Sub

' ============================================================
' DASHBOARD PROGRESS
' ============================================================
Public Sub SetupDashboard()
    On Error Resume Next
    Dim dsh As Worksheet
    Set dsh = ThisWorkbook.Worksheets("Dashboard")
    If dsh Is Nothing Then Exit Sub

    UpdateDashboardProgress
    ApplyAllConditionalRules
    On Error GoTo 0
End Sub

Public Sub UpdateDashboardProgress()
    Dim dsh As Worksheet
    Set dsh = ThisWorkbook.Worksheets("Dashboard")
    If dsh Is Nothing Then Exit Sub

    Dim meta() As String
    Dim parts() As String
    Dim totalQ As Long, totalA As Long
    totalQ = 0: totalA = 0

    If Len(META_DATA) = 0 Then Exit Sub
    meta = Split(META_DATA, vbLf)

    Dim dashRow As Long
    dashRow = 5  ' First credit row on dashboard

    Dim i As Long
    For i = LBound(meta) To UBound(meta)
        If Len(meta(i)) = 0 Then GoTo NextMeta
        parts = Split(meta(i), "|")
        If UBound(parts) < 3 Then GoTo NextMeta

        Dim sName As String, qCount As Long
        sName = parts(0)
        qCount = CLng(parts(3))

        ' Count answered in this credit sheet
        Dim ws As Worksheet, answered As Long, visible As Long
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(sName)
        On Error GoTo 0
        If ws Is Nothing Then GoTo NextMeta

        answered = 0: visible = 0
        Dim r As Long
        For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If Len(CStr(ws.Cells(r, 5).Value)) > 0 Then  ' Has question type = is a question row
                If Not ws.Rows(r).Hidden Then
                    visible = visible + 1
                    If Len(CStr(ws.Cells(r, 7).Value)) > 0 Then
                        answered = answered + 1
                    End If
                End If
            End If
        Next r

        totalQ = totalQ + visible
        totalA = totalA + answered

        ' Update dashboard row
        If dashRow <= dsh.Cells(dsh.Rows.Count, 1).End(xlUp).Row + 5 Then
            dsh.Cells(dashRow, 4).Value = answered
            dsh.Cells(dashRow, 5).Value = visible
            If visible > 0 Then
                dsh.Cells(dashRow, 6).Value = answered / visible
            Else
                dsh.Cells(dashRow, 6).Value = 0
            End If
            dashRow = dashRow + 1
        End If
NextMeta:
    Next i

    ' Update totals
    dsh.Cells(2, 4).Value = totalA
    dsh.Cells(2, 5).Value = totalQ
    If totalQ > 0 Then
        dsh.Cells(2, 6).Value = totalA / totalQ
    Else
        dsh.Cells(2, 6).Value = 0
    End If
End Sub

' ============================================================
' N/A TOGGLE
' ============================================================
Public Sub ToggleNA()
    Dim sName As String
    sName = ActiveSheet.Name
    If sName = "Dashboard" Or sName = "History" Or sName = "SearchResults" Then Exit Sub

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Check current state - look at row 2 font color
    If ws.Cells(2, 1).Font.Color = RGB(180, 180, 180) Then
        ' Currently N/A - re-enable
        Dim r2 As Long
        For r2 = 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            ws.Rows(r2).Font.Color = RGB(0, 0, 0)
        Next r2
        ws.Cells(1, 8).Value = ""
    Else
        ' Mark as N/A
        Dim r3 As Long
        For r3 = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            ws.Rows(r3).Font.Color = RGB(180, 180, 180)
        Next r3
        ws.Cells(1, 8).Value = "N/A"
    End If
    UpdateDashboardProgress
End Sub

' ============================================================
' REVIEW MODE - Highlight unanswered
' ============================================================
Public Sub ReviewMode()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws.Name = "Dashboard" Or ws.Name = "History" Or ws.Name = "SearchResults" Then Exit Sub

    Dim r As Long, unanswered As Long
    unanswered = 0

    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Len(CStr(ws.Cells(r, 5).Value)) > 0 Then  ' Question row
            If Not ws.Rows(r).Hidden Then
                If Len(CStr(ws.Cells(r, 7).Value)) = 0 Then
                    ' Highlight unanswered
                    ws.Cells(r, 7).Interior.Color = RGB(255, 243, 224)
                    ws.Cells(r, 7).Borders.Color = RGB(255, 152, 0)
                    unanswered = unanswered + 1
                Else
                    ' Clear highlight
                    ws.Cells(r, 7).Interior.Color = RGB(255, 255, 255)
                    ws.Cells(r, 7).Borders.Color = RGB(200, 200, 200)
                End If
            End If
        End If
    Next r

    MsgBox unanswered & " unanswered question(s) highlighted in orange on " & ws.Name, vbInformation, "Review Mode"
End Sub

' ============================================================
' SEARCH
' ============================================================
Public Sub SearchQuestions()
    Dim query As String
    query = InputBox("Search across all questions:" & vbCrLf & vbCrLf & "Enter search term(s):", "Search Green Star Questions")
    If Len(query) = 0 Then Exit Sub

    Dim searchTerm As String
    searchTerm = LCase(Trim(query))

    ' Create or clear SearchResults sheet
    Dim sr As Worksheet
    On Error Resume Next
    Set sr = ThisWorkbook.Worksheets("SearchResults")
    On Error GoTo 0
    If sr Is Nothing Then
        Set sr = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        sr.Name = "SearchResults"
    End If
    sr.Cells.Clear

    ' Header
    sr.Cells(1, 1).Value = "Search Results for: """ & query & """"
    sr.Cells(1, 1).Font.Bold = True
    sr.Cells(1, 1).Font.Size = 14

    sr.Cells(2, 1).Value = "Credit"
    sr.Cells(2, 2).Value = "Ref"
    sr.Cells(2, 3).Value = "Question"
    sr.Cells(2, 4).Value = "Type"
    sr.Cells(2, 5).Value = "Current Response"
    Dim c As Long
    For c = 1 To 5
        sr.Cells(2, c).Font.Bold = True
        sr.Cells(2, c).Interior.Color = RGB(31, 78, 40)
        sr.Cells(2, c).Font.Color = RGB(255, 255, 255)
    Next c

    sr.Columns(1).ColumnWidth = 25
    sr.Columns(2).ColumnWidth = 8
    sr.Columns(3).ColumnWidth = 60
    sr.Columns(4).ColumnWidth = 16
    sr.Columns(5).ColumnWidth = 40

    Dim resultRow As Long
    resultRow = 3

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Dashboard" Or ws.Name = "History" Or ws.Name = "SearchResults" Then GoTo NextSheet
        Dim r As Long
        For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If Len(CStr(ws.Cells(r, 5).Value)) > 0 Then
                Dim haystack As String
                haystack = LCase(CStr(ws.Cells(r, 1).Value) & " " & CStr(ws.Cells(r, 6).Value) & " " & CStr(ws.Cells(r, 8).Value))
                If InStr(haystack, searchTerm) > 0 Then
                    sr.Cells(resultRow, 1).Value = ws.Name
                    sr.Cells(resultRow, 2).Value = ws.Cells(r, 1).Value
                    sr.Cells(resultRow, 3).Value = ws.Cells(r, 6).Value
                    sr.Cells(resultRow, 4).Value = ws.Cells(r, 5).Value
                    sr.Cells(resultRow, 5).Value = ws.Cells(r, 7).Value
                    ' Add hyperlink to jump to the question
                    sr.Hyperlinks.Add sr.Cells(resultRow, 2), "", "'" & ws.Name & "'!A" & r, "Go to question"
                    resultRow = resultRow + 1
                End If
            End If
        Next r
NextSheet:
    Next ws

    sr.Cells(1, 3).Value = (resultRow - 3) & " result(s) found"

    sr.Activate
End Sub

' ============================================================
' VERSION HISTORY
' ============================================================
Public Sub LogChange(sheetName As String, row As Long, newValue As Variant)
    On Error Resume Next
    Dim hsh As Worksheet
    Set hsh = ThisWorkbook.Worksheets("History")
    If hsh Is Nothing Then Exit Sub

    Dim nextRow As Long
    nextRow = hsh.Cells(hsh.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 3 Then nextRow = 3

    ' Keep max 500 entries
    If nextRow > 502 Then
        hsh.Rows("3:103").Delete
        nextRow = hsh.Cells(hsh.Rows.Count, 1).End(xlUp).Row + 1
    End If

    hsh.Cells(nextRow, 1).Value = Now
    hsh.Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    hsh.Cells(nextRow, 2).Value = sheetName
    hsh.Cells(nextRow, 3).Value = "Row " & row
    hsh.Cells(nextRow, 4).Value = CStr(newValue)
    On Error GoTo 0
End Sub

Public Sub ShowHistory()
    On Error Resume Next
    ThisWorkbook.Worksheets("History").Activate
    On Error GoTo 0
End Sub

' ============================================================
' DARK MODE
' ============================================================
Public Sub ToggleDarkMode()
    gDarkMode = Not gDarkMode
    Dim ws As Worksheet

    If gDarkMode Then
        For Each ws In ThisWorkbook.Worksheets
            ws.Tab.Color = RGB(30, 33, 39)
            Dim lastR As Long, lastC As Long
            lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lastC = 8
            Dim r As Long, cl As Long
            For r = 1 To lastR
                For cl = 1 To lastC
                    If ws.Cells(r, cl).Interior.Color = RGB(255, 255, 255) Or _
                       ws.Cells(r, cl).Interior.ColorIndex = xlNone Then
                        ws.Cells(r, cl).Interior.Color = CLR_DARK_CARD
                        ws.Cells(r, cl).Font.Color = CLR_DARK_TEXT
                    End If
                Next cl
            Next r
        Next ws
        Application.StatusBar = "Dark mode ON"
    Else
        For Each ws In ThisWorkbook.Worksheets
            ws.Tab.Color = xlNone
            lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lastC = 8
            For r = 1 To lastR
                For cl = 1 To lastC
                    If ws.Cells(r, cl).Interior.Color = CLR_DARK_CARD Then
                        ws.Cells(r, cl).Interior.Color = RGB(255, 255, 255)
                        ws.Cells(r, cl).Font.Color = RGB(0, 0, 0)
                    End If
                Next cl
            Next r
        Next ws
        Application.StatusBar = "Dark mode OFF"
    End If
End Sub

' ============================================================
' NAVIGATION HELPERS
' ============================================================
Public Sub GoToDashboard()
    ThisWorkbook.Worksheets("Dashboard").Activate
End Sub

Public Sub RefreshAll()
    ApplyAllConditionalRules
    UpdateDashboardProgress
    MsgBox "All conditional rules applied and dashboard updated.", vbInformation, "Refresh Complete"
End Sub
'''

print("  VBA modules prepared")

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 6: Build the workbook with openpyxl
# ═══════════════════════════════════════════════════════════════════════════════
print("Building workbook...")

wb = Workbook()

# ── Styles ──
header_font = Font(name="Calibri", bold=True, size=14, color="FFFFFF")
credit_font = Font(name="Calibri", bold=True, size=12, color="FFFFFF")
level_font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
criteria_font = Font(name="Calibri", bold=True, size=11, color="1F4E28")
question_font = Font(name="Calibri", size=10)
data_flag_font = Font(name="Calibri", size=10, italic=True, color="2E75B6")
condition_font = Font(name="Calibri", bold=True, size=10, color="7030A0")
guidance_font = Font(name="Calibri", size=9, italic=True, color="666666")
wrap = Alignment(wrap_text=True, vertical="top")
center_wrap = Alignment(wrap_text=True, vertical="center", horizontal="center")

green_fill = PatternFill(start_color="1F4E28", end_color="1F4E28", fill_type="solid")
dark_green_fill = PatternFill(start_color="0D3318", end_color="0D3318", fill_type="solid")
level_fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
criteria_fill = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")
question_fill = PatternFill(start_color="F1F8E9", end_color="F1F8E9", fill_type="solid")
condition_fill = PatternFill(start_color="EDE7F6", end_color="EDE7F6", fill_type="solid")
data_fill = PatternFill(start_color="E3F2FD", end_color="E3F2FD", fill_type="solid")
white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
guidance_fill = PatternFill(start_color="FFFDE7", end_color="FFFDE7", fill_type="solid")

thin_border = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)

COL_WIDTHS = {"A": 8, "B": 20, "C": 22, "D": 28, "E": 16, "F": 55, "G": 50, "H": 45}

yn_dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
yn_dv.error = "Please select Yes or No"
yn_dv.errorTitle = "Invalid entry"
yn_dv.prompt = "Select Yes or No"
yn_dv.promptTitle = "Condition"

# ── Dashboard sheet ──
dsh = wb.active
dsh.title = "Dashboard"
dsh.sheet_properties.tabColor = "1F4E28"

# Dashboard header
dsh.merge_cells("A1:F1")
dsh.cell(row=1, column=1, value="Green Star Buildings v1.1 — Submission Dashboard")
dsh.cell(row=1, column=1).font = Font(name="Calibri", bold=True, size=18, color="FFFFFF")
dsh.cell(row=1, column=1).fill = green_fill
dsh.cell(row=1, column=1).alignment = Alignment(vertical="center")
dsh.row_dimensions[1].height = 45

# Summary row
labels = ["", "", "", "Answered", "Total", "Progress"]
for i, label in enumerate(labels):
    cell = dsh.cell(row=2, column=i + 1, value=label)
    cell.font = Font(name="Calibri", bold=True, size=11)
    cell.alignment = center_wrap

dsh.cell(row=2, column=1, value="TOTAL")
dsh.cell(row=2, column=1).font = Font(name="Calibri", bold=True, size=12, color="1F4E28")
dsh.cell(row=2, column=4, value=0)
dsh.cell(row=2, column=5, value=total_questions)
dsh.cell(row=2, column=6, value=0)
dsh.cell(row=2, column=6).number_format = '0%'
dsh.row_dimensions[2].height = 30

# Blank row
dsh.row_dimensions[3].height = 10

# Column headers for credit list
credit_headers = ["Credit", "Category", "Colour", "Answered", "Total Visible", "Progress"]
for i, h in enumerate(credit_headers):
    cell = dsh.cell(row=4, column=i + 1, value=h)
    cell.font = Font(name="Calibri", bold=True, size=10, color="FFFFFF")
    cell.fill = dark_green_fill
    cell.alignment = center_wrap
    cell.border = thin_border
dsh.row_dimensions[4].height = 25

# Credit rows
for ci, c in enumerate(all_credits):
    row = ci + 5
    sname = c["sheet_name"][:31]
    cat = c["category"]
    color = CATEGORY_COLORS.get(cat, "333333")
    qcount = len(c["questions"])

    # Credit name with hyperlink
    cell = dsh.cell(row=row, column=1, value=sname)
    cell.hyperlink = f"#{sname}!A1"
    cell.font = Font(name="Calibri", size=10, color="1F4E28", underline="single")
    cell.border = thin_border

    dsh.cell(row=row, column=2, value=cat).border = thin_border
    dsh.cell(row=row, column=2).font = Font(name="Calibri", size=10)

    # Category colour indicator
    color_cell = dsh.cell(row=row, column=3)
    color_cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    color_cell.border = thin_border

    # Answered (VBA will update)
    dsh.cell(row=row, column=4, value=0).border = thin_border
    dsh.cell(row=row, column=4).alignment = Alignment(horizontal="center")

    # Total
    dsh.cell(row=row, column=5, value=qcount).border = thin_border
    dsh.cell(row=row, column=5).alignment = Alignment(horizontal="center")

    # Progress
    dsh.cell(row=row, column=6, value=0).border = thin_border
    dsh.cell(row=row, column=6).number_format = '0%'
    dsh.cell(row=row, column=6).alignment = Alignment(horizontal="center")

    dsh.row_dimensions[row].height = 22

# Dashboard column widths
dsh.column_dimensions["A"].width = 30
dsh.column_dimensions["B"].width = 15
dsh.column_dimensions["C"].width = 8
dsh.column_dimensions["D"].width = 12
dsh.column_dimensions["E"].width = 14
dsh.column_dimensions["F"].width = 12

# Toolbar buttons row
btn_row = total_credits + 6
dsh.cell(row=btn_row, column=1, value="Actions:").font = Font(bold=True, size=11)
buttons = [
    ("Review Mode", "ReviewMode"),
    ("Search", "SearchQuestions"),
    ("Dark Mode", "ToggleDarkMode"),
    ("History", "ShowHistory"),
    ("Refresh All", "RefreshAll"),
]
for i, (label, _) in enumerate(buttons):
    cell = dsh.cell(row=btn_row, column=i + 2, value=f"[ {label} ]")
    cell.font = Font(name="Calibri", size=10, color="1F4E28", bold=True)
    cell.alignment = Alignment(horizontal="center")
    cell.border = thin_border

# Add instructions
inst_row = btn_row + 2
instructions = [
    "HOW TO USE:",
    "• Click any credit name above to navigate to its questions",
    "• Fill in the Response column (G) for each question",
    "• Y/N questions use dropdown validation - dependent questions show/hide automatically",
    "• Use the action buttons above: Review highlights gaps, Search finds questions across credits",
    "• The Guidance column (H) shows submission guidelines, tips, and evidence requirements",
    "• To mark a credit as N/A: go to the credit sheet and run the N/A toggle (macros menu)",
    "• Progress updates automatically when you change responses",
    "",
    "KEYBOARD SHORTCUTS (via macros):",
    "• Ctrl+Shift+D = Dashboard  |  Ctrl+Shift+S = Search  |  Ctrl+Shift+R = Review",
]
for i, line in enumerate(instructions):
    cell = dsh.cell(row=inst_row + i, column=1, value=line)
    cell.font = Font(name="Calibri", size=10, color="555555", italic=(i > 0))
    if i == 0:
        cell.font = Font(name="Calibri", size=11, color="1F4E28", bold=True)

dsh.freeze_panes = "A5"

# ── Credit sheets ──
question_row_map = {}  # (sheet_name, ref) -> row number (for conditional rules)

for ci, credit in enumerate(all_credits):
    sname = credit["sheet_name"][:31]
    ws = wb.create_sheet(title=sname)
    cat = credit["category"]
    cat_color = CATEGORY_COLORS.get(cat, "333333")
    ws.sheet_properties.tabColor = cat_color

    # Column widths
    for col_letter, width in COL_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    # Header row
    headers = ["Ref", "Credit", "Performance Level", "Criteria",
               "Question Type", "Question", "Response", "Guidance"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = dark_green_fill
        cell.alignment = center_wrap
        cell.border = thin_border
    ws.row_dimensions[1].height = 35
    ws.freeze_panes = "A2"

    # Add data validation for Y/N
    ws.add_data_validation(yn_dv)

    row = 2

    # Credit title row
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    title = credit.get("title", sname)
    cell = ws.cell(row=row, column=1, value=title)
    cell.font = credit_font
    cell.fill = green_fill
    cell.alignment = wrap
    cell.border = thin_border
    ws.row_dimensions[row].height = 30
    row += 1

    # "Back to Dashboard" link row
    ws.cell(row=row, column=1, value="<< Dashboard")
    ws.cell(row=row, column=1).hyperlink = "#Dashboard!A1"
    ws.cell(row=row, column=1).font = Font(name="Calibri", size=9, color="1F4E28", underline="single")
    ws.row_dimensions[row].height = 18
    row += 1

    for section in credit["sections"]:
        # Level header
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        cell = ws.cell(row=row, column=1, value=section["title"])
        cell.font = level_font
        cell.fill = level_fill
        cell.alignment = wrap
        cell.border = thin_border
        ws.row_dimensions[row].height = 22
        row += 1

        for crit in section["criteria"]:
            # Criteria header
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
            cell = ws.cell(row=row, column=1, value=crit["name"])
            cell.font = criteria_font
            cell.fill = criteria_fill
            cell.alignment = wrap
            cell.border = thin_border
            ws.row_dimensions[row].height = 20
            row += 1

            for q in crit["questions"]:
                is_yn = q["type"] == "Condition (Y/N)"
                is_data = q["type"] == "Data"

                # Build guidance text
                guidance_text = get_guidance_text(
                    credit["sheet_name"], crit["name"], q["type"], q["data_note"]
                )

                values = [q["ref"], q["credit"], q["level"], q["criteria"],
                          q["type"], q["question"], "", guidance_text]

                for col, val in enumerate(values, 1):
                    cell = ws.cell(row=row, column=col, value=val)
                    cell.alignment = wrap
                    cell.border = thin_border

                    if col == 7:
                        # Response column - white, unlocked
                        cell.fill = white_fill
                        cell.font = question_font
                    elif col == 8:
                        # Guidance column
                        cell.fill = guidance_fill
                        cell.font = guidance_font
                    elif is_yn:
                        cell.fill = condition_fill
                        cell.font = condition_font if col == 5 else question_font
                    elif is_data and col == 5:
                        cell.fill = data_fill
                        cell.font = data_flag_font
                    else:
                        cell.fill = question_fill
                        cell.font = question_font

                # Y/N dropdown
                if is_yn:
                    yn_dv.add(ws.cell(row=row, column=7))

                ws.row_dimensions[row].height = 60

                # Track row for conditional rules
                question_row_map[(sname, q["ref"])] = row

                row += 1

    # Protect sheet - only column G (Response) is editable
    for r in range(1, row):
        for c in range(1, 9):
            ws.cell(row=r, column=c).protection = Protection(locked=True)
        # Unlock response column
        if ws.cell(row=r, column=5).value in ["Descriptive", "Data", "Condition (Y/N)"]:
            ws.cell(row=r, column=7).protection = Protection(locked=False)

    ws.protection.sheet = True
    ws.protection.password = "greenstar"
    ws.protection.enable()

# ── History sheet (hidden) ──
hsh = wb.create_sheet(title="History")
hsh.cell(row=1, column=1, value="Version History").font = Font(bold=True, size=14, color="1F4E28")
hsh.cell(row=2, column=1, value="Timestamp").font = Font(bold=True)
hsh.cell(row=2, column=2, value="Sheet").font = Font(bold=True)
hsh.cell(row=2, column=3, value="Location").font = Font(bold=True)
hsh.cell(row=2, column=4, value="New Value").font = Font(bold=True)
hsh.column_dimensions["A"].width = 20
hsh.column_dimensions["B"].width = 25
hsh.column_dimensions["C"].width = 12
hsh.column_dimensions["D"].width = 50
for c in range(1, 5):
    hsh.cell(row=2, column=c).fill = dark_green_fill
    hsh.cell(row=2, column=c).font = Font(bold=True, color="FFFFFF")
    hsh.cell(row=2, column=c).border = thin_border
hsh.freeze_panes = "A3"
hsh.sheet_state = "hidden"

# ── Dashboard protection ──
dsh.protection.sheet = True
dsh.protection.password = "greenstar"
dsh.protection.enable()

print(f"  {total_credits} credit sheets created")
print(f"  {total_questions} questions with guidance")

# ═══════════════════════════════════════════════════════════════════════════════
# STEP 7: Save as .xlsx first, then inject VBA to make .xlsm
# ═══════════════════════════════════════════════════════════════════════════════
print("Saving workbook...")

# Save as xlsx first
xlsx_path = "Green_Star_Buildings_v1.1_Interactive.xlsx"
wb.save(xlsx_path)
print(f"  Saved {xlsx_path}")

# Now inject VBA to create .xlsm
print("Injecting VBA macros...")

vba_modules = [
    ("ThisWorkbook", True, THISWORKBOOK_CODE),
    ("GreenStarMacros", False, MACROS_CODE),
]

try:
    vba_bin = build_vba_project_bin(vba_modules)
    print(f"  vbaProject.bin: {len(vba_bin)} bytes")

    xlsm_path = "Green_Star_Buildings_v1.1_Interactive.xlsm"

    # Read the xlsx as a zip, modify content types and add vba, write as xlsm
    with zipfile.ZipFile(xlsx_path, 'r') as zin:
        with zipfile.ZipFile(xlsm_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                if item.filename == '[Content_Types].xml':
                    # Add VBA content type
                    ct_xml = data.decode('utf-8')
                    # Add macro content type
                    if 'vbaProject.bin' not in ct_xml:
                        ct_xml = ct_xml.replace(
                            '</Types>',
                            '<Override PartName="/xl/vbaProject.bin" '
                            'ContentType="application/vnd.ms-office.vbaProject"/>\n</Types>'
                        )
                    # Change workbook content type from xlsx to xlsm
                    ct_xml = ct_xml.replace(
                        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
                        'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
                    )
                    zout.writestr(item, ct_xml)

                elif item.filename == 'xl/_rels/workbook.xml.rels':
                    # Add relationship to vbaProject.bin
                    rels_xml = data.decode('utf-8')
                    if 'vbaProject.bin' not in rels_xml:
                        rels_xml = rels_xml.replace(
                            '</Relationships>',
                            '<Relationship Id="rIdVBA" Type='
                            '"http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
                            'Target="vbaProject.bin"/>\n</Relationships>'
                        )
                    zout.writestr(item, rels_xml)
                else:
                    zout.writestr(item, data)

            # Add vbaProject.bin
            zout.writestr('xl/vbaProject.bin', vba_bin)

    print(f"  Saved {xlsm_path}")

except Exception as e:
    print(f"  WARNING: VBA injection failed: {e}")
    print(f"  The .xlsx file is still fully functional (without macros)")
    xlsm_path = None

# ═══════════════════════════════════════════════════════════════════════════════
# DONE
# ═══════════════════════════════════════════════════════════════════════════════
import os
print(f"\nDone!")
print(f"  Credits: {total_credits}")
print(f"  Questions: {total_questions}")
print(f"  Conditional rules: {len(conditional_rules)}")
if xlsm_path and os.path.exists(xlsm_path):
    print(f"  Output: {xlsm_path} ({os.path.getsize(xlsm_path):,} bytes)")
print(f"  Backup: {xlsx_path} ({os.path.getsize(xlsx_path):,} bytes)")
print(f"\nTo use: Open the .xlsm file and enable macros.")
print(f"Protection password: greenstar")
