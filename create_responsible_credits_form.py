import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation

wb = openpyxl.Workbook()

# ── Styles ──────────────────────────────────────────────────────────────────
header_font = Font(name="Calibri", bold=True, size=14, color="FFFFFF")
credit_font = Font(name="Calibri", bold=True, size=12, color="FFFFFF")
level_font = Font(name="Calibri", bold=True, size=11, color="FFFFFF")
criteria_font = Font(name="Calibri", bold=True, size=11, color="1F4E28")
question_font = Font(name="Calibri", size=10)
data_flag_font = Font(name="Calibri", size=10, italic=True, color="2E75B6")
condition_font = Font(name="Calibri", bold=True, size=10, color="7030A0")
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

thin_border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

COL_WIDTHS = {"A": 8, "B": 20, "C": 22, "D": 28, "E": 16, "F": 55, "G": 50, "H": 40}

# Reusable Yes/No dropdown
yn_dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
yn_dv.error = "Please select Yes or No"
yn_dv.errorTitle = "Invalid entry"
yn_dv.prompt = "Select Yes or No"
yn_dv.promptTitle = "Yes / No"

# Track which rows need the dropdown per sheet
yn_cells = {}  # sheet title -> list of cell refs


def setup_sheet(ws, title):
    ws.title = title
    for col_letter, width in COL_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    headers = [
        "Ref", "Credit", "Performance Level", "Criteria",
        "Question Type", "Question", "Response",
        "Data Collection / Research Notes",
    ]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = dark_green_fill
        cell.alignment = center_wrap
        cell.border = thin_border
    ws.row_dimensions[1].height = 35
    ws.freeze_panes = "A2"
    yn_cells[title] = []
    return 2


def add_credit_header(ws, row, credit_name, outcome):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    cell = ws.cell(row=row, column=1, value=f"{credit_name} — {outcome}")
    cell.font = credit_font
    cell.fill = green_fill
    cell.alignment = wrap
    cell.border = thin_border
    ws.row_dimensions[row].height = 30
    return row + 1


def add_level_header(ws, row, level_name):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    cell = ws.cell(row=row, column=1, value=level_name)
    cell.font = level_font
    cell.fill = level_fill
    cell.alignment = wrap
    cell.border = thin_border
    ws.row_dimensions[row].height = 22
    return row + 1


def add_criteria_header(ws, row, criteria_name):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    cell = ws.cell(row=row, column=1, value=criteria_name)
    cell.font = criteria_font
    cell.fill = criteria_fill
    cell.alignment = wrap
    cell.border = thin_border
    ws.row_dimensions[row].height = 20
    return row + 1


def add_question(ws, row, ref, credit, level, criteria, q_type, question, data_note=""):
    is_yn = q_type == "Condition (Y/N)"
    values = [ref, credit, level, criteria, q_type, question, "", data_note]
    for col, val in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.alignment = wrap
        cell.border = thin_border
        if is_yn:
            cell.fill = condition_fill
            cell.font = condition_font if col == 5 else question_font
        elif col == 8 and data_note:
            cell.fill = data_fill
            cell.font = data_flag_font
        else:
            cell.fill = question_fill if col != 7 else white_fill
            cell.font = question_font
    ws.row_dimensions[row].height = 45
    if is_yn:
        yn_cells[ws.title].append(f"G{row}")
    return row + 1


def apply_dropdowns(ws):
    """Add a single Yes/No data validation to the sheet covering all Y/N cells."""
    cells = yn_cells.get(ws.title, [])
    if not cells:
        return
    dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    dv.error = "Please select Yes or No"
    dv.errorTitle = "Invalid entry"
    dv.prompt = "Select Yes or No"
    dv.promptTitle = "Yes / No"
    for c in cells:
        dv.add(c)
    ws.add_data_validation(dv)


# ════════════════════════════════════════════════════════════════════════════
# CREDIT 1: INDUSTRY DEVELOPMENT
# ════════════════════════════════════════════════════════════════════════════
ws = wb.active
row = setup_sheet(ws, "Industry Development")
row = add_credit_header(ws, row, "Industry Development",
    "The development facilitates industry transformation through partnership, collaboration and data sharing.")

row = add_level_header(ws, row, "Credit Achievement (1 point)")

# ── Green Star Accredited Professional ──
row = add_criteria_header(ws, row, "Green Star Accredited Professional")
ref_base = "ID"
q = 1

row = add_question(ws, row, f"{ref_base}.{q}", "Industry Development", "Credit Achievement",
    "Green Star Accredited Professional", "Descriptive",
    "Identify the GSAP(s) engaged, including name, organisation, accreditation number, and Green Star Buildings accreditation held.",
    "GSAP workforce capacity and distribution across projects.")
q += 1

row = add_question(ws, row, f"{ref_base}.{q}", "Industry Development", "Credit Achievement",
    "Green Star Accredited Professional", "Descriptive",
    "State the date and project phase when the GSAP was first engaged.",
    "Timing of sustainability expertise integration relative to design stage.")
q += 1

row = add_question(ws, row, f"{ref_base}.{q}", "Industry Development", "Credit Achievement",
    "Green Star Accredited Professional", "Condition (Y/N)",
    "Was the GSAP engaged within one month of project registration?",
    "")
q += 1

row = add_question(ws, row, f"{ref_base}.{q}", "Industry Development", "Credit Achievement",
    "Green Star Accredited Professional", "Descriptive",
    "Summarise the GSAP's scope of advisory and coordination activities on Green Star strategy, process and certification.",
    "Depth of sustainability advisory services on projects.")
q += 1

row = add_question(ws, row, f"{ref_base}.{q}", "Industry Development", "Credit Achievement",
    "Green Star Accredited Professional", "Condition (Y/N)",
    "Was the GSAP role fulfilled by more than one individual or organisation?",
    "Continuity of sustainability expertise across project lifecycle.")
q += 1

row = add_question(ws, row, f"{ref_base}.{q}", "Industry Development", "Credit Achievement",
    "Green Star Accredited Professional", "Descriptive",
    "If multiple GSAPs, explain transitions and confirm each held valid Green Star Buildings accreditation throughout their engagement.",
    "")
q += 1

row = add_question(ws, row, f"{ref_base}.{q}", "Industry Development", "Credit Achievement",
    "Green Star Accredited Professional", "Condition (Y/N)",
    "Is the GSAP nominated as the Project Contact for GBCA communications?",
    "")
q += 1

row = add_question(ws, row, f"{ref_base}.{q}", "Industry Development", "Credit Achievement",
    "Green Star Accredited Professional", "Descriptive",
    "Describe how ongoing GSAP involvement was maintained throughout the project (e.g. design meetings, workshops).",
    "")
q += 1

# ── Financial Transparency ──
row = add_criteria_header(ws, row, "Financial Transparency")

row = add_question(ws, row, f"{ref_base}.{q}", "Industry Development", "Credit Achievement",
    "Financial Transparency", "Condition (Y/N)",
    "Was the Financial Transparency template completed in its latest version and submitted in Excel format?",
    "Industry-wide benchmarking of sustainable building costs.")
q += 1

row = add_question(ws, row, f"{ref_base}.{q}", "Industry Development", "Credit Achievement",
    "Financial Transparency", "Descriptive",
    "Identify who prepared the cost data (e.g. quantity surveyor, head contractor, cost consultant).",
    "")
q += 1

row = add_question(ws, row, f"{ref_base}.{q}", "Industry Development", "Credit Achievement",
    "Financial Transparency", "Descriptive",
    "Explain how documentation and implementation costs for sustainable practices were isolated from the base (non-Green Star) requirement.",
    "Cost premiums/savings of green building practices.")
q += 1

row = add_question(ws, row, f"{ref_base}.{q}", "Industry Development", "Credit Achievement",
    "Financial Transparency", "Data",
    "Provide total project construction cost and total additional cost for sustainable practices (documentation + implementation).",
    "Cost-benefit analysis of green certification across the industry.")
q += 1

# ── Marketing Sustainability Achievements ──
row = add_criteria_header(ws, row, "Marketing Sustainability Achievements")

row = add_question(ws, row, f"{ref_base}.{q}", "Industry Development", "Credit Achievement",
    "Marketing Sustainability Achievements", "Descriptive",
    "List which three or more marketing activities were undertaken: (a) case study to GBCA, (b) digital screens, (c) construction hoarding, (d) marketing/communications strategy.",
    "Industry adoption of sustainability marketing practices.")
q += 1

row = add_question(ws, row, f"{ref_base}.{q}", "Industry Development", "Credit Achievement",
    "Marketing Sustainability Achievements", "Descriptive",
    "Describe how sustainability achievements are communicated to building users, the public, or prospective tenants/buyers.",
    "Effectiveness and reach of green building awareness campaigns.")
q += 1

row = add_question(ws, row, f"{ref_base}.{q}", "Industry Development", "Credit Achievement",
    "Marketing Sustainability Achievements", "Data",
    "Identify the target audience and estimated reach for each marketing activity.",
    "Public awareness exposure to green building benefits.")
q += 1

apply_dropdowns(ws)


# ════════════════════════════════════════════════════════════════════════════
# CREDIT 2: RESPONSIBLE CONSTRUCTION
# ════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet()
row = setup_sheet(ws2, "Responsible Construction")
row = add_credit_header(ws2, row, "Responsible Construction",
    "The builder's construction practices reduce impacts and promote opportunities for improved environmental and social outcomes.")
ref_base = "RC"
q = 1

row = add_level_header(ws2, row, "Minimum Expectation (Nil points)")

# ── Environmental Management System ──
row = add_criteria_header(ws2, row, "Environmental Management System")

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management System", "Condition (Y/N)",
    "Is any site works contract valued at $10 million or more?",
    "Contract sizes relative to EMS certification thresholds.")
q += 1

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management System", "Descriptive",
    "For contracts under $10M, identify the EMS framework used and explain how it complies (e.g. NSW EMS Guidelines or equivalent).",
    "EMS framework adoption rates in construction.")
q += 1

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management System", "Descriptive",
    "For contracts $10M+, state the certified standard (ISO 14001, BS 7750, or EMAS) and confirm certification validity for the full duration of site works.",
    "Uptake of certified environmental management in construction.")
q += 1

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management System", "Condition (Y/N)",
    "Were different head contractors used for demolition, early works, and main works?",
    "")
q += 1

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management System", "Descriptive",
    "If multiple head contractors, confirm each had an EMS in place and explain how contract values were apportioned.",
    "")
q += 1

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management System", "Descriptive",
    "Explain how the EMS addresses implementation of the EMP and the key environmental impacts targeted.",
    "Relationship between management systems and on-site environmental outcomes.")
q += 1

# ── Environmental Management Plan ──
row = add_criteria_header(ws2, row, "Environmental Management Plan")

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management Plan", "Descriptive",
    "Outline the project-specific EMP, including key impact areas addressed (e.g. noise, dust, stormwater, vegetation).",
    "Most common environmental risks managed during construction.")
q += 1

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management Plan", "Condition (Y/N)",
    "Did the EMP cover the full duration of all site works?",
    "Construction duration and environmental management coverage.")
q += 1

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management Plan", "Data",
    "State the EMP start and end dates.",
    "")
q += 1

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management Plan", "Descriptive",
    "Describe the audit and reporting regime, including frequency and how non-conformances were closed out.",
    "Environmental compliance enforcement during construction.")
q += 1

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management Plan", "Data",
    "Provide the total number of audits, non-conformances identified, and percentage closed out.",
    "Quantitative construction environmental management performance.")
q += 1

# ── Construction and Demolition Waste Diversion ──
row = add_criteria_header(ws2, row, "Construction and Demolition Waste Diversion")

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Minimum Expectation",
    "Construction and Demolition Waste Diversion", "Data",
    "State total site waste (tonnes), total diverted from landfill (tonnes), and diversion rate (%).",
    "Construction waste diversion rate benchmarking.")
q += 1

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Minimum Expectation",
    "Construction and Demolition Waste Diversion", "Condition (Y/N)",
    "Does the diversion rate meet the 80% threshold?",
    "")
q += 1

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Minimum Expectation",
    "Construction and Demolition Waste Diversion", "Descriptive",
    "List waste streams and diversion pathways (recycling, reuse, recovery). Note any excluded streams (special/excavation waste) with justification.",
    "Waste stream composition and recycling pathways in construction.")
q += 1

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Minimum Expectation",
    "Construction and Demolition Waste Diversion", "Descriptive",
    "Confirm waste contractors provided a Disclosure Statement aligned with the Green Star C&D Waste Reporting Criteria.",
    "Supply chain transparency in waste reporting.")
q += 1

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Minimum Expectation",
    "Construction and Demolition Waste Diversion", "Data",
    "Provide a breakdown by material type (e.g. concrete, timber, steel, plasterboard) showing tonnes generated and diverted.",
    "Material-specific waste benchmarking across projects.")
q += 1

# ── Sustainability Training ──
row = add_criteria_header(ws2, row, "Sustainability Training")

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Minimum Expectation",
    "Sustainability Training", "Data",
    "State total site workers on site 3+ days, number trained, and resulting percentage.",
    "Sustainability education reach in the construction workforce.")
q += 1

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Minimum Expectation",
    "Sustainability Training", "Condition (Y/N)",
    "Does the training rate meet the 95% threshold?",
    "")
q += 1

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Minimum Expectation",
    "Sustainability Training", "Descriptive",
    "Summarise training content covering: (a) project sustainability attributes, (b) value of certification, (c) site workers' role in delivery.",
    "Training content quality for workforce sustainability literacy.")
q += 1

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Minimum Expectation",
    "Sustainability Training", "Descriptive",
    "Describe the delivery method (e.g. induction, toolbox talks) and how attendance was tracked.",
    "Effective training delivery models for sustainability in construction.")
q += 1

# ── Credit Achievement ──
row = add_level_header(ws2, row, "Credit Achievement (1 point)")
row = add_criteria_header(ws2, row, "Increased Construction and Demolition Waste Diversion")

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Credit Achievement",
    "Increased C&D Waste Diversion", "Data",
    "State total site waste (tonnes), total diverted (tonnes), and diversion rate (%).",
    "Higher-tier waste diversion benchmarking.")
q += 1

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Credit Achievement",
    "Increased C&D Waste Diversion", "Condition (Y/N)",
    "Does the diversion rate meet the 90% threshold?",
    "")
q += 1

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Credit Achievement",
    "Increased C&D Waste Diversion", "Descriptive",
    "Confirm waste contractors/facilities provided a Compliance Verification Summary per the Green Star C&D Waste Reporting Criteria.",
    "Third-party waste reporting verification practices.")
q += 1

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Credit Achievement",
    "Increased C&D Waste Diversion", "Descriptive",
    "Identify the waste reporting auditor(s) and their credentials per the Green Star C&D Waste Reporting Criteria.",
    "Auditor capacity and verification standards in waste management.")
q += 1

row = add_question(ws2, row, f"{ref_base}.{q}", "Responsible Construction", "Credit Achievement",
    "Increased C&D Waste Diversion", "Data",
    "List waste processing facilities used, their location, waste types processed, and any GECA C&D Waste Services Standard certification held.",
    "Waste processing infrastructure availability and certification uptake.")
q += 1

apply_dropdowns(ws2)


# ════════════════════════════════════════════════════════════════════════════
# CREDIT 3: VERIFICATION AND HANDOVER
# ════════════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet()
row = setup_sheet(ws3, "Verification and Handover")
row = add_credit_header(ws3, row, "Verification and Handover",
    "The building has been optimised and handed over to deliver a higher level of performance in operation.")
ref_base = "VH"
q = 1

row = add_level_header(ws3, row, "Minimum Expectation (Nil points)")

# ── Metering and Monitoring ──
row = add_criteria_header(ws3, row, "Metering and Monitoring")

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Metering and Monitoring", "Descriptive",
    "Outline the metering strategy for energy and water across all distinct uses, major uses, and tenancies/units, referencing the CIBSE TM39 schedule.",
    "Metering granularity across building types.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Metering and Monitoring", "Data",
    "State the total number of energy and water meters (utility + sub-meters) and distinct end-uses metered.",
    "Metering density benchmarking across building typologies.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Metering and Monitoring", "Descriptive",
    "Confirm all meters provide up to 1-hour interval readings, are validated per NABERS Metering Rules, and are NMI pattern-approved or meet an equivalent standard.",
    "Metering quality standards adoption.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Metering and Monitoring", "Descriptive",
    "Describe the automatic monitoring system, including consumption trend reporting and alarm/alert functionality for the facilities manager.",
    "Monitoring system capabilities for operational performance.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Metering and Monitoring", "Condition (Y/N)",
    "Is this a Class 2 build-to-sell apartment project?",
    "")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Metering and Monitoring", "Descriptive",
    "If Class 2 build-to-sell, confirm base building trends are provided to the FM and explain how unit meters are handled.",
    "")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Metering and Monitoring", "Condition (Y/N)",
    "Does the metering strategy rely on connection of tenant meters?",
    "")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Metering and Monitoring", "Descriptive",
    "If relying on tenant meters, describe the fitout guide or lease clauses ensuring meter connection and monitoring requirements.",
    "")
q += 1

# ── Commissioning and Tuning ──
row = add_criteria_header(ws3, row, "Commissioning and Tuning")

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Descriptive",
    "Summarise environmental performance targets (energy, water, IEQ, airtightness) set prior to schematic design.",
    "Target-setting practices driving building performance outcomes.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Condition (Y/N)",
    "Was the design intent report or OPR signed off by the building owner?",
    "")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Data",
    "Provide numerical targets for: (a) energy use intensity, (b) water consumption, (c) IEQ parameters, (d) airtightness rate.",
    "Design target benchmarking against actual operational performance.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Descriptive",
    "Summarise the services and maintainability review: participants, key outcomes, and close-out status in the Services and Maintainability Report.",
    "Stakeholder collaboration in pre-construction reviews.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Descriptive",
    "Identify the commissioning standard followed (e.g. AIRAH DA27, ASHRAE 202, CIBSE Code M). Outline the commissioning plan scope and program.",
    "Commissioning standard adoption rates.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Descriptive",
    "List all nominated building systems commissioned (e.g. HVAC, BMCS, lighting, electrical, hydraulic, fire, lifts).",
    "Extent of building systems commissioning across projects.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Descriptive",
    "Explain how airtightness targets were set (per ATTMA Guide) and how the air barrier schematic was reviewed before end of design development.",
    "Airtightness design integration practices.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Descriptive",
    "Describe airtightness testing: practitioner's ATTMA level, standard followed (AS/NZS ISO 9972), areas tested (whole building or sample), and selection of high-risk assemblies.",
    "Airtightness testing practices and practitioner qualifications.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Data",
    "Provide airtightness results (air permeability rates) per tested area. Note whether targets were met and any improvement opportunities identified.",
    "Building envelope quality benchmarking.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Descriptive",
    "Describe the tuning commitment: contractual arrangement, tuning plan, and team roles (FM, ICA, head contractor, subcontractors).",
    "Post-occupancy building optimisation approaches.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Data",
    "State tuning duration (min. 12 months), frequency of adjustments (min. quarterly), and planned start date.",
    "Tuning duration and frequency indicators.")
q += 1

# ── Building Information ──
row = add_criteria_header(ws3, row, "Building Information")

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Building Information", "Descriptive",
    "Summarise the O&M information provided: maintenance procedures, schedules, contacts, warranties, and as-built drawings for nominated systems.",
    "Handover documentation completeness.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Building Information", "Descriptive",
    "Explain how O&M information guides the FM team on keeping records current and responding to monitoring system alerts/faults.",
    "")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Building Information", "Descriptive",
    "Confirm a CIBSE TM31 building logbook was prepared covering all nominated systems and delivered to the owner prior to occupation.",
    "Building logbook adoption rates.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Building Information", "Descriptive",
    "Describe building user information: availability to occupants, relevance to audience, and digital format used (e.g. website, app, signage).",
    "Approaches to engaging occupants in sustainable operations.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Building Information", "Condition (Y/N)",
    "Is the building user information in an editable digital format accessible to the FM team?",
    "Digital information management maturity in building operations.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Minimum Expectation",
    "Building Information", "Descriptive",
    "State the format and platform used for building user information.",
    "")
q += 1

# ── Credit Achievement ──
row = add_level_header(ws3, row, "Credit Achievement (1 point)")

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Credit Achievement",
    "General", "Condition (Y/N)",
    "Is the Total Building Services Value over $20 million? (If yes, both Soft Landings and ICA criteria must be met.)",
    "Building services expenditure relative to commissioning requirements.")
q += 1

# ── Soft Landings Approach ──
row = add_criteria_header(ws3, row, "Soft Landings Approach")

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Credit Achievement",
    "Soft Landings Approach", "Descriptive",
    "Describe implementation of CIBSE ANZ Soft Landings Stages 1-4.",
    "Soft landings adoption and implementation quality.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Credit Achievement",
    "Soft Landings Approach", "Condition (Y/N)",
    "Are the sample worksheets for Stages 1-3 completed and Stage 4 actions identified?",
    "")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Credit Achievement",
    "Soft Landings Approach", "Descriptive",
    "Describe the FM team's involvement: commissioning participation, O&M manual development and sign-off, and pre-handover training received.",
    "FM involvement in building transition and performance gap.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Credit Achievement",
    "Soft Landings Approach", "Descriptive",
    "Explain arrangements for FM access to design and construction team members for two years post practical completion.",
    "Post-handover support duration and structure.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Credit Achievement",
    "Soft Landings Approach", "Condition (Y/N)",
    "Has Stage 5 (post-occupancy evaluation) been planned or implemented?",
    "Voluntary post-occupancy evaluation uptake.")
q += 1

# ── Independent Commissioning Agent ──
row = add_criteria_header(ws3, row, "Independent Commissioning Agent")

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Credit Achievement",
    "Independent Commissioning Agent", "Descriptive",
    "Identify the ICA: qualifications, commissioning knowledge, and experience with 2+ similar projects.",
    "ICA workforce capacity and qualification levels.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Credit Achievement",
    "Independent Commissioning Agent", "Descriptive",
    "Confirm the ICA was appointed before design development and is independent of all design/installation consultants and contractors.",
    "Independence arrangements in commissioning oversight.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Credit Achievement",
    "Independent Commissioning Agent", "Descriptive",
    "Summarise the ICA's involvement across phases: design development, tender, construction, commissioning, and tuning.",
    "Breadth of ICA involvement for commissioning effectiveness.")
q += 1

row = add_question(ws3, row, f"{ref_base}.{q}", "Verification and Handover", "Credit Achievement",
    "Independent Commissioning Agent", "Condition (Y/N)",
    "Is the ICA role fulfilled by more than one person?",
    "")
q += 1

apply_dropdowns(ws3)


# ════════════════════════════════════════════════════════════════════════════
# CREDIT 4: RESPONSIBLE RESOURCE MANAGEMENT
# ════════════════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet()
row = setup_sheet(ws4, "Responsible Resource Mgmt")
row = add_credit_header(ws4, row, "Responsible Resource Management",
    "Operational waste and resources can be separated and recovered in a safe and efficient manner.")
ref_base = "RRM"
q = 1

row = add_level_header(ws4, row, "Minimum Expectation (Nil points)")

# ── Collection of Waste Streams ──
row = add_criteria_header(ws4, row, "Collection of Waste Streams")

row = add_question(ws4, row, f"{ref_base}.{q}", "Responsible Resource Management", "Minimum Expectation",
    "Collection of Waste Streams", "Descriptive",
    "List all separately collected waste streams (min: general waste, paper/cardboard, glass, plastic, plus one additional). Justify the additional stream selected.",
    "Diversity of operational waste separation across building types.")
q += 1

row = add_question(ws4, row, f"{ref_base}.{q}", "Responsible Resource Management", "Minimum Expectation",
    "Collection of Waste Streams", "Condition (Y/N)",
    "Is any single non-food waste stream expected to exceed 5% of total annual operational waste by volume?",
    "Dominant waste streams for targeted reduction strategies.")
q += 1

row = add_question(ws4, row, f"{ref_base}.{q}", "Responsible Resource Management", "Minimum Expectation",
    "Collection of Waste Streams", "Descriptive",
    "If yes, identify the stream(s) exceeding 5% and describe their separate collection provisions.",
    "")
q += 1

row = add_question(ws4, row, f"{ref_base}.{q}", "Responsible Resource Management", "Minimum Expectation",
    "Collection of Waste Streams", "Descriptive",
    "Describe bin/chute intake locations, proximity to waste generation points, and labelling approach.",
    "Waste infrastructure design for occupant convenience.")
q += 1

row = add_question(ws4, row, f"{ref_base}.{q}", "Responsible Resource Management", "Minimum Expectation",
    "Collection of Waste Streams", "Condition (Y/N)",
    "Does the project include cold shell or excluded tenancy spaces outside the rating scope?",
    "")
q += 1

row = add_question(ws4, row, f"{ref_base}.{q}", "Responsible Resource Management", "Minimum Expectation",
    "Collection of Waste Streams", "Descriptive",
    "If yes, describe fitout guide, lease clauses, or contracts ensuring waste separation in those spaces.",
    "")
q += 1

row = add_question(ws4, row, f"{ref_base}.{q}", "Responsible Resource Management", "Minimum Expectation",
    "Collection of Waste Streams", "Condition (Y/N)",
    "Is co-mingled recycling used for any waste streams?",
    "Prevalence of co-mingled vs source-separated recycling.")
q += 1

row = add_question(ws4, row, f"{ref_base}.{q}", "Responsible Resource Management", "Minimum Expectation",
    "Collection of Waste Streams", "Descriptive",
    "If co-mingled, identify which streams and confirm acceptance by the waste collection service.",
    "")
q += 1

# ── Dedicated Waste Storage Area ──
row = add_criteria_header(ws4, row, "Dedicated Waste Storage Area")

row = add_question(ws4, row, f"{ref_base}.{q}", "Responsible Resource Management", "Minimum Expectation",
    "Dedicated Waste Storage Area", "Descriptive",
    "Describe the storage area(s): location, total area, and layout for keeping waste streams separate.",
    "Waste storage design patterns across building types.")
q += 1

row = add_question(ws4, row, f"{ref_base}.{q}", "Responsible Resource Management", "Minimum Expectation",
    "Dedicated Waste Storage Area", "Data",
    "Provide forecasted waste generation rates, collection frequency per stream, and storage capacity calculations. Identify the best practice guideline used.",
    "Waste generation rate data for cross-project benchmarking.")
q += 1

row = add_question(ws4, row, f"{ref_base}.{q}", "Responsible Resource Management", "Minimum Expectation",
    "Dedicated Waste Storage Area", "Descriptive",
    "Describe collection vehicle access: parking, driveways, height clearances, and manoeuvring per AS 2890.2:2018.",
    "Waste collection logistics design.")
q += 1

row = add_question(ws4, row, f"{ref_base}.{q}", "Responsible Resource Management", "Minimum Expectation",
    "Dedicated Waste Storage Area", "Condition (Y/N)",
    "Is this a tenanted building where excluded tenancies contribute to the waste storage strategy?",
    "")
q += 1

row = add_question(ws4, row, f"{ref_base}.{q}", "Responsible Resource Management", "Minimum Expectation",
    "Dedicated Waste Storage Area", "Descriptive",
    "If yes, explain how waste from excluded tenancies was estimated and factored into storage sizing.",
    "")
q += 1

# ── Safe and Efficient Access to Waste Storage ──
row = add_criteria_header(ws4, row, "Safe and Efficient Access to Waste Storage")

row = add_question(ws4, row, f"{ref_base}.{q}", "Responsible Resource Management", "Minimum Expectation",
    "Safe and Efficient Access to Waste Storage", "Descriptive",
    "Identify the waste specialist/contractor who signed off on designs, including their organisation and relevant experience (min. 3 years).",
    "Waste specialist involvement in building design.")
q += 1

row = add_question(ws4, row, f"{ref_base}.{q}", "Responsible Resource Management", "Minimum Expectation",
    "Safe and Efficient Access to Waste Storage", "Descriptive",
    "Summarise the sign-off findings confirming storage areas are adequately sized and located for safe collection.",
    "Waste management design validation practices.")
q += 1

row = add_question(ws4, row, f"{ref_base}.{q}", "Responsible Resource Management", "Minimum Expectation",
    "Safe and Efficient Access to Waste Storage", "Data",
    "Provide: building GFA, number of occupants/units, waste storage area (m²), and estimated annual operational waste (tonnes or m³/year).",
    "Waste generation benchmarks normalised by building size and occupancy.")
q += 1

apply_dropdowns(ws4)


# ════════════════════════════════════════════════════════════════════════════
# CREDIT 5: RESPONSIBLE PROCUREMENT
# ════════════════════════════════════════════════════════════════════════════
ws5 = wb.create_sheet()
row = setup_sheet(ws5, "Responsible Procurement")
row = add_credit_header(ws5, row, "Responsible Procurement",
    "The procurement process for key products, materials, and services follows best practice environmental and social principles.")
ref_base = "RP"
q = 1

row = add_level_header(ws5, row, "Credit Achievement (1 point)")

# ── Risk and Opportunity Assessment ──
row = add_criteria_header(ws5, row, "Risk and Opportunity Assessment")

row = add_question(ws5, row, f"{ref_base}.{q}", "Responsible Procurement", "Credit Achievement",
    "Risk and Opportunity Assessment", "Condition (Y/N)",
    "Was the risk and opportunity assessment completed before appointment of the head contractor?",
    "Timing of supply chain risk assessment relative to procurement.")
q += 1

row = add_question(ws5, row, f"{ref_base}.{q}", "Responsible Procurement", "Credit Achievement",
    "Risk and Opportunity Assessment", "Condition (Y/N)",
    "Did the building owner provide input into the assessment?",
    "Stakeholder involvement in supply chain risk assessment.")
q += 1

row = add_question(ws5, row, f"{ref_base}.{q}", "Responsible Procurement", "Credit Achievement",
    "Risk and Opportunity Assessment", "Descriptive",
    "Identify who conducted the assessment.",
    "")
q += 1

row = add_question(ws5, row, f"{ref_base}.{q}", "Responsible Procurement", "Credit Achievement",
    "Risk and Opportunity Assessment", "Descriptive",
    "List the 10+ key supply chain items (min. 2 building services, 1 building material). Briefly justify each selection.",
    "Supply chain risk hotspots across building projects.")
q += 1

row = add_question(ws5, row, f"{ref_base}.{q}", "Responsible Procurement", "Credit Achievement",
    "Risk and Opportunity Assessment", "Descriptive",
    "Describe how risks and opportunities were evaluated per ISO 20400 Clause 4.3: human rights, labour, environment, fair practices, consumer issues, community.",
    "Depth of sustainability risk analysis in procurement.")
q += 1

row = add_question(ws5, row, f"{ref_base}.{q}", "Responsible Procurement", "Credit Achievement",
    "Risk and Opportunity Assessment", "Data",
    "For each key item, summarise priority risks/opportunities and risk ratings (high/medium/low) per ISO 20400 issue area.",
    "Aggregated supply chain risk profiles for industry guidance.")
q += 1

row = add_question(ws5, row, f"{ref_base}.{q}", "Responsible Procurement", "Credit Achievement",
    "Risk and Opportunity Assessment", "Descriptive",
    "Explain the methodology used to analyse and prioritise risks. Note any tools or references beyond ISO 20400 Annex A.",
    "Risk assessment methodologies for industry knowledge sharing.")
q += 1

# ── Responsible Procurement Plan ──
row = add_criteria_header(ws5, row, "Responsible Procurement Plan")

row = add_question(ws5, row, f"{ref_base}.{q}", "Responsible Procurement", "Credit Achievement",
    "Responsible Procurement Plan", "Descriptive",
    "Outline the plan's environmental, social, and economic objectives addressing the identified risks and opportunities.",
    "Procurement sustainability objective ambition across projects.")
q += 1

row = add_question(ws5, row, f"{ref_base}.{q}", "Responsible Procurement", "Credit Achievement",
    "Responsible Procurement Plan", "Descriptive",
    "Describe data collection, monitoring, and reporting requirements per ISO 20400 Clause 6.5. State metrics tracked and frequency.",
    "Procurement monitoring and reporting approaches.")
q += 1

row = add_question(ws5, row, f"{ref_base}.{q}", "Responsible Procurement", "Credit Achievement",
    "Responsible Procurement Plan", "Descriptive",
    "Describe the framework for incentivising contractors and trades. Provide examples of incentive mechanisms.",
    "Supply chain incentive models for sustainability outcomes.")
q += 1

row = add_question(ws5, row, f"{ref_base}.{q}", "Responsible Procurement", "Credit Achievement",
    "Responsible Procurement Plan", "Descriptive",
    "Explain how the plan was embedded in tender documentation for the head contractor and relevant trades.",
    "Integration of sustainability requirements into procurement workflows.")
q += 1

row = add_question(ws5, row, f"{ref_base}.{q}", "Responsible Procurement", "Credit Achievement",
    "Responsible Procurement Plan", "Condition (Y/N)",
    "Was the head contractor engaged under a design and construct (D&C) contract?",
    "Procurement models and their impact on sustainability integration.")
q += 1

row = add_question(ws5, row, f"{ref_base}.{q}", "Responsible Procurement", "Credit Achievement",
    "Responsible Procurement Plan", "Descriptive",
    "If D&C, explain the head contractor's role in developing the plan and how it was embedded in subcontractor tenders.",
    "")
q += 1

row = add_question(ws5, row, f"{ref_base}.{q}", "Responsible Procurement", "Credit Achievement",
    "Responsible Procurement Plan", "Descriptive",
    "Describe plan implementation during construction: data collection, monitoring, and reporting activities per ISO 20400 Clause 7.",
    "Real-world sustainable procurement implementation effectiveness.")
q += 1

row = add_question(ws5, row, f"{ref_base}.{q}", "Responsible Procurement", "Credit Achievement",
    "Responsible Procurement Plan", "Data",
    "List key items with procurement actions taken. Summarise the sustainability outcome per item (e.g. modern slavery risk mitigated, local supply used).",
    "Outcome-level data on responsible procurement for advocacy/policy.")
q += 1

row = add_question(ws5, row, f"{ref_base}.{q}", "Responsible Procurement", "Credit Achievement",
    "Responsible Procurement Plan", "Data",
    "Note any supply chain risks that materialised and corrective actions taken. State items fully vs partially implemented.",
    "Procurement plan implementation rates and real-world risk events.")
q += 1

apply_dropdowns(ws5)


# ════════════════════════════════════════════════════════════════════════════
# Save
# ════════════════════════════════════════════════════════════════════════════
output_path = "/home/user/submissionforms/Green_Star_Responsible_Credits_Submission_Questions.xlsx"
wb.save(output_path)
print(f"Saved to {output_path}")
