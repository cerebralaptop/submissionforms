import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

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

# Column widths: A=Ref, B=Credit, C=Performance Level, D=Criteria, E=Type, F=Question, G=Response, H=Data Collection Notes
COL_WIDTHS = {"A": 8, "B": 20, "C": 22, "D": 28, "E": 16, "F": 65, "G": 60, "H": 45}


def setup_sheet(ws, title):
    ws.title = title
    for col_letter, width in COL_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

    headers = [
        "Ref",
        "Credit",
        "Performance Level",
        "Criteria",
        "Question Type",
        "Question",
        "Response",
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
    return 2  # next row


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
    values = [ref, credit, level, criteria, q_type, question, "", data_note]
    for col, val in enumerate(values, 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.alignment = wrap
        cell.border = thin_border
        if q_type == "Condition (Y/N)":
            cell.fill = condition_fill
            if col == 5:
                cell.font = condition_font
            else:
                cell.font = question_font
        elif col == 8 and data_note:
            cell.fill = data_fill
            cell.font = data_flag_font
        else:
            cell.fill = question_fill if col != 7 else white_fill
            cell.font = question_font
    ws.row_dimensions[row].height = 60
    return row + 1


# ════════════════════════════════════════════════════════════════════════════
# CREDIT 1: INDUSTRY DEVELOPMENT
# ════════════════════════════════════════════════════════════════════════════
ws = wb.active
row = setup_sheet(ws, "Industry Development")
row = add_credit_header(ws, row, "Industry Development",
    "The development facilitates industry transformation through partnership, collaboration and data sharing.")

# --- Credit Achievement ---
row = add_level_header(ws, row, "Credit Achievement (1 point)")

# Criteria: Green Star Accredited Professional
row = add_criteria_header(ws, row, "Green Star Accredited Professional")
ref_base = "ID"
q_num = 1

row = add_question(ws, row, f"{ref_base}.{q_num}", "Industry Development", "Credit Achievement",
    "Green Star Accredited Professional", "Descriptive",
    "Identify the Green Star Accredited Professional(s) (GSAP) engaged on the project, including their name, organisation, accreditation number, and the Green Star Buildings accreditation they hold.",
    "Tracks GSAP workforce capacity and distribution across projects.")
q_num += 1

row = add_question(ws, row, f"{ref_base}.{q_num}", "Industry Development", "Credit Achievement",
    "Green Star Accredited Professional", "Descriptive",
    "State the date the GSAP was first engaged on the project and the project phase at that time (e.g. concept design, schematic design). Confirm whether this was within one month of project registration.",
    "Measures timing of sustainability expertise integration relative to design stage.")
q_num += 1

row = add_question(ws, row, f"{ref_base}.{q_num}", "Industry Development", "Credit Achievement",
    "Green Star Accredited Professional", "Descriptive",
    "Describe the scope of the GSAP's engagement, including the advisory, coordination and support activities undertaken with the project team on Green Star strategy, principles, structure, timing and certification process.",
    "Captures the depth of sustainability advisory services on projects.")
q_num += 1

row = add_question(ws, row, f"{ref_base}.{q_num}", "Industry Development", "Credit Achievement",
    "Green Star Accredited Professional", "Condition (Y/N)",
    "Has the GSAP role been fulfilled by more than one individual or organisation during the project? (Yes/No)",
    "Tracks continuity of sustainability expertise across project lifecycle.")
q_num += 1

row = add_question(ws, row, f"{ref_base}.{q_num}", "Industry Development", "Credit Achievement",
    "Green Star Accredited Professional", "Descriptive",
    "If multiple GSAPs were involved, explain the transition between individuals and confirm each was accredited for Green Star Buildings for the duration of their engagement.",
    "")
q_num += 1

row = add_question(ws, row, f"{ref_base}.{q_num}", "Industry Development", "Credit Achievement",
    "Green Star Accredited Professional", "Descriptive",
    "Confirm the GSAP is nominated as the 'Project Contact' for GBCA communications. Describe how ongoing input was maintained throughout the project (e.g. attendance at design meetings, coordination workshops).",
    "")
q_num += 1

# Criteria: Financial Transparency
row = add_criteria_header(ws, row, "Financial Transparency")

row = add_question(ws, row, f"{ref_base}.{q_num}", "Industry Development", "Credit Achievement",
    "Financial Transparency", "Descriptive",
    "Confirm that the Financial Transparency disclosure template has been completed in its latest version and submitted in Excel format. Identify who prepared the cost data (e.g. quantity surveyor, head contractor, cost consultant).",
    "Enables industry-wide benchmarking of sustainable building costs.")
q_num += 1

row = add_question(ws, row, f"{ref_base}.{q_num}", "Industry Development", "Credit Achievement",
    "Financial Transparency", "Descriptive",
    "Describe the methodology used to determine the documentation cost and implementation cost of sustainable building practices, including how costs beyond the base (non-Green Star) requirement were isolated or estimated.",
    "Supports research into cost premiums/savings of green building practices.")
q_num += 1

row = add_question(ws, row, f"{ref_base}.{q_num}", "Industry Development", "Credit Achievement",
    "Financial Transparency", "Data",
    "Provide the total project construction cost and the total additional cost attributed to sustainable building practices (documentation + implementation).",
    "Critical data point for cost-benefit analysis of green certification across the industry.")
q_num += 1

# Criteria: Marketing Sustainability Achievements
row = add_criteria_header(ws, row, "Marketing Sustainability Achievements")

row = add_question(ws, row, f"{ref_base}.{q_num}", "Industry Development", "Credit Achievement",
    "Marketing Sustainability Achievements", "Descriptive",
    "List which three (or more) of the following marketing activities have been undertaken: (a) case study provided to GBCA, (b) digital screens installed to promote rating, (c) construction hoarding displays GBCA and targeted rating, (d) GBCA and rating central to marketing/communications strategy and promotional material.",
    "Tracks industry adoption of sustainability marketing practices.")
q_num += 1

row = add_question(ws, row, f"{ref_base}.{q_num}", "Industry Development", "Credit Achievement",
    "Marketing Sustainability Achievements", "Descriptive",
    "Describe how the sustainability achievements and benefits of the project are communicated to building users, the public, and/or prospective tenants/buyers through the selected marketing channels.",
    "Measures effectiveness and reach of green building awareness campaigns.")
q_num += 1

row = add_question(ws, row, f"{ref_base}.{q_num}", "Industry Development", "Credit Achievement",
    "Marketing Sustainability Achievements", "Data",
    "Identify the target audience(s) for each sustainability marketing activity and the estimated reach (e.g. building occupants, public foot traffic, website visitors).",
    "Quantifies public awareness exposure to green building benefits for future advocacy research.")
q_num += 1


# ════════════════════════════════════════════════════════════════════════════
# CREDIT 2: RESPONSIBLE CONSTRUCTION
# ════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet()
row = setup_sheet(ws2, "Responsible Construction")
row = add_credit_header(ws2, row, "Responsible Construction",
    "The builder's construction practices reduce impacts and promote opportunities for improved environmental and social outcomes.")
ref_base = "RC"
q_num = 1

# --- Minimum Expectation ---
row = add_level_header(ws2, row, "Minimum Expectation (Nil points)")

# Environmental Management System
row = add_criteria_header(ws2, row, "Environmental Management System")

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management System", "Condition (Y/N)",
    "Is the total contract value for any site works package (demolition, early works, or main works) $10 million or more? (Yes/No)",
    "Benchmarks contract sizes relative to EMS certification thresholds.")
q_num += 1

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management System", "Descriptive",
    "For contracts valued at less than $10 million, describe how the Environmental Management System (EMS) complies with the NSW Environmental Management System Guidelines or another recognised framework. Identify the framework used.",
    "Tracks which EMS frameworks are most commonly adopted by the construction industry.")
q_num += 1

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management System", "Descriptive",
    "For contracts valued at $10 million or more, identify the standard to which the EMS is certified (AS/NZS ISO 14001, BS 7750, or EMAS) and confirm the certification was valid for the entire duration of site activities.",
    "Measures uptake of certified environmental management in the construction sector.")
q_num += 1

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management System", "Condition (Y/N)",
    "Were there different head contractors for demolition, early works, and main works? (Yes/No)",
    "")
q_num += 1

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management System", "Descriptive",
    "If multiple head contractors were engaged, confirm each had an EMS in place for their scope of works and explain how the contract values were apportioned.",
    "")
q_num += 1

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management System", "Descriptive",
    "Explain how the EMS includes actions related to the implementation of the Environmental Management Plan (EMP), and describe the key environmental impacts targeted.",
    "Documents the relationship between management systems and on-site environmental outcomes.")
q_num += 1

# Environmental Management Plan
row = add_criteria_header(ws2, row, "Environmental Management Plan")

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management Plan", "Descriptive",
    "Describe the project-specific Environmental Management Plan (EMP), including the key environmental performance conditions and impact areas it addresses (e.g. noise, dust, stormwater, vegetation protection).",
    "Identifies the most common environmental risks managed during construction.")
q_num += 1

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management Plan", "Descriptive",
    "Confirm the EMP was in place for the full duration of all site works (demolition, early works, and main works). State the start and end dates of the EMP's application.",
    "Tracks construction duration and environmental management coverage periods.")
q_num += 1

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management Plan", "Descriptive",
    "Describe the monitoring, auditing and reporting regime under the EMP, including the frequency of site audits and how non-conformances were managed and closed out.",
    "Captures data on environmental compliance enforcement during construction.")
q_num += 1

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Minimum Expectation",
    "Environmental Management Plan", "Data",
    "State the total number of environmental audits conducted, the number of non-conformances identified, and the percentage closed out during construction.",
    "Provides quantitative data on construction environmental management performance.")
q_num += 1

# Construction and Demolition Waste Diversion
row = add_criteria_header(ws2, row, "Construction and Demolition Waste Diversion")

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Minimum Expectation",
    "Construction and Demolition Waste Diversion", "Data",
    "State the total mass of site waste generated (in tonnes) and the total mass diverted from landfill. Provide the diversion rate as a percentage. Confirm this meets or exceeds the 80% threshold.",
    "Critical benchmark data for construction waste diversion rates across the industry.")
q_num += 1

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Minimum Expectation",
    "Construction and Demolition Waste Diversion", "Descriptive",
    "Describe the waste streams generated on site and the diversion pathways used (e.g. recycling, reuse, recovery). Identify any waste streams excluded from the calculation (e.g. special waste, excavation waste) and justify the exclusions.",
    "Maps waste stream composition and recycling pathways in the construction sector.")
q_num += 1

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Minimum Expectation",
    "Construction and Demolition Waste Diversion", "Descriptive",
    "Explain how the waste contractors and processing facilities have provided a Disclosure Statement outlining alignment with the Green Star Construction and Demolition Waste Reporting Criteria.",
    "Tracks supply chain transparency in waste management reporting.")
q_num += 1

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Minimum Expectation",
    "Construction and Demolition Waste Diversion", "Data",
    "Provide a breakdown of waste by material type (e.g. concrete, timber, steel, plasterboard, mixed) showing tonnes generated and tonnes diverted for each.",
    "Enables material-specific waste benchmarking across construction projects.")
q_num += 1

# Sustainability Training
row = add_criteria_header(ws2, row, "Sustainability Training")

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Minimum Expectation",
    "Sustainability Training", "Data",
    "State the total number of contractors and subcontractors present on site for at least three days during all site works, the number who received sustainability training, and the resulting percentage.",
    "Measures the reach of sustainability education in the construction workforce.")
q_num += 1

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Minimum Expectation",
    "Sustainability Training", "Descriptive",
    "Describe the content of the sustainability training provided, including how it covered: (a) sustainability attributes of the project and their benefits, (b) the value of certification, and (c) the role site workers play in delivering a sustainable building.",
    "Captures training content quality and scope for workforce sustainability literacy research.")
q_num += 1

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Minimum Expectation",
    "Sustainability Training", "Descriptive",
    "Explain the delivery method and timing of the training (e.g. site induction, toolbox talks, dedicated sessions) and how attendance was tracked and verified.",
    "Documents effective training delivery models for sustainability in construction.")
q_num += 1

# --- Credit Achievement ---
row = add_level_header(ws2, row, "Credit Achievement (1 point)")
row = add_criteria_header(ws2, row, "Increased Construction and Demolition Waste Diversion")

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Credit Achievement",
    "Increased Construction and Demolition Waste Diversion", "Data",
    "State the total mass of site waste generated (in tonnes) and the total mass diverted from landfill. Confirm the diversion rate meets or exceeds 90%.",
    "Higher-tier waste diversion benchmarking for industry best practice.")
q_num += 1

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Credit Achievement",
    "Increased Construction and Demolition Waste Diversion", "Descriptive",
    "Describe how waste contractors and processing facilities comply with the Green Star Construction and Demolition Waste Reporting Criteria, including provision of the Compliance Verification Summary.",
    "Assesses third-party waste reporting verification practices.")
q_num += 1

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Credit Achievement",
    "Increased Construction and Demolition Waste Diversion", "Descriptive",
    "Identify the auditor(s) who verified waste reporting, including their credentials as specified in the Green Star Construction and Demolition Waste Reporting Criteria.",
    "Tracks auditor capacity and verification standards in waste management.")
q_num += 1

row = add_question(ws2, row, f"{ref_base}.{q_num}", "Responsible Construction", "Credit Achievement",
    "Increased Construction and Demolition Waste Diversion", "Data",
    "List the waste processing facilities used, their location, and the types of waste they processed. Indicate whether any hold GECA Construction and Demolition Waste Services Standard certification.",
    "Maps waste processing infrastructure availability and certification uptake.")
q_num += 1


# ════════════════════════════════════════════════════════════════════════════
# CREDIT 3: VERIFICATION AND HANDOVER
# ════════════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet()
row = setup_sheet(ws3, "Verification and Handover")
row = add_credit_header(ws3, row, "Verification and Handover",
    "The building has been optimised and handed over to deliver a higher level of performance in operation.")
ref_base = "VH"
q_num = 1

# --- Minimum Expectation ---
row = add_level_header(ws3, row, "Minimum Expectation (Nil points)")

# Metering and Monitoring
row = add_criteria_header(ws3, row, "Metering and Monitoring")

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Metering and Monitoring", "Descriptive",
    "Describe the metering strategy for the building, identifying how energy and water metering is provided for all distinct uses, major uses, and separate tenancies or units. Reference the metering schedule prepared in accordance with CIBSE TM39 (steps 7-10).",
    "Captures metering granularity across building types for energy/water benchmarking research.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Metering and Monitoring", "Data",
    "State the total number of energy meters and water meters installed, including utility meters and sub-meters. Provide the number of distinct end-uses metered.",
    "Quantifies metering density for benchmarking across building typologies.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Metering and Monitoring", "Descriptive",
    "Confirm all meters (utility and sub-meters) provide continual information at up to 1-hour intervals to the monitoring system, are commissioned and validated in accordance with NABERS Metering and Consumption Rules, and are pattern approved by the NMI or meet another recognised standard.",
    "Tracks metering quality standards adoption.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Metering and Monitoring", "Descriptive",
    "Describe the automatic monitoring system implemented, including how it provides consumption trend reports and generates alarms when energy or water use exceeds set parameters. Explain how alerts are communicated to the facilities manager.",
    "Documents monitoring system capabilities for operational performance research.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Metering and Monitoring", "Condition (Y/N)",
    "Is this a Class 2 build-to-sell apartment project? (Yes/No)",
    "")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Metering and Monitoring", "Descriptive",
    "If this is a Class 2 build-to-sell project, confirm that base building consumption trends are provided to the facilities manager and explain how individual unit meters are handled.",
    "")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Metering and Monitoring", "Condition (Y/N)",
    "Does the base building metering strategy rely on connection of tenant meters? (Yes/No)",
    "")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Metering and Monitoring", "Descriptive",
    "If the strategy relies on tenant meters, describe the tenancy fitout guide and/or model lease clauses used to ensure tenant meter connection and monitoring system programming requirements are met.",
    "")
q_num += 1

# Commissioning and Tuning
row = add_criteria_header(ws3, row, "Commissioning and Tuning")

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Descriptive",
    "Describe the environmental performance targets documented prior to schematic design, covering energy use, water consumption, indoor environment quality, and airtightness. Confirm the design intent report or OPR was signed off by the building owner.",
    "Captures the target-setting practices that drive building performance outcomes.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Data",
    "Provide the specific numerical targets set for: (a) energy use intensity (kWh/m²/yr or similar), (b) water consumption (kL/yr or similar), (c) indoor environment quality parameters, and (d) airtightness (air permeability rate).",
    "Enables benchmarking of design targets against actual operational performance across projects.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Descriptive",
    "Describe the services and maintainability review conducted prior to construction, including who participated (building owner, design consultants, architect, facilities manager, head contractor, ICA) and the key outcomes documented in the Services and Maintainability Report.",
    "Documents stakeholder collaboration in pre-construction reviews.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Descriptive",
    "Identify the recognised commissioning standard or guideline followed (e.g. AIRAH DA27, ASHRAE Standard 202-2024, CIBSE Commissioning Code M, SA TA 5342:2021). Describe the commissioning plan, including the process, activities and program for commissioning all nominated building systems.",
    "Tracks adoption of commissioning standards across the industry.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Descriptive",
    "List all nominated building systems included in the commissioning scope (e.g. HVAC, BMCS, lighting, electrical, hydraulic, fire, lifts). Confirm commissioning requirements were included in construction documentation prior to the start of relevant trade packages.",
    "Maps the extent of building systems commissioning across projects.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Descriptive",
    "Describe how airtightness targets were set based on the ATTMA Australia Guide for Airtightness Targets, including how targets were defined for different building compartments where applicable. Explain how the air barrier system schematic was reviewed prior to end of design development to reduce risks.",
    "Captures industry approach to airtightness design integration.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Descriptive",
    "Describe the airtightness testing undertaken, including: the testing practitioner's ATTMA membership level, the testing standard followed (AS/NZS ISO 9972 Method 1), the areas tested (whole building or sample), and how both typical and high-risk assemblies were selected.",
    "Provides data on airtightness testing practices and practitioner qualifications.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Data",
    "Provide the airtightness test results (air permeability rates achieved) for each tested area. State whether the targets were met and, if not, identify the opportunities for improvement shared with the building owner.",
    "Key performance data for benchmarking building envelope quality across the industry.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Descriptive",
    "Describe the building systems tuning commitment, including: the contractual arrangement between building owner and head contractor, the tuning manual/plan, and the roles and responsibilities of the tuning team (facilities manager, ICA, head contractor, subcontractors).",
    "Documents industry approaches to post-occupancy building optimisation.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Commissioning and Tuning", "Data",
    "State the planned duration of the tuning process (minimum 12 months) and the frequency of adjustments and measurements (minimum quarterly). Identify the start date of the tuning period.",
    "Tracks tuning duration and frequency as indicators of operational optimisation commitment.")
q_num += 1

# Building Information
row = add_criteria_header(ws3, row, "Building Information")

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Building Information", "Descriptive",
    "Describe the operations and maintenance information provided to the building owner, including how it covers maintenance procedures, schedules, service contacts, warranties, and as-built drawings for all nominated building systems.",
    "Assesses completeness of handover documentation practices.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Building Information", "Descriptive",
    "Explain how the operations and maintenance information includes guidance for the facilities management team on keeping the information up to date and assessing, correcting, and validating alerts or faults from the monitoring system.",
    "")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Building Information", "Descriptive",
    "Describe the building logbook developed in accordance with CIBSE TM31 Building Logbook Toolkit. Confirm it covers all nominated building systems and was presented to the building owner prior to occupation.",
    "Tracks adoption of structured building logbook practices.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Building Information", "Descriptive",
    "Describe the building user information provided, including: how it is publicly available to intended building users, how the content is relevant and easily understood by the target audience, and the digital format used (e.g. website, app, digital signage).",
    "Captures approaches to engaging building occupants in sustainable operations.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Minimum Expectation",
    "Building Information", "Descriptive",
    "Confirm the building user information is provided in an editable, digital format accessible to the facilities management team for updates. Describe the format and platform used.",
    "Tracks digital information management maturity in building operations.")
q_num += 1

# --- Credit Achievement ---
row = add_level_header(ws3, row, "Credit Achievement (1 point)")

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Credit Achievement",
    "General", "Condition (Y/N)",
    "Is the Total Building Services Value of the project over $20 million? (Yes/No) Note: If yes, both Soft Landings Approach and Independent Commissioning Agent criteria must be met.",
    "Benchmarks building services expenditure relative to commissioning requirements.")
q_num += 1

# Soft Landings Approach
row = add_criteria_header(ws3, row, "Soft Landings Approach")

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Credit Achievement",
    "Soft Landings Approach", "Descriptive",
    "Describe how Stages 1 to 4 of the CIBSE ANZ Soft Landings Framework Australia and New Zealand have been implemented on the project. Confirm the sample worksheets for Stages 1-3 are completed and actions for Stage 4 are identified.",
    "Tracks adoption and implementation quality of soft landings framework across projects.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Credit Achievement",
    "Soft Landings Approach", "Descriptive",
    "Describe how the facilities management team (or building owner's representative) was involved in the soft landings approach, including their role in: commissioning and handover, developing the brief technical guide and operations manual, sign-off on the O&M manual, and training received before handover.",
    "Documents FM involvement in building transition for performance gap research.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Credit Achievement",
    "Soft Landings Approach", "Descriptive",
    "Explain the arrangements in place for the facilities management team to have continued access to critical design and construction team members for two years after practical completion.",
    "Measures post-handover support duration and structure.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Credit Achievement",
    "Soft Landings Approach", "Condition (Y/N)",
    "Has Stage 5 (post-occupancy evaluation) of the Soft Landings Framework been planned or implemented? (Yes/No)",
    "Tracks voluntary post-occupancy evaluation uptake.")
q_num += 1

# Independent Commissioning Agent
row = add_criteria_header(ws3, row, "Independent Commissioning Agent")

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Credit Achievement",
    "Independent Commissioning Agent", "Descriptive",
    "Identify the Independent Commissioning Agent (ICA) appointed, including their qualifications (registered professional engineer or qualified technician), demonstrated knowledge of commissioning, and experience with at least 2 projects of similar scope.",
    "Tracks ICA workforce capacity and qualification levels.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Credit Achievement",
    "Independent Commissioning Agent", "Descriptive",
    "Confirm the ICA was appointed prior to design development and is independent of any consultant, contractor, or sub-contractor involved in design or installation. Describe who the ICA reports to and their relationship to the project owner.",
    "Documents independence arrangements in commissioning oversight.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Credit Achievement",
    "Independent Commissioning Agent", "Descriptive",
    "Describe the ICA's involvement across project phases (design development, tender, construction, commissioning, and tuning), including key activities and oversight provided at each stage.",
    "Captures the breadth of ICA involvement for commissioning effectiveness research.")
q_num += 1

row = add_question(ws3, row, f"{ref_base}.{q_num}", "Verification and Handover", "Credit Achievement",
    "Independent Commissioning Agent", "Condition (Y/N)",
    "Is the ICA role fulfilled by more than one person? (Yes/No) If yes, confirm each meets the qualification and independence requirements.",
    "")
q_num += 1


# ════════════════════════════════════════════════════════════════════════════
# CREDIT 4: RESPONSIBLE RESOURCE MANAGEMENT
# ════════════════════════════════════════════════════════════════════════════
ws4 = wb.create_sheet()
row = setup_sheet(ws4, "Responsible Resource Mgmt")
row = add_credit_header(ws4, row, "Responsible Resource Management",
    "Operational waste and resources can be separated and recovered in a safe and efficient manner.")
ref_base = "RRM"
q_num = 1

# --- Minimum Expectation ---
row = add_level_header(ws4, row, "Minimum Expectation (Nil points)")

# Collection of Waste Streams
row = add_criteria_header(ws4, row, "Collection of Waste Streams")

row = add_question(ws4, row, f"{ref_base}.{q_num}", "Responsible Resource Management", "Minimum Expectation",
    "Collection of Waste Streams", "Descriptive",
    "List all the waste streams the building enables to be collected separately, including as a minimum: general waste, paper and cardboard, glass, plastic, and one additional stream (e.g. organics, e-waste, batteries). Identify the additional stream selected and justify why it was chosen.",
    "Tracks the diversity of operational waste separation across building types.")
q_num += 1

row = add_question(ws4, row, f"{ref_base}.{q_num}", "Responsible Resource Management", "Minimum Expectation",
    "Collection of Waste Streams", "Condition (Y/N)",
    "Is any single non-food waste stream (excluding the listed recycling streams) expected to represent more than 5% of total annual operational waste by volume? (Yes/No)",
    "Identifies dominant waste streams for targeted waste reduction strategies.")
q_num += 1

row = add_question(ws4, row, f"{ref_base}.{q_num}", "Responsible Resource Management", "Minimum Expectation",
    "Collection of Waste Streams", "Descriptive",
    "If yes, identify the waste stream(s) exceeding 5% and describe the separate collection provisions made for each.",
    "")
q_num += 1

row = add_question(ws4, row, f"{ref_base}.{q_num}", "Responsible Resource Management", "Minimum Expectation",
    "Collection of Waste Streams", "Descriptive",
    "Describe the location and distribution of chute intakes, bins, or storage containers across the building. Explain how they are positioned close to points of waste generation and how they are labelled for each waste stream.",
    "Documents waste infrastructure design approaches for occupant convenience research.")
q_num += 1

row = add_question(ws4, row, f"{ref_base}.{q_num}", "Responsible Resource Management", "Minimum Expectation",
    "Collection of Waste Streams", "Condition (Y/N)",
    "Does the project include cold shell or excluded tenancy spaces where fitout is not within the scope of the rating? (Yes/No)",
    "")
q_num += 1

row = add_question(ws4, row, f"{ref_base}.{q_num}", "Responsible Resource Management", "Minimum Expectation",
    "Collection of Waste Streams", "Descriptive",
    "If the project includes cold shell or excluded tenancy spaces, describe the tenancy fitout guide, model lease clauses, or supply contracts used to ensure waste stream separation requirements are met within those spaces.",
    "")
q_num += 1

row = add_question(ws4, row, f"{ref_base}.{q_num}", "Responsible Resource Management", "Minimum Expectation",
    "Collection of Waste Streams", "Condition (Y/N)",
    "Is co-mingled recycling used for any waste streams? (Yes/No)",
    "Tracks prevalence of co-mingled vs source-separated recycling.")
q_num += 1

row = add_question(ws4, row, f"{ref_base}.{q_num}", "Responsible Resource Management", "Minimum Expectation",
    "Collection of Waste Streams", "Descriptive",
    "If co-mingled recycling is used, identify which streams are co-mingled and confirm this is accepted by the waste collection service. Confirm remaining streams are still collected separately.",
    "")
q_num += 1

# Dedicated Waste Storage Area
row = add_criteria_header(ws4, row, "Dedicated Waste Storage Area")

row = add_question(ws4, row, f"{ref_base}.{q_num}", "Responsible Resource Management", "Minimum Expectation",
    "Dedicated Waste Storage Area", "Descriptive",
    "Describe the dedicated waste storage area(s), including location within the building, total area provided, and how the space is laid out to keep all applicable waste streams separate prior to off-site collection.",
    "Captures waste storage design patterns across building types.")
q_num += 1

row = add_question(ws4, row, f"{ref_base}.{q_num}", "Responsible Resource Management", "Minimum Expectation",
    "Dedicated Waste Storage Area", "Data",
    "Provide the forecasted waste generation rates used to size the storage area, the collection frequency assumed for each waste stream, and the resulting storage capacity calculations. Identify the third-party best practice guideline used for waste generation rates.",
    "Provides waste generation rate data for cross-project benchmarking research.")
q_num += 1

row = add_question(ws4, row, f"{ref_base}.{q_num}", "Responsible Resource Management", "Minimum Expectation",
    "Dedicated Waste Storage Area", "Descriptive",
    "Describe how collection vehicles can safely access the waste storage area, including provisions for parking adjacent to the area, driveways, height clearances, and manoeuvring areas in accordance with AS 2890.2:2018.",
    "Documents waste collection logistics design for operational efficiency research.")
q_num += 1

row = add_question(ws4, row, f"{ref_base}.{q_num}", "Responsible Resource Management", "Minimum Expectation",
    "Dedicated Waste Storage Area", "Condition (Y/N)",
    "Is this a tenanted building where excluded tenancy spaces contribute to the waste storage strategy? (Yes/No)",
    "")
q_num += 1

row = add_question(ws4, row, f"{ref_base}.{q_num}", "Responsible Resource Management", "Minimum Expectation",
    "Dedicated Waste Storage Area", "Descriptive",
    "If yes, explain how potential waste quantities from excluded tenancy spaces have been accounted for in the storage area sizing, including the estimation methodology used.",
    "")
q_num += 1

# Safe and Efficient Access to Waste Storage
row = add_criteria_header(ws4, row, "Safe and Efficient Access to Waste Storage")

row = add_question(ws4, row, f"{ref_base}.{q_num}", "Responsible Resource Management", "Minimum Expectation",
    "Safe and Efficient Access to Waste Storage", "Descriptive",
    "Identify the waste specialist and/or waste contractor who signed off on the waste storage designs. Include their name, organisation, and confirm they have a minimum of three years' experience developing operational waste management plans.",
    "Tracks waste specialist involvement in building design.")
q_num += 1

row = add_question(ws4, row, f"{ref_base}.{q_num}", "Responsible Resource Management", "Minimum Expectation",
    "Safe and Efficient Access to Waste Storage", "Descriptive",
    "Describe the key findings of the waste specialist/contractor sign-off, including confirmation that the storage areas are adequately sized and located for the safe and convenient storage and collection of all identified waste streams.",
    "Documents waste management design validation practices.")
q_num += 1

row = add_question(ws4, row, f"{ref_base}.{q_num}", "Responsible Resource Management", "Minimum Expectation",
    "Safe and Efficient Access to Waste Storage", "Data",
    "State the total building GFA, the number of occupants or units served, the total waste storage area provided (m²), and the estimated total annual operational waste (tonnes/year or m³/year).",
    "Key data for developing waste generation benchmarks normalised by building size and occupancy.")
q_num += 1


# ════════════════════════════════════════════════════════════════════════════
# CREDIT 5: RESPONSIBLE PROCUREMENT
# ════════════════════════════════════════════════════════════════════════════
ws5 = wb.create_sheet()
row = setup_sheet(ws5, "Responsible Procurement")
row = add_credit_header(ws5, row, "Responsible Procurement",
    "The procurement process for key products, materials, and services follows best practice environmental and social principles.")
ref_base = "RP"
q_num = 1

# --- Credit Achievement ---
row = add_level_header(ws5, row, "Credit Achievement (1 point)")

# Risk and Opportunity Assessment
row = add_criteria_header(ws5, row, "Risk and Opportunity Assessment")

row = add_question(ws5, row, f"{ref_base}.{q_num}", "Responsible Procurement", "Credit Achievement",
    "Risk and Opportunity Assessment", "Descriptive",
    "Confirm the risk and opportunity assessment was completed prior to appointment of the head contractor for main works. Identify who conducted the assessment (project design team) and confirm input was obtained from the building owner.",
    "Tracks timing and stakeholder involvement in supply chain risk assessment.")
q_num += 1

row = add_question(ws5, row, f"{ref_base}.{q_num}", "Responsible Procurement", "Credit Achievement",
    "Risk and Opportunity Assessment", "Descriptive",
    "List the 10 or more key items identified in the project's supply chain, confirming at least two are building services and at least one is a building material. For each item, briefly explain why it was selected as a key item.",
    "Maps supply chain risk hotspots across building projects for industry-wide risk profiling.")
q_num += 1

row = add_question(ws5, row, f"{ref_base}.{q_num}", "Responsible Procurement", "Credit Achievement",
    "Risk and Opportunity Assessment", "Descriptive",
    "Describe how environmental, social, and human health risks and opportunities were identified and evaluated for each key item, with reference to the following issue areas from ISO 20400:2017 Clause 4.3: (a) human rights, (b) labour practices, (c) the environment, (d) fair operating practices, (e) consumer issues, and (f) community involvement and development.",
    "Captures the depth and breadth of sustainability risk analysis in procurement decisions.")
q_num += 1

row = add_question(ws5, row, f"{ref_base}.{q_num}", "Responsible Procurement", "Credit Achievement",
    "Risk and Opportunity Assessment", "Data",
    "For each of the 10+ key items, summarise the top-priority risks and opportunities identified and the risk rating assigned (e.g. high, medium, low) for each ISO 20400 issue area.",
    "Enables aggregated analysis of supply chain risk profiles to inform industry guidance.")
q_num += 1

row = add_question(ws5, row, f"{ref_base}.{q_num}", "Responsible Procurement", "Credit Achievement",
    "Risk and Opportunity Assessment", "Descriptive",
    "Explain the methodology or framework used to analyse and prioritise the risks and opportunities. Identify any tools, databases, or references used beyond ISO 20400:2017 Annex A.",
    "Documents risk assessment methodologies for knowledge sharing across the industry.")
q_num += 1

# Responsible Procurement Plan
row = add_criteria_header(ws5, row, "Responsible Procurement Plan")

row = add_question(ws5, row, f"{ref_base}.{q_num}", "Responsible Procurement", "Credit Achievement",
    "Responsible Procurement Plan", "Descriptive",
    "Describe the project-level environmental, social, and economic objectives set in the responsible procurement plan to address the risks and implement the opportunities identified in the assessment.",
    "Captures the range and ambition of procurement sustainability objectives across projects.")
q_num += 1

row = add_question(ws5, row, f"{ref_base}.{q_num}", "Responsible Procurement", "Credit Achievement",
    "Responsible Procurement Plan", "Descriptive",
    "Explain the data collection, impact measurement, monitoring, and reporting requirements outlined in the plan, with reference to ISO 20400:2017 Clause 6.5. Describe what metrics are tracked and how frequently.",
    "Documents procurement monitoring and reporting approaches for industry benchmarking.")
q_num += 1

row = add_question(ws5, row, f"{ref_base}.{q_num}", "Responsible Procurement", "Credit Achievement",
    "Responsible Procurement Plan", "Descriptive",
    "Describe the framework established for incentivising contractors and trades to achieve the plan's objectives. Provide examples of incentive mechanisms used.",
    "Tracks effectiveness of supply chain incentive models for sustainability outcomes.")
q_num += 1

row = add_question(ws5, row, f"{ref_base}.{q_num}", "Responsible Procurement", "Credit Achievement",
    "Responsible Procurement Plan", "Descriptive",
    "Explain how the responsible procurement plan was embedded in tender documentation for the head contractor and relevant trades. If the head contractor was engaged under a design and construct contract, describe how the plan was developed prior to procurement activities and embedded in subcontractor tenders.",
    "Measures integration of sustainability requirements into standard procurement workflows.")
q_num += 1

row = add_question(ws5, row, f"{ref_base}.{q_num}", "Responsible Procurement", "Credit Achievement",
    "Responsible Procurement Plan", "Condition (Y/N)",
    "Was the head contractor engaged under a design and construct (D&C) contract? (Yes/No)",
    "Tracks procurement models and their impact on sustainability integration.")
q_num += 1

row = add_question(ws5, row, f"{ref_base}.{q_num}", "Responsible Procurement", "Credit Achievement",
    "Responsible Procurement Plan", "Descriptive",
    "If a D&C contract was used, explain how the head contractor was involved in developing the responsible procurement plan and how the plan was embedded in subcontractor and trade tender documentation.",
    "")
q_num += 1

row = add_question(ws5, row, f"{ref_base}.{q_num}", "Responsible Procurement", "Credit Achievement",
    "Responsible Procurement Plan", "Descriptive",
    "Describe how the plan was implemented during construction in partnership with relevant contractors and trades. Provide examples of data collection, monitoring, and reporting activities carried out, with reference to ISO 20400:2017 Clause 7.",
    "Captures real-world implementation of sustainable procurement practices for effectiveness research.")
q_num += 1

row = add_question(ws5, row, f"{ref_base}.{q_num}", "Responsible Procurement", "Credit Achievement",
    "Responsible Procurement Plan", "Data",
    "List the key items for which responsible procurement actions were taken during construction. For each, summarise the sustainability outcome achieved (e.g. modern slavery risk mitigated, environmental impact reduced, local supply chain used).",
    "Provides outcome-level data on responsible procurement effectiveness for advocacy and policy research.")
q_num += 1

row = add_question(ws5, row, f"{ref_base}.{q_num}", "Responsible Procurement", "Credit Achievement",
    "Responsible Procurement Plan", "Data",
    "Identify any supply chain risks that materialised during construction and the corrective actions taken. State the number of items where procurement plans were fully implemented versus partially implemented.",
    "Tracks procurement plan implementation rates and real-world supply chain risk events.")
q_num += 1


# ════════════════════════════════════════════════════════════════════════════
# Save
# ════════════════════════════════════════════════════════════════════════════
output_path = "/home/user/submissionforms/Green_Star_Responsible_Credits_Submission_Questions.xlsx"
wb.save(output_path)
print(f"Saved to {output_path}")
