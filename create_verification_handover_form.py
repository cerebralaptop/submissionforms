#!/usr/bin/env python3
"""
Script to create Verification and Handover credit submission form for Green Star Buildings v1.1
"""

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

def create_submission_form():
    wb = Workbook()
    ws = wb.active
    ws.title = "Verification & Handover"

    # Define styles
    header_font = Font(bold=True, size=12, color="FFFFFF")
    header_fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
    section_font = Font(bold=True, size=11)
    section_fill = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")
    subsection_fill = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    wrap_alignment = Alignment(wrap_text=True, vertical='top')

    # Set column widths
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 50
    ws.column_dimensions['D'].width = 15
    ws.column_dimensions['E'].width = 40
    ws.column_dimensions['F'].width = 50

    # Title
    ws.merge_cells('A1:F1')
    ws['A1'] = "GREEN STAR BUILDINGS v1.1 - VERIFICATION AND HANDOVER CREDIT SUBMISSION FORM"
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].alignment = Alignment(horizontal='center')

    # Credit information
    ws.merge_cells('A2:F2')
    ws['A2'] = "Outcome: The building has been optimised and handed over to deliver a higher level of performance in operation."
    ws['A2'].font = Font(italic=True, size=10)
    ws['A2'].alignment = wrap_alignment

    # Headers
    row = 4
    headers = ['Criteria', 'Requirement', 'Submission Question', 'Compliance (Y/N/NA)', 'Evidence Provided', 'Documentation Required']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Data structure for the form
    form_data = [
        # MINIMUM EXPECTATION SECTION
        {
            'criteria': 'MINIMUM EXPECTATION',
            'requirement': '',
            'question': '',
            'evidence': '',
            'documentation': '',
            'is_section': True
        },
        # Metering and Monitoring
        {
            'criteria': 'Metering and Monitoring',
            'requirement': 'Metering Distribution',
            'question': 'Does the building have accessible energy and water metering for all distinct uses and major uses?',
            'evidence': '',
            'documentation': 'As-built drawings showing location of all energy and water meters',
            'is_subsection': True
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Is metering provided for each separate tenancy or unit?',
            'evidence': '',
            'documentation': ''
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Is floor by floor metering provided where entire floors have a single use?',
            'evidence': '',
            'documentation': ''
        },
        {
            'criteria': '',
            'requirement': 'Metering Schedule',
            'question': 'Has a metering schedule been provided in accordance with CIBSE TM39?',
            'evidence': '',
            'documentation': 'Metering schedule'
        },
        {
            'criteria': '',
            'requirement': 'Metering Attributes',
            'question': 'Do all meters provide continual information (up to 1-hour interval readings)?',
            'evidence': '',
            'documentation': 'Product data sheets demonstrating pattern approval by NMI or recognised standard'
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Are meters commissioned and validated per NABERS Metering and Consumption Rules?',
            'evidence': '',
            'documentation': 'Completed metering validation documents'
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Are meters pattern approved by NMI or recognised standard?',
            'evidence': '',
            'documentation': 'Product data sheets or certificates'
        },
        {
            'criteria': '',
            'requirement': 'Automatic Monitoring',
            'question': 'Is an automatic monitoring system implemented that reports consumption trends?',
            'evidence': '',
            'documentation': 'Letter of confirmation from contractor/metering provider'
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Does the system raise alarms when energy/water use exceeds parameters?',
            'evidence': '',
            'documentation': 'Extracts from commissioning reports'
        },
        # Commissioning and Tuning
        {
            'criteria': 'Commissioning and Tuning',
            'requirement': 'Environmental Performance Targets',
            'question': 'Were environmental performance targets documented prior to schematic design?',
            'evidence': '',
            'documentation': 'Extracts from design intent report or OPR',
            'is_subsection': True
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Does the report include targets for energy use, water consumption, IEQ and airtightness?',
            'evidence': '',
            'documentation': ''
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Has the building owner signed off on the report/document?',
            'evidence': '',
            'documentation': 'Signed design intent report or OPR'
        },
        {
            'criteria': '',
            'requirement': 'Services and Maintainability Review',
            'question': 'Was a services and maintainability review conducted prior to construction?',
            'evidence': '',
            'documentation': 'Evidence of review being conducted'
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Did the review involve building owner, design consultants, architect, FM, head contractor and ICA?',
            'evidence': '',
            'documentation': ''
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Were all items addressed, closed-out and documented in Services and Maintainability Report?',
            'evidence': '',
            'documentation': 'Services and Maintainability Report signed by all parties'
        },
        {
            'criteria': '',
            'requirement': 'Building Commissioning',
            'question': 'Were commissioning requirements included in construction documentation prior to construction?',
            'evidence': '',
            'documentation': 'Extracts from construction documentation'
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Was a commissioning plan developed prior to start of commissioning?',
            'evidence': '',
            'documentation': 'Extracts of commissioning plan'
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Was commissioning completed per a recognised standard (AIRAH DA27, ASHRAE, CIBSE, SA TA 5342)?',
            'evidence': '',
            'documentation': 'Extracts from commissioning report'
        },
        {
            'criteria': '',
            'requirement': 'Airtightness',
            'question': 'Were airtightness targets set prior to schematic design based on ATTMA Australia Guide?',
            'evidence': '',
            'documentation': 'Evidence of how airtightness targets were determined'
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Was the air barrier system schematic reviewed prior to end of design development?',
            'evidence': '',
            'documentation': 'Evidence of air barrier system schematic review (marked up drawings)'
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Was an approach to delivering airtightness developed prior to construction?',
            'evidence': '',
            'documentation': 'Airtightness testing and commissioning plan'
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Was airtightness testing carried out per AS/NZS ISO 9972 Method 1 by ATTMA member?',
            'evidence': '',
            'documentation': 'Building airtightness testing report or ATTMA Green Star Commissioning report'
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Were test results shared with building owner (regardless of targets being met)?',
            'evidence': '',
            'documentation': 'Signed confirmation from building owner'
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Were airtightness test results included in energy modelling?',
            'evidence': '',
            'documentation': 'Extracts from energy modelling report'
        },
        {
            'criteria': '',
            'requirement': 'Building Systems Tuning',
            'question': 'Has building owner/developer committed to tuning process with quarterly adjustments for 12 months?',
            'evidence': '',
            'documentation': 'Commitment/contract from building owner and head contractor'
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Does commitment include a building tuning manual/plan with roles and responsibilities?',
            'evidence': '',
            'documentation': 'Extracts of building tuning plan'
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Does the building tuning team include FM, ICA/owner rep, head contractor and relevant subcontractors?',
            'evidence': '',
            'documentation': ''
        },
        # Building Information
        {
            'criteria': 'Building Information',
            'requirement': 'Operations and Maintenance Information',
            'question': 'Has O&M information been provided to building owner for all nominated building systems?',
            'evidence': '',
            'documentation': 'Extracts of operations and maintenance information',
            'is_subsection': True
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Does O&M info include maintenance instructions, procedures, schedules, service contacts, warranties and as-built drawings?',
            'evidence': '',
            'documentation': ''
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Do appropriate user groups have access to information for best practice environmental outcomes?',
            'evidence': '',
            'documentation': ''
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Does O&M info include guidance on keeping information up to date and validating alerts/faults?',
            'evidence': '',
            'documentation': ''
        },
        {
            'criteria': '',
            'requirement': 'Building Logbook',
            'question': 'Has a building logbook been developed per CIBSE TM31 for all nominated building systems?',
            'evidence': '',
            'documentation': 'Building logbook'
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Was the logbook presented to building owner prior to occupation?',
            'evidence': '',
            'documentation': 'Evidence of provision to building owner'
        },
        {
            'criteria': '',
            'requirement': 'Building User Information',
            'question': 'Has building user information been provided to building owner and FM prior to occupation?',
            'evidence': '',
            'documentation': 'Building user information'
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Is the building user information publicly available to intended users?',
            'evidence': '',
            'documentation': ''
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Is the information communicated in a way easily understood by the target audience?',
            'evidence': '',
            'documentation': ''
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Is the information in an editable, digital format accessible for updates by FM?',
            'evidence': '',
            'documentation': 'Evidence of editable format provision'
        },
        # CREDIT ACHIEVEMENT SECTION
        {
            'criteria': 'CREDIT ACHIEVEMENT (1 point)',
            'requirement': '',
            'question': '',
            'evidence': '',
            'documentation': '',
            'is_section': True
        },
        {
            'criteria': '',
            'requirement': 'Pathway Selection',
            'question': 'Which pathway is being pursued? (Select one or both if Total Building Services Value > $20m)',
            'evidence': '',
            'documentation': '',
            'is_subsection': True
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Total Building Services Value (if applicable): $_________',
            'evidence': '',
            'documentation': ''
        },
        {
            'criteria': '',
            'requirement': '',
            'question': '[ ] Soft Landings Approach',
            'evidence': '',
            'documentation': ''
        },
        {
            'criteria': '',
            'requirement': '',
            'question': '[ ] Independent Commissioning Agent',
            'evidence': '',
            'documentation': ''
        },
        {
            'criteria': '',
            'requirement': '',
            'question': '[ ] Both (required if Total Building Services Value > $20m)',
            'evidence': '',
            'documentation': ''
        },
        # Soft Landings Approach
        {
            'criteria': 'Soft Landings Approach',
            'requirement': 'CIBSE ANZ Framework Implementation',
            'question': 'Have Stages 1 to 4 of the CIBSE ANZ Soft Landings Framework been implemented?',
            'evidence': '',
            'documentation': 'Evidence of CIBSE ANZ framework implementation',
            'is_subsection': True
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Are sample worksheets from CIBSE ANZ for Stages 1-3 completed?',
            'evidence': '',
            'documentation': 'Completed worksheets'
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Are actions for Stage 4 identified?',
            'evidence': '',
            'documentation': ''
        },
        {
            'criteria': '',
            'requirement': 'FM Team Involvement',
            'question': 'Was the FM team (or building owner rep) involved in commissioning and handover process?',
            'evidence': '',
            'documentation': ''
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Did FM team take part in developing brief technical guide and O&M manual?',
            'evidence': '',
            'documentation': ''
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Did FM team sign-off on the O&M manual?',
            'evidence': '',
            'documentation': 'Signed O&M manual'
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Was FM team trained before handover (including BMS demonstration)?',
            'evidence': '',
            'documentation': 'Training records'
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Does FM team have continued access to design/construction team for 2 years post PC?',
            'evidence': '',
            'documentation': 'Contractual evidence of continued access'
        },
        # Independent Commissioning Agent
        {
            'criteria': 'Independent Commissioning Agent',
            'requirement': 'ICA Appointment',
            'question': 'Was an ICA appointed prior to design development?',
            'evidence': '',
            'documentation': 'Letter from building owner confirming ICA appointment',
            'is_subsection': True
        },
        {
            'criteria': '',
            'requirement': 'ICA Qualifications',
            'question': 'Is the ICA an advocate for and reporting directly to the project owner?',
            'evidence': '',
            'documentation': 'CV of ICA with qualifications and experience'
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Is the ICA independent of any consultant/contractor involved in design or installation?',
            'evidence': '',
            'documentation': ''
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Is the ICA a registered professional engineer or qualified technician with demonstrated commissioning competency?',
            'evidence': '',
            'documentation': ''
        },
        {
            'criteria': '',
            'requirement': '',
            'question': 'Does the ICA have experience commissioning at least 2 similar projects?',
            'evidence': '',
            'documentation': ''
        },
        {
            'criteria': '',
            'requirement': 'ICA Involvement',
            'question': 'Did ICA advise, monitor and verify commissioning/tuning throughout design development, tender, construction, commissioning and tuning phases?',
            'evidence': '',
            'documentation': 'Evidence of ICA involvement from design stage through to tuning'
        },
    ]

    # Write data to worksheet
    row = 5
    for item in form_data:
        is_section = item.get('is_section', False)
        is_subsection = item.get('is_subsection', False)

        if is_section:
            ws.merge_cells(f'A{row}:F{row}')
            cell = ws.cell(row=row, column=1, value=item['criteria'])
            cell.font = section_font
            cell.fill = section_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='left', vertical='center')
        else:
            ws.cell(row=row, column=1, value=item['criteria']).alignment = wrap_alignment
            ws.cell(row=row, column=2, value=item['requirement']).alignment = wrap_alignment
            ws.cell(row=row, column=3, value=item['question']).alignment = wrap_alignment
            ws.cell(row=row, column=4, value='').alignment = Alignment(horizontal='center', vertical='top')
            ws.cell(row=row, column=5, value=item['evidence']).alignment = wrap_alignment
            ws.cell(row=row, column=6, value=item['documentation']).alignment = wrap_alignment

            for col in range(1, 7):
                cell = ws.cell(row=row, column=col)
                cell.border = thin_border
                if is_subsection:
                    cell.fill = subsection_fill
                    if col == 1:
                        cell.font = Font(bold=True)

        row += 1

    # Add notes section
    row += 1
    ws.merge_cells(f'A{row}:F{row}')
    ws.cell(row=row, column=1, value="NOTES:").font = Font(bold=True)
    row += 1
    notes = [
        "1. Complete all applicable sections based on the performance level being targeted.",
        "2. For Minimum Expectation: All three criteria (Metering and Monitoring, Commissioning and Tuning, Building Information) must be met.",
        "3. For Credit Achievement (1 point): In addition to Minimum Expectation, either Soft Landings Approach OR Independent Commissioning Agent must be met.",
        "4. For projects with Total Building Services Value > $20m: BOTH Soft Landings Approach AND Independent Commissioning Agent must be met for Credit Achievement.",
        "5. Attach all supporting documentation as referenced in the 'Documentation Required' column.",
        "6. Mark Y (Yes), N (No), or NA (Not Applicable) in the Compliance column.",
    ]
    for note in notes:
        ws.merge_cells(f'A{row}:F{row}')
        ws.cell(row=row, column=1, value=note).alignment = wrap_alignment
        row += 1

    # Set row heights for better readability
    for r in range(5, row):
        ws.row_dimensions[r].height = 30

    # Save workbook
    filename = '/home/user/submissionforms/Verification_and_Handover_Submission_Form.xlsx'
    wb.save(filename)
    print(f"Created: {filename}")
    return filename

if __name__ == "__main__":
    create_submission_form()
