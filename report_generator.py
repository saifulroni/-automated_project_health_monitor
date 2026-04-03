from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import HexColor, white
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table,
    TableStyle, HRFlowable
)
from reportlab.lib.units import mm
from reportlab.lib import colors
from datetime import datetime
import pandas as pd
from rag_engine import run_rag_engine

RAG_HEX = {
    'Red': '#FF4444',
    'Amber': '#FFA500',
    'Green': '#44BB44',
    'Grey': '#AAAAAA'
}

def build_pdf_report(tasks_filepath, output_path):
    df, project_rags, critical_path = run_rag_engine(tasks_filepath)
    today = pd.Timestamp(datetime.today().date())

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=15 * mm,
        leftMargin=15 * mm,
        topMargin=15 * mm,
        bottomMargin=15 * mm
    )

    styles = getSampleStyleSheet()
    story = []

    # Styles
    title_style = ParagraphStyle(
        'Title',
        fontSize=18,
        fontName='Helvetica-Bold',
        textColor=HexColor('#1F3864'),
        spaceAfter=4
    )

    sub_style = ParagraphStyle(
        'Sub',
        fontSize=10,
        fontName='Helvetica',
        textColor=HexColor('#555555'),
        spaceAfter=12
    )

    section_style = ParagraphStyle(
        'Section',
        fontSize=12,
        fontName='Helvetica-Bold',
        textColor=HexColor('#1F3864'),
        spaceBefore=10,
        spaceAfter=6
    )

    wrap_style = ParagraphStyle(
        'Wrap',
        fontSize=8,
        fontName='Helvetica',
        leading=10,
        wordWrap='LTR'
    )

    footer_style = ParagraphStyle(
        'Footer',
        fontSize=7,
        textColor=HexColor('#888888'),
        alignment=1
    )

    # Title
    story.append(Paragraph("LEMU-TAJIN LAUNCH — WEEKLY PMO HEALTH REPORT", title_style))
    story.append(Paragraph(
        f"Week of {today.strftime('%d %B %Y')}  |  Project Health Monitor",
        sub_style
    ))
    story.append(HRFlowable(width="100%", thickness=2, color=HexColor('#1F3864')))
    story.append(Spacer(1, 8))

    
    story.append(Paragraph("1. PORTFOLIO HEALTH OVERVIEW", section_style))

    port_data = [['Project', 'RAG Status', 'Total Tasks', '% Complete', 'Overdue']]

    for project, rag in project_rags.items():
        proj_df = df[df['Project_Name'] == project]
        overdue = len(proj_df[(proj_df['RAG'] == 'Red') & (proj_df['Status'] != 'Completed')])

        port_data.append([
            Paragraph(project, wrap_style),
            rag,
            str(len(proj_df)),
            f"{round(proj_df['Pct_Complete'].mean(), 1)}%",
            str(overdue)
        ])

    port_table = Table(port_data, colWidths=[65 * mm, 30 * mm, 25 * mm, 28 * mm, 22 * mm])

    port_style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), HexColor('#1F3864')),
        ('TEXTCOLOR', (0, 0), (-1, 0), white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [HexColor('#F0F4FF'), white]),
    ])

    for row_idx, (_, rag) in enumerate(project_rags.items(), 1):
        color = HexColor(RAG_HEX.get(rag, '#AAAAAA'))
        port_style.add('BACKGROUND', (1, row_idx), (1, row_idx), color)
        port_style.add('TEXTCOLOR', (1, row_idx), (1, row_idx), white)

    port_table.setStyle(port_style)
    story.append(port_table)
    story.append(Spacer(1, 10))

    
    story.append(Paragraph("2. OVERDUE & BLOCKED TASKS (ALL PROJECTS)", section_style))

    overdue_df = df[(df['RAG'] == 'Red') & (df['Status'] != 'Completed')]

    if len(overdue_df) > 0:
        overdue_data = [['Task ID', 'Task Name', 'Project', 'Owner', 'Days Late', 'Status', 'Last Comment']]

        for _, row in overdue_df.iterrows():
            days_late = max((today - row['Planned_End_Date']).days, 0)

            overdue_data.append([
                row['Task_ID'],
                Paragraph(str(row['Task_Name']), wrap_style),
                Paragraph(str(row['Project_Name']), wrap_style),
                row['Owner'],
                f"+{days_late}d",
                row['Status'],
                Paragraph(str(row['Comments']) if pd.notna(row['Comments']) else '', wrap_style)
            ])

        od_table = Table(
            overdue_data,
            colWidths=[15 * mm, 42 * mm, 30 * mm, 18 * mm, 15 * mm, 20 * mm, 45 * mm]
        )

        od_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), HexColor('#CC0000')),
            ('TEXTCOLOR', (0, 0), (-1, 0), white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ]))

        story.append(od_table)
    else:
        story.append(Paragraph("✓ No overdue tasks this week.", styles['Normal']))

    story.append(Spacer(1, 10))

    
        # SECTION 3
    story.append(Paragraph("3. TASKS DUE IN NEXT 7 DAYS — AMBER OR RED", section_style))

    upcoming = df[
        (df['Planned_End_Date'] >= today) &
        (df['Planned_End_Date'] <= today + pd.Timedelta(days=7)) &
        (df['RAG'].isin(['Amber', 'Red'])) &
        (df['Status'] != 'Completed')
    ]

    if len(upcoming) > 0:
        up_data = [['Task ID', 'Task Name', 'Project', 'Owner', 'Due Date', 'RAG', '% Done']]

        for _, row in upcoming.iterrows():
            up_data.append([
                row['Task_ID'],
                Paragraph(str(row['Task_Name']), wrap_style),
                Paragraph(str(row['Project_Name']), wrap_style),
                row['Owner'],
                row['Planned_End_Date'].strftime('%d %b'),
                row['RAG'],
                f"{int(row['Pct_Complete'])}%"
            ])

        up_table = Table(
            up_data,
            colWidths=[15 * mm, 42 * mm, 32 * mm, 20 * mm, 18 * mm, 16 * mm, 16 * mm]
        )

        up_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), HexColor('#1F3864')),
            ('TEXTCOLOR', (0, 0), (-1, 0), white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('ALIGN', (0, 1), (0, -1), 'CENTER'),
            ('ALIGN', (3, 1), (6, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [HexColor('#F8FAFD'), white]),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ])

        for row_idx in range(1, len(up_data)):
            rag = up_data[row_idx][5]
            up_style.add('BACKGROUND', (5, row_idx), (5, row_idx), HexColor(RAG_HEX.get(rag, '#AAAAAA')))
            up_style.add('TEXTCOLOR', (5, row_idx), (5, row_idx), white)
            up_style.add('FONTNAME', (5, row_idx), (5, row_idx), 'Helvetica-Bold')

        up_table.setStyle(up_style)
        story.append(up_table)
    else:
        story.append(Paragraph("✓ No at-risk tasks due in the next 7 days.", styles['Normal']))

    story.append(Spacer(1, 10))
    
        # SECTION 4
    story.append(Paragraph("4. CRITICAL PATH STATUS", section_style))

    cp_df = df[df['Is_Critical'] == True]
    cp_data = [['Task ID', 'Task Name', 'Owner', 'End Date', 'RAG', '% Done']]

    for _, row in cp_df.iterrows():
        cp_data.append([
            row['Task_ID'],
            Paragraph(str(row['Task_Name']), wrap_style),
            row['Owner'],
            row['Planned_End_Date'].strftime('%d %b') if pd.notna(row['Planned_End_Date']) else '',
            row['RAG'],
            f"{int(row['Pct_Complete'])}%"
        ])

    cp_table = Table(
        cp_data,
        colWidths=[15 * mm, 62 * mm, 22 * mm, 20 * mm, 16 * mm, 18 * mm]
    )

    cp_style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), HexColor('#1F3864')),
        ('TEXTCOLOR', (0, 0), (-1, 0), white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('ALIGN', (0, 1), (0, -1), 'CENTER'),
        ('ALIGN', (2, 1), (5, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [HexColor('#F8FAFD'), white]),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ])

    for row_idx in range(1, len(cp_data)):
        rag = cp_data[row_idx][4]
        cp_style.add('BACKGROUND', (4, row_idx), (4, row_idx), HexColor(RAG_HEX.get(rag, '#AAAAAA')))
        cp_style.add('TEXTCOLOR', (4, row_idx), (4, row_idx), white)
        cp_style.add('FONTNAME', (4, row_idx), (4, row_idx), 'Helvetica-Bold')

    cp_table.setStyle(cp_style)
    story.append(cp_table)

    # Footer
    story.append(Spacer(1, 15))
    story.append(HRFlowable(width="100%", thickness=1, color=HexColor('#AAAAAA')))
    story.append(Paragraph(
        f"Lemu Tajin Launch PMO | Project Monitoring Report {today.strftime('%d %B %Y')}",
        footer_style
    ))

    doc.build(story)
    print(f"PDF saved: {output_path}")


if __name__ == '__main__':
    build_pdf_report('Tasks.xlsx', f"outputs/lemu-tajin_report_{datetime.today().strftime('%Y%m%d')}.pdf")