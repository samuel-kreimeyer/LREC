#!/usr/bin/env python3
"""
Script to generate Certificate of Attendance PDFs from an Excel spreadsheet.
Run with: source venv/bin/activate && python scripts/generate_coas.py

Requirements: pandas, openpyxl, reportlab (installed in venv).
"""

import os
import pandas as pd
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

def create_certificate_pdf(name, speaker, title, date, output_path):
    """
    Generate a single PDF certificate.
    """
    doc = SimpleDocTemplate(output_path, pagesize=landscape(letter), leftMargin=0.5*inch, rightMargin=0.5*inch, topMargin=0.5*inch, bottomMargin=0.5*inch)
    story = []
    styles = getSampleStyleSheet()

    # Custom style for title: centered, bold, large font
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        spaceAfter=30,
        alignment=TA_CENTER,
        textColor=colors.black
    )

    # Custom style for body text: centered, normal font
    body_style = ParagraphStyle(
        'CustomBody',
        parent=styles['Normal'],
        fontSize=12,
        spaceAfter=12,
        alignment=TA_CENTER,
        textColor=colors.black
    )

    # Custom style for club name: centered, bold
    club_style = ParagraphStyle(
        'CustomClub',
        parent=styles['Normal'],
        fontSize=18,
        spaceAfter=10,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold',
        textColor=colors.black
    )

    # Header with image left and club name
    try:
        img = Image('skyline.png', width=1.5*inch, height=1*inch)
    except:
        img = Paragraph("Image not found", body_style)  # Fallback if image missing

    club_para = Paragraph("LITTLE ROCK ENGINEERS CLUB", club_style)

    header_data = [
        [img, club_para]
    ]

    header_table = Table(header_data, colWidths=[1.5*inch, 8.5*inch])  # Left for image, right wide for centered text
    header_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('ALIGN', (0,0), (0,0), 'LEFT'),
        ('ALIGN', (1,0), (1,0), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 20),
        ('RIGHTPADDING', (0,0), (-1,-1), 20),
        ('TOPPADDING', (0,0), (-1,-1), 10),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
    ]))

    story.append(header_table)
    story.append(Spacer(0.25, 0.3 * inch))

    # Title
    story.append(Paragraph("CERTIFICATE OF ATTENDANCE", title_style))
    story.append(Spacer(1, 0.2 * inch))

    # Certify text
    story.append(Paragraph("This is to certify that", body_style))

    # Name
    name_para = Paragraph(f"<b>{name}</b>", body_style)
    story.append(name_para)

    # Presentation by
    story.append(Paragraph("Earned one (1) Professional Development Hour (PDH) by attending the presentation by:", body_style))

    # Speaker
    speaker_para = Paragraph(f"<b>{speaker}</b>", body_style)
    story.append(speaker_para)

    # Title
    title_para = Paragraph(f"<i>{title}</i>", body_style)
    story.append(title_para)

    story.append(Spacer(1, 2.5 * inch))  # Adjusted spacer to fit on single page

    # Date at bottom
    date_line = f"Conducted in Little Rock, Arkansas on <b>{date}</b>"
    story.append(Paragraph(date_line, body_style))

    # Build PDF
    doc.build(story)

def main():
    # Prompt for spreadsheet path
    spreadsheet_path = input("Enter the path to the spreadsheet (default: COA forms.xlsx): ").strip()
    if not spreadsheet_path:
        spreadsheet_path = "COA forms.xlsx"

    if not os.path.exists(spreadsheet_path):
        print(f"Error: Spreadsheet '{spreadsheet_path}' not found.")
        return

    try:
        # Read Excel
        df = pd.read_excel(spreadsheet_path)
        print(f"Loaded spreadsheet with columns: {list(df.columns)}")
        print(f"Number of rows: {len(df)}")

        required_cols = ['Name', 'Speaker', 'Title', 'Date']
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            print(f"Error: Missing required columns: {missing_cols}")
            return

        # Generate PDF for each row
        for index, row in df.iterrows():
            name = row['Name']
            speaker = row['Speaker']
            title = row['Title']
            date = row['Date']

            # Sanitize for filename
            safe_name = "".join(c for c in str(name) if c.isalnum() or c in (' ', '-', '_')).rstrip()
            safe_date = str(date).split()[0].replace('/', '_').replace('-', '_')
            output_filename = f"COA_{safe_name}_{safe_date}.pdf"
            output_path = os.path.join(os.getcwd(), output_filename)

            print(f"Generating {output_filename}...")
            create_certificate_pdf(name, speaker, title, date, output_path)

        print("All certificates generated successfully!")

    except Exception as e:
        print(f"Error processing spreadsheet: {e}")

if __name__ == "__main__":
    main()