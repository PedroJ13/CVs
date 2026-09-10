from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


OUTPUT_PATH = Path(r"C:\Work\CVs\Output\Pedro_Gutierrez_DistantJob_Senior_SQL_DBA_Cover_Letter.pdf")


def build_pdf():
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)

    document = SimpleDocTemplate(
        str(OUTPUT_PATH),
        pagesize=LETTER,
        rightMargin=0.75 * inch,
        leftMargin=0.75 * inch,
        topMargin=0.62 * inch,
        bottomMargin=0.62 * inch,
        pageCompression=1,
        title="Pedro Gutierrez Cover Letter Senior SQL DBA",
        author="Pedro Javier Gutierrez Armas",
    )

    styles = getSampleStyleSheet()
    name_style = ParagraphStyle(
        "Name",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=21,
        textColor=colors.HexColor("#17365D"),
        alignment=TA_CENTER,
        spaceAfter=3,
    )
    contact_style = ParagraphStyle(
        "Contact",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.8,
        leading=11,
        textColor=colors.HexColor("#404040"),
        alignment=TA_CENTER,
    )
    date_style = ParagraphStyle(
        "Date",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10.2,
        leading=13,
        alignment=TA_LEFT,
        spaceAfter=12,
    )
    body_style = ParagraphStyle(
        "Body",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10.2,
        leading=14.2,
        textColor=colors.black,
        alignment=TA_LEFT,
        spaceAfter=10,
    )
    subject_style = ParagraphStyle(
        "Subject",
        parent=body_style,
        fontName="Helvetica-Bold",
        spaceAfter=12,
    )

    story = [
        Paragraph("PEDRO JAVIER GUTIERREZ ARMAS", name_style),
        Paragraph(
            "San Jose, Costa Rica | +506 6351-5860 | pj13eros@hotmail.com | linkedin.com/in/pedrogutierrez13",
            contact_style,
        ),
        Spacer(1, 8),
        Table(
            [[""]],
            colWidths=[7.0 * inch],
            rowHeights=[1.2],
            style=TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#2F75B5"))]),
        ),
        Spacer(1, 18),
        Paragraph("September 10, 2026", date_style),
        Paragraph("Dear Hiring Team,", body_style),
        Paragraph("Re: Senior SQL DBA", subject_style),
        Paragraph(
            "I am writing to express my interest in the Senior SQL DBA position with DistantJob's client. I bring more than 15 years of experience working with Microsoft SQL Server across database administration, SQL development, production support, performance optimization, data integration, and enterprise reporting environments.",
            body_style,
        ),
        Paragraph(
            "Throughout my career, I have supported production databases, developed and optimized complex T-SQL queries and stored procedures, improved indexing strategies, and resolved database and data workflow issues. At Intertec International, my optimization work reduced query execution times by up to 50%. I have also worked with SQL Server Always On, Extended Events, and Query Store to analyze workload behavior and troubleshoot performance problems.",
            body_style,
        ),
        Paragraph(
            "My experience includes leading the migration of ten SQL Server databases, maintaining data integrity during production changes, and collaborating with development and business teams to identify risks and deliver reliable solutions. I have also established database management and security practices, documented technical environments, and trained development teams on SQL programming and performance best practices.",
            body_style,
        ),
        Paragraph(
            "In my current role, I continue developing and reviewing SQL-based data solutions while using Git-based workflows, Cursor, and an AI-assisted pull request review process. This experience has strengthened my ability to evaluate database changes for maintainability, performance, and downstream impact.",
            body_style,
        ),
        Paragraph(
            "I am particularly interested in this opportunity because it combines production ownership, proactive performance improvement, and close collaboration with developers. I would welcome the opportunity to discuss how my SQL Server background could support your client's customized, multi-database SaaS environment.",
            body_style,
        ),
        Spacer(1, 4),
        Paragraph("Sincerely,", body_style),
        Paragraph("<b>Pedro Javier Gutierrez Armas</b>", body_style),
    ]

    document.build(story)
    print(OUTPUT_PATH)


if __name__ == "__main__":
    build_pdf()
