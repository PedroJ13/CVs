from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(r"C:\Work\CVs")
SOURCE = ROOT / "Base" / "Pedro_Gutierrez_CV_SQL_Server_DBA.docx"
OUTPUT = ROOT / "Output" / "Pedro_Gutierrez_FullStack_Principal_SQL_DBA_CV.docx"

BLUE = "1F4E79"
DARK = "222222"
GRAY = "5B5B5B"
LIGHT_LINE = "BDD7EE"


def clear_body(doc):
    body = doc._element.body
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)


def write(paragraph, text, size=9.1, bold=False, italic=False, color=None, font="Aptos"):
    run = paragraph.add_run(text)
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    return run


def set_bottom_border(paragraph):
    p_pr = paragraph._p.get_or_add_pPr()
    borders = p_pr.find(qn("w:pBdr"))
    if borders is None:
        borders = OxmlElement("w:pBdr")
        p_pr.append(borders)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "4")
    bottom.set(qn("w:space"), "2")
    bottom.set(qn("w:color"), LIGHT_LINE)
    borders.append(bottom)


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_borders(cell, color="D9D9D9", size="4"):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = qn(f"w:{edge}")
        node = borders.find(tag)
        if node is None:
            node = OxmlElement(f"w:{edge}")
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), size)
        node.set(qn("w:color"), color)


def set_cell_margins(cell, top=65, start=80, bottom=65, end=80):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for side, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def section_heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(7)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    write(p, text.upper(), 10.2, True, color=BLUE, font="Aptos Display")
    set_bottom_border(p)


def bullet(doc, text, size=9.0):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Inches(0.18)
    p.paragraph_format.first_line_indent = Inches(-0.12)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = 1.01
    p.paragraph_format.keep_together = True
    write(p, text, size=size)


def role(doc, title, company, dates, place, points, technologies, page_break_before=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4.5)
    p.paragraph_format.space_after = Pt(0.5)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.page_break_before = page_break_before
    write(p, f"{title} | {company}", 10, True, color=DARK)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(1.5)
    p.paragraph_format.keep_with_next = True
    write(p, f"{dates} | {place}", 8.5, italic=True, color=GRAY)

    for point in points:
        bullet(doc, point)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(1.5)
    p.paragraph_format.keep_together = True
    write(p, "Technologies: ", 8.5, True, color=DARK)
    write(p, technologies, 8.5, color=GRAY)


def add_skills_table(doc):
    rows = [
        (
            "SQL Server operations",
            "AWS RDS, production, development and staging environments, SSMS, DBeaver, SQL Agent, Database Mail, maintenance, backup and restore, monitoring and incident support",
        ),
        (
            "Performance engineering",
            "Advanced T-SQL, stored procedures, functions, indexing, execution-pattern analysis, long-running queries, set-based optimization and bottleneck troubleshooting",
        ),
        (
            "Security and standards",
            "RBAC, least privilege, JIT privileged access, login and user auditing, schema permissions, SID mapping, credential rotation, runbooks and operational procedures",
        ),
        (
            "Platforms and integration",
            "SQL Server, AWS RDS, SSIS, MySQL, PostgreSQL, Snowflake, Azure-based environments, ETL/ELT, data warehousing and Power BI",
        ),
        (
            "Automation and delivery",
            "PowerShell, Python, Git, controlled deployment and rollback scripts, diagnostic utilities, Agile collaboration, Cursor and AI-assisted pull-request review",
        ),
    ]
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    widths = [Inches(1.72), Inches(5.68)]
    for row in table.rows:
        for index, cell in enumerate(row.cells):
            cell.width = widths[index]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    for index, text in enumerate(("Area", "Tools and practices")):
        cell = table.rows[0].cells[index]
        cell.text = ""
        set_cell_shading(cell, BLUE)
        set_cell_borders(cell)
        set_cell_margins(cell)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        write(p, text, 8.7, True, color="FFFFFF")

    for row_index, (area, details) in enumerate(rows, start=1):
        cells = table.add_row().cells
        for index, value in enumerate((area, details)):
            cell = cells[index]
            cell.width = widths[index]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cell.text = ""
            set_cell_borders(cell)
            set_cell_margins(cell)
            if row_index % 2 == 0:
                set_cell_shading(cell, "F3F7FA")
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            write(p, value, 8.35, bold=(index == 0))


def build_document():
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc = Document(SOURCE)
    clear_body(doc)

    section = doc.sections[0]
    section.top_margin = Inches(0.43)
    section.bottom_margin = Inches(0.43)
    section.left_margin = Inches(0.55)
    section.right_margin = Inches(0.55)

    normal = doc.styles["Normal"]
    normal.font.name = "Aptos"
    normal.font.size = Pt(9.1)
    normal.paragraph_format.space_after = Pt(2)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0.5)
    write(p, "PEDRO JAVIER GUTIERREZ ARMAS", 18, True, color=BLUE, font="Aptos Display")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(1)
    write(p, "Principal SQL Database Administrator | Production Support and Performance Engineering", 9.7, True, color="404040")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    write(p, "San Jose, Costa Rica | +506 6351-5860 | pj13eros@hotmail.com | linkedin.com/in/pedrogutierrez13", 8.5, color="4C4C4C")

    section_heading(doc, "Professional Summary")
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.04
    p.paragraph_format.space_after = Pt(4)
    write(
        p,
        "Principal-level SQL Server DBA, database engineer, and SQL developer with 15+ years of experience administering, securing, developing, optimizing, and supporting enterprise databases and data platforms. Current hands-on experience supporting SQL Server on AWS RDS across production and non-production environments, with strong expertise in advanced T-SQL, indexing, execution analysis, backup and restore, SQL Agent monitoring, database security, production incident response, and controlled deployments. Proven record of improving query execution time by up to 50%, establishing database access and operating standards, leading data recovery and migration work, documenting repeatable procedures, and partnering with developers and stakeholders to resolve complex database issues.",
    )

    section_heading(doc, "Technical Skills")
    add_skills_table(doc)

    section_heading(doc, "Professional Experience")
    role(
        doc,
        "SQL Server DBA / Database Engineer",
        "MWR Life",
        "September 2024 - Present",
        "Remote / Costa Rica",
        [
            "Administer and support SQL Server databases hosted on AWS RDS across Production, Development, and Staging, covering access, connectivity, backup and restore, maintenance, monitoring, and production incident response.",
            "Designed a role-based access model using least-privilege principles, standardized account conventions, JIT privileged access, permission audits, and controlled post-restore procedures for lower environments.",
            "Develop T-SQL stored procedures, functions, deployment and rollback scripts, diagnostic procedures, and support utilities while resolving authentication, SID mapping, SSL, firewall, and application connectivity issues.",
            "Led a controlled production data recovery and integration remediation effort, validating source backups, recovering hundreds of records, correcting stored procedures, and completing post-deployment verification.",
            "Analyze long-running queries, indexes, execution patterns, query variants, and cursor-based processes; design set-based alternatives and validate output equivalence before release.",
            "Built monitoring dashboards and runbooks for SQL Agent jobs, alerts, execution history, integrations, data-quality exceptions, restore procedures, access controls, and incident response.",
        ],
        "SQL Server, AWS RDS, T-SQL, SSMS, DBeaver, SQL Agent, Database Mail, PowerShell, JSON/OpenJSON, Git, RBAC, JIT access, backup/restore, monitoring and performance tuning",
    )

    role(
        doc,
        "DBA / SQL Developer",
        "Health Catalyst",
        "May 2021 - September 2024",
        "San Jose, Costa Rica",
        [
            "Developed and supported SQL Server, SSIS, Snowflake, Python, and PowerShell data workflows for production and client-facing requirements.",
            "Designed customized ETL applications that reduced manual processing time by up to 40%.",
            "Automated validation and anomaly-detection routines, improving data-validation accuracy by 30% and supporting reliable reporting operations.",
        ],
        "SQL Server, T-SQL, SSIS, Snowflake, Python, PowerShell, Power BI, Pandas and Excel",
        page_break_before=True,
    )

    role(
        doc,
        "SQL Developer",
        "Intertec International",
        "February 2019 - April 2021",
        "San Jose, Costa Rica",
        [
            "Developed and optimized advanced T-SQL queries, stored procedures, and database workflows supporting production business logic and analytics.",
            "Reduced query execution times by up to 50% through SQL tuning, indexing strategies, and performance analysis.",
            "Integrated SQL Server, Salesforce, and MySQL data while improving data accuracy by more than 25% through validation and error handling.",
        ],
        "SQL Server, T-SQL, SSIS, Salesforce and MySQL",
    )

    role(
        doc,
        "DBA / SQL Developer",
        "EL Tiempo",
        "September 2020 - November 2020",
        "Colombia",
        [
            "Led planning and execution of migration activities for ten SQL Server databases, delivering all phases on time and within scope.",
            "Coordinated migration risk analysis, data validation, integrity controls, deployment timing, and operational continuity with cross-functional teams.",
        ],
        "SQL Server, SSIS, Azure and SSRS",
    )

    role(
        doc,
        "DBA / SQL and BI Consultant",
        "Gold Data Networks",
        "January 2016 - February 2020",
        "Panama City, Panama",
        [
            "Designed relational databases, schemas, tables, stored procedures, and supporting objects for operational applications and reporting workloads.",
            "Delivered database, integration, and BI solutions aligned with application performance and infrastructure requirements.",
        ],
        "SQL Server, SSIS, SSRS, PostgreSQL, Power BI and Excel",
    )

    role(
        doc,
        "Data Warehouse DBA",
        "BAC Credomatic",
        "November 2017 - January 2019",
        "San Jose, Costa Rica",
        [
            "Developed and maintained ETL pipelines that loaded multiple source systems into the enterprise data warehouse.",
            "Designed and optimized tables, indexes, and views for reliable, high-performance analytical workloads.",
        ],
        "SQL Server, SSIS, SSAS, SSRS, Power BI and Azure",
    )

    section_heading(doc, "Additional Database Experience")
    bullet(doc, "Bosal, DBA and BI Consultant: designed the first stage of the main data warehouse, produced management reporting, and shared database improvement practices with development teams.", 8.7)
    bullet(doc, "Xetux Solutions, Database Manager: defined database-management policies and improved performance and data security across business areas.", 8.7)
    bullet(doc, "ACH Cloud Services, SQL and BI Consultant: administered and developed databases, created operational reporting, and documented the data environment.", 8.7)
    bullet(doc, "EducaTablet and VIGEOSOFT, Microsoft SQL Server DBA: administered production databases, improved database process response times, modeled OLTP systems, and trained development teams.", 8.7)
    bullet(doc, "Optica Caroni, Developer Analyst: delivered production workflow, database, and reporting projects using SQL Server, SSIS, SSRS, .NET, and VB6.", 8.7)

    section_heading(doc, "Education and Languages")
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(1)
    write(p, "Systems Analyst, Informatics - IUT Dr. Federico Rivero Palacio (2000 - 2004)")
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(1)
    write(p, "Training: SQL Admin Part 1; Analyzing and Visualizing Data with Power BI; Dataiku Core Designer; Python for Data Analysis")
    p = doc.add_paragraph()
    write(p, "Spanish: Native | English: Full Professional")

    doc.core_properties.author = "Pedro Javier Gutierrez Armas"
    doc.core_properties.title = "Principal SQL Database Administrator Resume"
    doc.core_properties.subject = "Resume tailored for the FullStack Principal SQL Database Administrator opportunity"
    doc.save(OUTPUT)
    print(OUTPUT)


if __name__ == "__main__":
    build_document()
