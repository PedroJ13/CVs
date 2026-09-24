from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor


OUTPUT = r"C:\Work\CVs\Output\Pedro_Gutierrez_Panama_SQL_Server_SSIS_ETL_BI_CV.docx"


def write(paragraph, value, size=9.2, bold=False, color="26333F"):
    run = paragraph.add_run(value)
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color)


def section(doc, label):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    write(p, label.upper(), 10.5, True, "174B70")


def bullet(doc, value):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Inches(0.22)
    p.paragraph_format.first_line_indent = Inches(-0.12)
    p.paragraph_format.space_after = Pt(1.7)
    p.paragraph_format.line_spacing = 1.02
    p.paragraph_format.keep_together = True
    write(p, value)


def role(doc, title, company, dates, place, points, technologies, break_before=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.page_break_before = break_before
    write(p, f"{title} | {company}", 10, True, "17365D")
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    write(p, f"{dates} | {place}", 8.7, color="596878")
    for point in points:
        bullet(doc, point)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    write(p, "Tecnologías: ", 8.6, True, "596878")
    write(p, technologies, 8.6, color="596878")


doc = Document()
page = doc.sections[0]
page.top_margin = Inches(0.52)
page.bottom_margin = Inches(0.52)
page.left_margin = Inches(0.65)
page.right_margin = Inches(0.65)
normal = doc.styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(9.2)
normal.paragraph_format.space_after = Pt(2)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(1)
write(p, "PEDRO JAVIER GUTIERREZ ARMAS", 16, True, "17365D")
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(1)
write(p, "Especialista SQL Server | SSIS, ETL, Data Warehouse y Power BI", 10.1, True, "174B70")
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(6)
write(p, "San José, Costa Rica | +506 6351-5860 | pj13eros@hotmail.com | linkedin.com/in/pedrogutierrez13", 8.6, color="596878")

section(doc, "Perfil profesional")
p = doc.add_paragraph()
p.paragraph_format.line_spacing = 1.04
write(p, "Especialista en SQL Server e ingeniería de datos con más de 15 años de experiencia en desarrollo SQL, administración de bases de datos, procesos ETL y soluciones de data warehouse y BI. Dominio de T-SQL, SSIS, procedimientos almacenados, consultas complejas, índices, optimización de procesos y modelado de datos. Experiencia en el sector bancario mediante el desarrollo y mantenimiento del data warehouse de BAC Credomatic, además de trabajo previo en Panamá como consultor DBA y BI. Ha preparado datasets y reportes en Power BI para análisis y toma de decisiones.")

section(doc, "Competencias técnicas")
for label, value in [
    ("SQL Server", "T-SQL, procedimientos almacenados, consultas complejas, vistas, índices, optimización y administración de bases de datos."),
    ("ETL y data warehouse", "SSIS, extracción, transformación y carga, integración de fuentes, SSAS, modelado dimensional y Kimball."),
    ("Calidad y rendimiento", "Validaciones, manejo de errores, integridad de datos, análisis de rendimiento y optimización de procesos."),
    ("BI y analítica", "Power BI, SSRS, Excel, datasets analíticos, visualización, tendencias y reportes ejecutivos."),
    ("Automatización", "Python, PowerShell, Snowflake SQL, dbt, PostgreSQL y MySQL."),
]:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(1.7)
    write(p, f"{label}: ", 9.1, True)
    write(p, value, 9.1)

section(doc, "Experiencia profesional")
role(doc, "Snowflake Developer / Data Engineer", "ServiceTitan", "Sep 2024 - Actualidad", "Remoto / Costa Rica", [
    "Migra lógica de reportes desde C# hacia Snowflake SQL y mantiene modelos dbt en capas silver y gold.",
    "Creó una prueba de concepto de data warehouse con metodología Kimball para mejorar estructura, reutilización y rendimiento.",
    "Optimiza cargas SQL y dbt, reduciendo en 20% el tiempo de procesamiento en una primera fase.",
], "Snowflake SQL, dbt, MetricFlow, Kimball, Snowflake Cortex")
role(doc, "Data Engineer", "SMASH Costa Rica", "May 2021 - Sep 2024", "San José, Costa Rica", [
    "Diseñó aplicaciones ETL y flujos productivos con SQL Server y SSIS según necesidades de clientes, reduciendo trabajo manual hasta en 40%.",
    "Automatizó análisis y validaciones para detectar anomalías y mejorar la exactitud de datos en 30%.",
    "Preparó, transformó y modeló datasets para Power BI, reportes y decisiones operativas.",
], "SQL Server, SSIS, Power BI, Python, PowerShell, Snowflake, Excel")
role(doc, "SQL Developer", "Intertec International", "Feb 2019 - Abr 2021", "San José, Costa Rica", [
    "Desarrolló y optimizó procedimientos almacenados, consultas SQL y flujos de base de datos para lógica empresarial y analítica.",
    "Integró diferentes fuentes mediante ETL, validaciones y manejo de errores, mejorando la exactitud de datos en más de 25%.",
    "Redujo tiempos de consulta hasta en 50% mediante optimización SQL, índices y análisis de rendimiento.",
], "SQL Server, T-SQL, SSIS, Salesforce, MySQL")
role(doc, "DBA / SQL Developer", "EL Tiempo", "Sep 2020 - Nov 2020", "Colombia", [
    "Lideró la migración de diez bases SQL Server, manteniendo integridad de datos y continuidad operativa.",
], "SQL Server, SSIS, Azure, SSRS", break_before=True)
role(doc, "DBA / SQL & BI Consultant", "Gold Data Networks", "Ene 2016 - Feb 2020", "Ciudad de Panamá, Panamá", [
    "Construyó bases relacionales y estructuras de tablas para aplicaciones y reportes; desarrolló objetos SQL y procedimientos almacenados.",
    "Implementó soluciones integrales de base de datos, ETL y BI alineadas con infraestructura y necesidades de negocio.",
], "SQL Server, SSIS, SSRS, PostgreSQL, Power BI, Excel")
role(doc, "Data Warehouse DBA", "BAC Credomatic", "Nov 2017 - Ene 2019", "San José, Costa Rica", [
    "Desarrolló y mantuvo pipelines SSIS desde múltiples fuentes hacia el data warehouse del banco.",
    "Definió y optimizó tablas, índices y vistas para cargas analíticas de alto rendimiento y análisis de grandes datasets.",
], "SQL Server, T-SQL, SSIS, SSAS, SSRS, Power BI, Azure")

section(doc, "Experiencia adicional")
bullet(doc, "Bosal: diseñó la primera etapa del data warehouse y generó reportes y presentaciones para la alta dirección.")
bullet(doc, "Xetux Solutions: creó un data lake de ventas, definió políticas de administración y mejoró rendimiento y seguridad de datos.")
bullet(doc, "ACH Cloud Services, EducaTablet y VIGEOSOFT: administración y desarrollo SQL Server, reportes, documentación y capacitación técnica.")

section(doc, "Formación e idiomas")
p = doc.add_paragraph()
write(p, "Analista de Sistemas, Informática - IUT Dr. Federico Rivero Palacio (2000-2004).")
p = doc.add_paragraph()
write(p, "Formación adicional: SQL Admin Part 1; Analyzing and Visualizing Data with Power BI; Python for Data Analysis.")
p = doc.add_paragraph()
write(p, "Español: nativo | Inglés: dominio profesional.")

doc.core_properties.author = "Pedro Javier Gutierrez Armas"
doc.core_properties.title = "Panama SQL Server SSIS ETL BI CV"
doc.save(OUTPUT)
print(OUTPUT)
