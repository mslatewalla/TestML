#!/usr/bin/env python3
"""Generate Mustafa Slatewalla Program/Delivery Management resume (2 pages)."""

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor, Twips
from pathlib import Path


NAVY = RGBColor(0x1B, 0x36, 0x5D)
ACCENT = RGBColor(0x2C, 0x5F, 0x8A)
BODY = RGBColor(0x2D, 0x2D, 0x2D)
MUTED = RGBColor(0x4A, 0x4A, 0x4A)


def set_run(run, *, size=10.5, bold=False, italic=False, color=BODY, font="Calibri"):
    run.font.name = font
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = color


def set_paragraph_format(p, *, before=0, after=0, line=1.08, align=None, left=0):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    if align is not None:
        p.alignment = align
    if left:
        pf.left_indent = Inches(left)


def add_bottom_border(paragraph, color="1B365D", size="12"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), size)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def heading(doc, text):
    p = doc.add_paragraph()
    set_paragraph_format(p, before=8, after=3, line=1.0)
    run = p.add_run(text.upper())
    set_run(run, size=11, bold=True, color=NAVY, font="Calibri")
    add_bottom_border(p, color="2C5F8A", size="8")
    return p


def body_para(doc, text, *, before=2, after=2, size=10.5, italic=False, color=BODY):
    p = doc.add_paragraph()
    set_paragraph_format(p, before=before, after=after, line=1.08)
    run = p.add_run(text)
    set_run(run, size=size, italic=italic, color=color)
    return p


def bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    set_paragraph_format(p, before=1, after=1, line=1.06)
    p.paragraph_format.left_indent = Inches(0.22)
    p.paragraph_format.first_line_indent = Inches(-0.14)
    # Clear default run and add formatted
    if p.runs:
        p.runs[0].text = ""
    run = p.add_run(text)
    set_run(run, size=10.5, color=BODY)
    return p


def job_header(doc, title, company, dates):
    p = doc.add_paragraph()
    set_paragraph_format(p, before=7, after=0, line=1.05)
    r1 = p.add_run(title)
    set_run(r1, size=11, bold=True, color=NAVY)
    r2 = p.add_run("  |  ")
    set_run(r2, size=10.5, color=MUTED)
    r3 = p.add_run(company)
    set_run(r3, size=10.5, italic=True, color=ACCENT)
    # dates on same line with tab
    tab_stops = p.paragraph_format.tab_stops
    tab_stops.add_tab_stop(Inches(7.1), WD_TAB_ALIGNMENT.RIGHT)
    r4 = p.add_run("\t" + dates)
    set_run(r4, size=10, italic=False, color=MUTED)
    return p


def set_list_bullet_font(doc):
    styles = doc.styles
    try:
        style = styles["List Bullet"]
        style.font.name = "Calibri"
        style.font.size = Pt(10.5)
        style.font.color.rgb = BODY
    except KeyError:
        pass


def build():
    doc = Document()
    set_list_bullet_font(doc)

    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    # Name
    name = doc.add_paragraph()
    set_paragraph_format(name, before=0, after=0, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
    r = name.add_run("MUSTAFA SLATEWALLA")
    set_run(r, size=20, bold=True, color=NAVY, font="Calibri")

    target = doc.add_paragraph()
    set_paragraph_format(target, before=1, after=1, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
    r = target.add_run("Program Management  ·  Delivery Management  ·  Portfolio Leadership")
    set_run(r, size=11, bold=True, color=ACCENT)

    contact = doc.add_paragraph()
    set_paragraph_format(contact, before=0, after=2, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
    r = contact.add_run("Bangalore, KA  |  +91-9591-990871  |  mslatewalla@gmail.com")
    set_run(r, size=10.5, color=MUTED)
    add_bottom_border(contact, color="1B365D", size="18")

    heading(doc, "Executive Profile")
    body_para(
        doc,
        "Strategic Program and Delivery Leader with 14+ years driving large-scale digital "
        "transformation across supply chain, legacy modernization, and e-commerce platforms. "
        "Proven success in program management and portfolio management of multi-million-dollar "
        "initiatives, leading distributed teams (20+) across geographies, and delivering "
        "measurable business impact. Strengths include resource planning, budget management, "
        "dependency management, RAID management, governance cadence, executive reporting, "
        "change management, and vendor / client management in high-complexity environments.",
        before=3,
        after=2,
    )

    heading(doc, "Core Competencies")
    comps = doc.add_paragraph()
    set_paragraph_format(comps, before=3, after=2, line=1.15)
    r = comps.add_run(
        "Program Management  •  Portfolio Management  •  Release Management  •  Resource Planning  •  "
        "Budget Management  •  Dependency Management  •  RAID Management  •  Governance Cadence  •  "
        "Executive Reporting  •  Delivery Excellence  •  Change Management  •  Agile Transformation  •  "
        "Sprint Planning  •  Backlog Management  •  Vendor / Client Management  •  Stakeholder Alignment  •  "
        "KPI / SLA & Metrics  •  Cloud Infrastructure (AWS / GCP)"
    )
    set_run(r, size=10.5, color=BODY)

    heading(doc, "Key Business Impact")
    for t in [
        "Delivered $12M+ in measurable business value through supply chain optimization and working capital improvement across enterprise client portfolios.",
        "Drove delivery excellence on legacy modernization, accelerating transformation by 30–35% and reducing manual review by 60% through structured process and change management.",
        "Optimized cloud infrastructure achieving $180K+ annual savings through resource planning, capacity planning, and budget management.",
        "Scaled engineering delivery from 0 to 35+ personnel while maintaining operational reliability above 99% through disciplined resource planning and governance cadence.",
        "Improved program predictability to 95%+ through KPI/SLA tracking, RAID management, and executive reporting to senior stakeholders.",
    ]:
        bullet(doc, t)

    heading(doc, "Professional Experience")

    job_header(
        doc,
        "Agile Project Lead – Legacy Code Modernization",
        "Publicis Sapient, Bangalore, India",
        "Sep 2025 – Present",
    )
    for t in [
        "Own program management for Legacy Code Modernization with enterprise clients; lead resource planning and delivery for a cross-functional team of 8+ senior engineers and data scientists.",
        "Lead the process and delivery excellence function across major international clients (Mn+ LOC of code); manage client relationships, cross-team dependencies, and integration across workstreams.",
        "Developed an MCP-integrated agentic framework (Crew.ai) for orchestrated modernization of Java EJB and Angular systems; applied change management and process improvements that reduced delivery timelines by 30%.",
        "Established RAID management and a metrics framework (code quality indices, architectural debt); instituted governance cadence and executive reporting aligning technical and business stakeholders.",
    ]:
        bullet(doc, t)

    job_header(
        doc,
        "Delivery Lead",
        "3sc Solution, Bangalore, India",
        "Jul 2022 – Sep 2025",
    )
    for t in [
        "Directed portfolio management of 3 large-scale Integrated Business Planning programs ($2–5M each) for Fortune 500 clients; orchestrated 20+ distributed engineers and Product Owners across the US and Europe with budget management and resource planning.",
        "Implemented a KPI/SLA framework covering schedule variance, cost performance, and defect density; maintained executive reporting dashboards that achieved 95%+ project predictability.",
        "Executed Agile transformation across 4 engineering teams through sprint planning and backlog management, reducing delivery timeline variability by 38% and increasing sprint velocity by 27%.",
        "Ran governance cadence, RAID logs, and dependency management across client, vendor, and internal workstreams to protect delivery excellence on committed milestones; orchestrated supply chain optimization delivering $12M working capital improvement and a 35% reduction in order fulfillment cycles.",
    ]:
        bullet(doc, t)

    job_header(
        doc,
        "Engineering Lead",
        "Locus.sh, Bangalore, India",
        "Sep 2020 – Jul 2022",
    )
    for t in [
        "Led an 18-member engineering team delivering a Last Mile tracking platform for 500+ logistics clients; owned vendor and client delivery commitments for a system processing 50K+ daily transactions at 99.85% uptime.",
        "Owned release management via a CI/CD pipeline that cut deployment time from 8 hours to 45 minutes; automated testing to 78% coverage enabled daily production releases.",
        "Optimized AWS spend by $180K annually through resource planning and capacity management; maintained sub-5-minute incident response.",
        "Reduced critical production incidents 70% year-over-year through RAID management, quality initiatives, and continuous monitoring.",
    ]:
        bullet(doc, t)

    job_header(
        doc,
        "Director of Engineering",
        "Altisource Business Solutions Pvt. Ltd, Bangalore, India",
        "Jul 2013 – Jun 2020",
    )
    for t in [
        "Directed engineering for Hubzu.com and Owners.com (2M+ monthly transactions); established a delivery excellence program achieving 99.95% system reliability.",
        "Designed an ML-based document classification system at 92% accuracy, reducing manual processing by 55% and enabling 10x transaction volume scale.",
        "Led complete AWS-to-GCP cloud migration with zero downtime using structured change management; implemented multi-region DR (RPO 1 hour, RTO 4 hours).",
        "Established OKR-based governance cadence aligning 35+ engineers with business objectives; built the Bangalore delivery center from 0 to 35 engineers through resource planning and budget management.",
    ]:
        bullet(doc, t)

    heading(doc, "Education & Certifications")
    for t in [
        "Bachelor of Science in Electrical, Electronics & Communications Engineering  |  University of Mumbai",
        "Agile Based Project Management Certification  |  Altisource",
        "Supply Chain Data Analytics Certification  |  IIM Mumbai",
    ]:
        p = doc.add_paragraph()
        set_paragraph_format(p, before=2, after=1, line=1.05)
        p.paragraph_format.left_indent = Inches(0.08)
        run = p.add_run(t)
        set_run(run, size=10.5, color=BODY)

    out = Path("/workspace/resumes/Mustafa_Slatewalla_Program_Delivery_Resume.docx")
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    print(f"Wrote {out}")


if __name__ == "__main__":
    build()
