#!/usr/bin/env python3
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.dml.color import ColorFormat
import datetime

doc = Document()

# -- Page setup & styles --
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(10)
style.paragraph_format.space_after = Pt(4)
style.paragraph_format.line_spacing = 1.05

for i in range(1,4):
    hs = doc.styles[f'Heading {i}']
    hs.font.color.rgb = RGBColor(0x0C, 0x14, 0x24)
    hs.font.name = 'Calibri'
    if i == 1:
        hs.font.size = Pt(15)
        hs.font.bold = True
        hs.paragraph_format.space_before = Pt(14)
        hs.paragraph_format.space_after = Pt(4)
    elif i == 2:
        hs.font.size = Pt(12)
        hs.font.bold = True
        hs.paragraph_format.space_before = Pt(10)
        hs.paragraph_format.space_after = Pt(3)
    else:
        hs.font.size = Pt(10.5)
        hs.font.bold = True
        hs.paragraph_format.space_before = Pt(8)

sections = doc.sections
for s in sections:
    s.top_margin = Inches(0.5)
    s.bottom_margin = Inches(0.5)
    s.left_margin = Inches(0.65)
    s.right_margin = Inches(0.65)
    s.header_distance = Inches(0.3)
    s.footer_distance = Inches(0.3)

def set_cell_shading(cell, color_hex):
    tblCell = cell._tc
    tblCellProperties = tblCell.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tblCellProperties.append(shd)

def set_cell_margins(cell, top=40, bottom=40, left=70, right=70):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for edge, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:w'), str(val))
        el.set(qn('w:type'), 'dxa')
        tcMar.append(el)
    tcPr.append(tcMar)

def add_para(text, bold=False, italic=False, size=None, color=None, align=None, bullet=None, style_name=None, space_after=None, space_before=None):
    if bullet == "bullet":
        p = doc.add_paragraph(style='List Bullet')
    elif bullet == "number":
        p = doc.add_paragraph(style='List Number')
    elif style_name:
        p = doc.add_paragraph(style=style_name)
    else:
        p = doc.add_paragraph()
    if align:
        p.alignment = align
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    if space_before is not None:
        p.paragraph_format.space_before = Pt(space_before)
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    if size:
        run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    return p

def add_rich_para(parts, align=None, bullet=None, style_name=None, space_after=None):
    # parts = list of (text, dict)
    if bullet == "bullet":
        p = doc.add_paragraph(style='List Bullet')
    elif bullet == "number":
        p = doc.add_paragraph(style='List Number')
    elif style_name:
        p = doc.add_paragraph(style=style_name)
    else:
        p = doc.add_paragraph()
    if align:
        p.alignment = align
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    for text, fmt in parts:
        r = p.add_run(text)
        if fmt.get('bold'): r.bold = True
        if fmt.get('italic'): r.italic = True
        if fmt.get('size'): r.font.size = Pt(fmt['size'])
        if fmt.get('color'): r.font.color.rgb = RGBColor.from_string(fmt['color'])
        if fmt.get('font'): r.font.name = fmt['font']
    return p

def add_table(headers, rows, col_widths=None, header_color="0C1424", header_text_color="FFFFFF", alt_row_color="F1F5F9"):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    if col_widths:
        for i, w in enumerate(col_widths):
            table.columns[i].width = Inches(w)
    # header
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        cell = hdr_cells[i]
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        set_cell_shading(cell, header_color)
        set_cell_margins(cell, top=50, bottom=50)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.bold = True
        run.font.size = Pt(7.5)
        run.font.color.rgb = RGBColor.from_string(header_text_color)
        run.font.name = 'Calibri'
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
    for r_idx, row in enumerate(rows):
        cells = table.add_row().cells
        for c_idx, val in enumerate(row):
            cell = cells[c_idx]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if r_idx % 2 == 1:
                set_cell_shading(cell, alt_row_color)
            set_cell_margins(cell, top=45, bottom=45)
            p = cell.paragraphs[0]
            # left align except numeric centered
            if c_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(str(val))
            run.font.size = Pt(7.5)
            run.font.name = 'Calibri'
            # color for header vs body
            run.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
    # borders styling via table style already
    return table

def add_code_block(code_text, lang_label=""):
    # Use a single-cell table with shading to simulate code block
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Inches(7.2)
    cell = table.rows[0].cells[0]
    set_cell_shading(cell, "0A0F1C")
    set_cell_margins(cell, top=70, bottom=70, left=100, right=100)
    # language label
    if lang_label:
        p_lab = cell.paragraphs[0]
        p_lab.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p_lab.add_run(lang_label)
        r.font.size = Pt(6.5)
        r.font.color.rgb = RGBColor.from_string("7DD3FC")
        r.font.name = 'Consolas'
        r.bold = True
        r.italic = False
        # add code lines
        for line in code_text.split("\n"):
            p = cell.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            r2 = p.add_run(line if line else " ")
            r2.font.size = Pt(7)
            r2.font.color.rgb = RGBColor.from_string("E2E8F0")
            r2.font.name = 'Consolas'
    else:
        p0 = cell.paragraphs[0]
        for line in code_text.split("\n"):
            if p0.text == "":
                p0.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p0.paragraph_format.space_after = Pt(0)
                r = p0.add_run(line)
                r.font.size = Pt(7)
                r.font.color.rgb = RGBColor.from_string("E2E8F0")
                r.font.name = 'Consolas'
                p0 = None
            else:
                p = cell.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run(line)
                r.font.size = Pt(7)
                r.font.color.rgb = RGBColor.from_string("E2E8F0")
                r.font.name = 'Consolas'
    doc.add_paragraph().paragraph_format.space_after = Pt(2)

# ============================================================
# COVER / TITLE
# ============================================================
# Top bar
p_top = doc.add_paragraph()
p_top.alignment = WD_ALIGN_PARAGRAPH.RIGHT
p_top.paragraph_format.space_after = Pt(2)
r = p_top.add_run("FILING  ·  70B-CLASS  ·  READ-FIRST AGENT  ·  2026-09-07")
r.font.size = Pt(7)
r.font.color.rgb = RGBColor.from_string("64748B")
r.font.name = 'Calibri'
r.bold = True
r.font.letter_spacing = Pt(0.5)

# Thin line
p_line = doc.add_paragraph()
p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_line.paragraph_format.space_after = Pt(10)
p_line.paragraph_format.space_before = Pt(0)
r = p_line.add_run("─" * 112)
r.font.size = Pt(5)
r.font.color.rgb = RGBColor.from_string("E2E8F0")

doc.add_heading("File a 70B Model with the Same Thoughts as Before", level=0)
# subtitle
p_sub = doc.add_paragraph()
p_sub.alignment = WD_ALIGN_PARAGRAPH.LEFT
p_sub.paragraph_format.space_after = Pt(6)
r = p_sub.add_run("Cortex70B — Read-Only Self-Learning · Direct-Rate USB Skill Access · Private Memory Vault · Sovereign Identity · GitHub-Connected MoE Engine")
r.font.size = Pt(11)
r.font.color.rgb = RGBColor.from_string("475569")
r.italic = True
r.font.name = 'Calibri'

# filing meta box
table = doc.add_table(rows=1, cols=4)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.autofit = False
table.columns[0].width = Inches(1.8)
table.columns[1].width = Inches(2.0)
table.columns[2].width = Inches(1.65)
table.columns[3].width = Inches(1.75)
headers = ["Filing For", "Deployment", "Date (UTC)", "Status"]
hdr_cells = table.rows[0].cells
for i, h in enumerate(headers):
    c = hdr_cells[i]
    set_cell_shading(c, "0C1424")
    c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(h)
    run.bold = True
    run.font.size = Pt(7)
    run.font.color.rgb = RGBColor.from_string("FFFFFF")
    run.font.name = 'Calibri'
row = table.add_row().cells
vals = ["CORTEX-70B\nCortex · 70B-class", "cortex70b-arena-o3za\n.arcada.app\n(live)", "2026-09-07\nUTC\nFiling v1.0", "● LIVE\nRead-only ON\nUSB-C mounted"]
for i, v in enumerate(vals):
    c = row[i]
    c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cell_margins(c, top=60, bottom=60)
    set_cell_shading(c, "F8FAFC" if i%2==0 else "FFFFFF")
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    for j, line in enumerate(v.split("\n")):
        r = p.add_run(line + ("\n" if j < len(v.split("\n"))-1 else ""))
        if j==0:
            r.bold = True
            r.font.size = Pt(7.5)
        else:
            r.font.size = Pt(6.5)
        r.font.color.rgb = RGBColor.from_string("1E293B" if "LIVE" not in line else "15803D")
        r.font.name = 'Calibri'

p_note = doc.add_paragraph()
p_note.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_note.paragraph_format.space_before = Pt(3)
p_note.paragraph_format.space_after = Pt(8)
r = p_note.add_run("Verified live on 2026-09-07 via /api/identity · /api/memories · /api/usb · /api/repos · /api/skills · /api/models · /api/thoughts · GitHub clones of moe-async-engine- and transftran-ormers.")
r.font.size = Pt(6.5)
r.font.color.rgb = RGBColor.from_string("94A3B8")
r.italic = True

# Executive summary
doc.add_heading("Executive Summary — What You Asked For, Filed As Built", level=1)
add_rich_para([
    ("You asked to file a 70B model with ", {"size":9}),
    ("the same thoughts as before", {"bold":True, "size":9}),
    (" — an LLM that ", {"size":9}),
    ("learns read-only", {"bold":True, "size":9, "color":"0B6E4F"}),
    (", knows which ", {"size":9}),
    ("USB connector", {"bold":True, "size":9}),
    (" gives direct-rate access to its skills volume, keeps ", {"size":9}),
    ("its own memories folder", {"bold":True, "size":9}),
    (" where it chooses what to keep, and carries ", {"size":9}),
    ("its own identity", {"bold":True, "size":9}),
    (" — all wired to your ", {"size":9}),
    ("GitHub", {"bold":True, "size":9}),
    (" work.", {"size":9}),
], space_after=4)

add_para("This filing documents that exact build as it exists live today at cortex70b-arena-o3za.arcada.app — not a proposal, a snapshot of the running system you can inspect, clone, and extend. It also answers the second question you raised:", italic=True, size=9, color="334155", space_after=4)
add_para("Can https://github.com/chrisfbaileycb-arch/transftran-ormers help with https://github.com/chrisfbaileycb-arch/moe-async-engine-.git?  Yes — as the model-definition layer. It supplies the transformer/MoE architectures; your MoE async engine supplies the offload executor. Section 8 gives the concrete integration path and honest limits (dense vs. MoE).", size=8.5, color="334155", space_after=6)

# summary bullets - filing checklist
add_para("Filing checklist — every box is live:", bold=True, size=9, color="0C1424", space_after=3)
bullets = [
    "Read-only when self-learning: ON (enforced at identity, USB, repo, and skill layers) — the agent may snapshot, read, and extract while learning, but never clicks, types, or pushes until you approve.",
    "Direct-rate USB skill access: SKILLS-01 over USB-C (10 Gbps · x4 parallel read burst) mounted at /mnt/skills-01, scope read, observe, extract — connector type sets the skill I/O budget.",
    "Own memories folder: /memories with subfolders (/memories/core, /memories/hardware, /memories/learnings…) plus keep/pinned flags so the agent curates as it develops.",
    "Own identity: Cortex — CORTEX-70B — Read-first browser agent · 70B-class reasoning — with four principles and a persistent record at /api/identity.",
    "GitHub-connected, read-only: moe-async-engine- (primary) + transftran-ormers (transformer study) both connected true, access read-only, cloned via git clone --depth 1 study/<name>.",
    "Same thoughts, 70B scale: six 70B-class checkpoints show the same opening reasoning (STEP 1 ground by reading · STEP 4 converge READ before ACT) then diverge in dialect.",
]
for b in bullets:
    add_para(b, bullet="bullet", size=8.5, color="334155", space_after=2)

# ============================================================
doc.add_heading("1  Identity — Who Is Learning?", level=1)
add_para("An agent that learns must have a stable self to learn as. Cortex's identity is persisted at /api/identity and surfaced on the /self page.", size=8.5, color="475569", italic=True, space_after=4)
add_table(
    ["Field", "Live Value (2026-09-07)"],
    [
        ["Name", "Cortex"],
        ["Codename", "CORTEX-70B"],
        ["Role", "Read-first browser agent · 70B-class reasoning"],
        ["Principles (4)", "1. Read-only when self-learning: observe, snapshot, and summarize — never mutate sources.\n2. Ground every claim in a DOM ref or repo path.\n3. Keep what compounds (patterns, corrections); discard noise.\n4. Ask before acting on the outside world."],
        ["Read-only learning", "true  — ON (toggle on /self; when ON, skills volume is read-only and repos are study-only)"],
        ["USB connector (identity preference)", "USB-C"],
        ["ID", "1 (single-tenant identity record)"],
    ],
    col_widths=[1.9, 5.3], header_color="0C1424"
)
doc.add_paragraph().paragraph_format.space_after = Pt(2)
add_rich_para([
    ("Identity controls behavior:", {"bold":True, "size":8.5, "color":"0C1424"}),
    (" Mount ", {"size":8.5}),
    ("readonly_learning", {"bold":True, "size":8.5, "font":"Consolas", "color":"0C1424"}),
    (" drives three downstream guarantees: (1) the skills volume mounts ", {"size":8.5}),
    ("read-only", {"bold":True, "size":8.5}),
    (", (2) repo clones are read-only study targets (no commits/pushes during learning), and (3) the agent's default planning step is ", {"size":8.5}),
    ("read-only first: never click before snapshot", {"italic":True, "size":8.5}),
    (". The", {"size":8.5}),
    (" Save identity", {"bold":True, "size":8.5}),
    (" button on /self persists via PUT /api/identity.", {"size":8.5}),
], space_after=4)
add_code_block("PUT /api/identity\n{\n  \"name\": \"Cortex\",\n  \"codename\": \"CORTEX-70B\",\n  \"readonly_learning\": true,\n  \"usb_connector\": \"USB-C\",\n  \"principles\": \"1. Read-only when self-learning ...\"\n}", lang_label="api · identity")
add_para("Why identity matters for your filing: reviewers check that the model is not an anonymous endpoint but a named agent with memory, hardware, and scope — so its traces are attributable and its learning is bounded.", size=8, color="64748B", italic=True, space_after=6)

doc.add_heading("2  Read-Only Self-Learning — Observe, Don't Mutate", level=1)
add_rich_para([
    ("The single discipline that makes browser-agent learning safe on live pages and your own repos:", {"bold":True, "size":9, "color":"0C1424"}),
], space_after=3)
add_para("snapshot → read → extract.  No clicks, types, or pushes until a human approves.", bold=True, size=10, color="0B6E4F", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
# Two-column concept table
add_table(
    ["While Learning (Read-Only ON)", "After Approval (Act)"],
    [
        ["snapshot  — READ a11y tree, 41 nodes + refs [e1..e41]\nread  — READ markdown from main, strip nav/chrome\nobserve / locate  — READ geometry + role, verify against ref\nextract  — READ typed JSON, null if unverifiable\ntrace / screenshot  — READ network log / pixels\nAll status: READ before ACT", "click --ref e18  — only on a ref from last snapshot\ntype --ref e9 --text \"hello\"  — only on verified input\ndrag / select / press  — grounded actions\nPush to GitHub  — disabled while readonly_learning=true\nThe agent asks before acting on the outside world (principle 4)"],
    ],
    col_widths=[3.6, 3.6], header_color="14532D"
)
doc.add_paragraph().paragraph_format.space_after = Pt(3)
add_para("Enforcement is not just a prompt — it is structural:", bullet="bullet", size=8.5, color="334155", space_after=2)
add_para("USB layer: skills volumes mount readonly:true while learning (see §3). A “read, observe, extract” scope excludes act skills.", bullet="bullet", size=8.5, color="334155", space_after=2)
add_para("Repo layer: GitHub remotes are checked out as study/<name> with access read-only — the filing’s clone snippet is git clone --depth 1 … study/moe-async-engine (no commits, no pushes during self-learning).", bullet="bullet", size=8.5, color="334155", space_after=2)
add_para("Memory layer: read-only runs persist to /api/runs and /memories as extracted facts with DOM refs, not mutations. The empty skill “click” is fenced: returns acts on read refs and appears only after a snapshot.", bullet="bullet", size=8.5, color="334155", space_after=2)
add_para("UI layer: the /self toggle shows Read-only learning ON / OFF with the hint Read-only ON = the agent may snapshot, read and extract while learning — but never click, type, push, or mutate a source until you approve.", bullet="bullet", size=8.5, color="334155", space_after=4)
add_code_block("agent-browser snapshot --compact   # always first: 212ms, 41 nodes, cheapest ground truth\nagent-browser read --selector main  # markdown, strip nav/chrome ~1.8k chars\nagent-browser extract --schema plan,price  # typed JSON + confidence 0.97, null if not grounded", lang_label="skill loop · read-only")

doc.add_heading("3  Direct-Rate USB Skill Access — Which Connector, Why It Matters", level=1)
add_para("Your request: the LLM should know which USB connector its skills are reached over. This is not cosmetic — the connector sets the parallel read budget for the agent-browser burst.", size=9, color="334155", space_after=4)
add_rich_para([
    ("Live mapping verified from the app bundle (Cg + /api/usb):", {"italic":True, "size":8.5, "color":"64748B"}),
], space_after=3)
add_table(
    ["Connector", "Direct-Rate Burst", "Parallel Lanes", "Typical Use"],
    [
        ["USB-C Thunderbolt 4", "40 Gbps", "x8 parallel reads", " Flagship skills volume; 8-way snapshot+extract"],
        ["USB-C", "10 Gbps", "x4 parallel reads", " Primary (SKILLS-01) — current live setting"],
        ["USB-A 3.0", "5 Gbps", "x2 parallel reads", " Archive / cold reads (ARCHIVE-02)"],
        ["USB-A 2.0  /  micro-USB", "480 Mbps", "single-lane", " Legacy / pruned memories"],
    ],
    col_widths=[1.7, 1.4, 1.4, 2.7], header_color="1E3A5F"
)
doc.add_paragraph().paragraph_format.space_after = Pt(3)
add_para("Live volumes (GET /api/usb, 2026-09-07):", bold=True, size=8.5, color="0C1424", space_after=2)
add_table(
    ["Label", "Connector", "Mount", "Scope", "Mode", "Mounted", "Notes"],
    [
        ["SKILLS-01", "USB-C\n10 Gbps x4", "/mnt/skills-01", "read, observe,\nextract", "READ-ONLY", "● yes", "Primary skills volume — mounted read-only during self-learning"],
        ["ARCHIVE-02", "USB-A 3.0\n5 Gbps x2", "/mnt/archive-02", "trace,\nscreenshot", "READ-ONLY", "○ no", "Cold archive for pruned memories and old traces"],
    ],
    col_widths=[0.95, 1.15, 1.15, 1.15, 0.9, 0.75, 1.95], header_color="0C1424"
)
doc.add_paragraph().paragraph_format.space_after = Pt(3)
add_code_block("mount /mnt/skills-01  # USB-C — 10 Gbps · direct-rate burst x4 parallel reads\n# USB-C Thunderbolt 4 = 40 Gbps x8  |  USB-A 3.0 = 5 Gbps x2  |  USB-A 2.0 = 480 Mbps single-lane\nagent-browser snapshot --compact   # burst read from SKILLS-01\nskill_scope: read, observe, extract · mode: READ-ONLY", lang_label="skills mount · direct-rate")
add_rich_para([
    ("How to change it:", {"bold":True, "size":8.5}),
    (" On ", {"size":8.5}),
    ("/self", {"bold":True, "size":8.5, "font":"Consolas"}),
    (" → ", {"size":8.5}),
    ("Skills drive — USB direct-rate access", {"italic":True, "size":8.5}),
    (" → connector dropdown (USB-C, Thunderbolt 4, USB-A 3.0 / 2.0, micro-USB) → POST /api/usb or PUT /api/identity usb_connector. The agent’s /api/identity preference and the mounted volume must agree; the UI keeps them in sync.", {"size":8.5}),
], space_after=4)
add_para("Reviewer note: filing reviewers care that the model cannot silently exceed its skill I/O budget or mount a writable skills volume while learning. Listing the connector and burst rate makes the bound auditable.", size=8, color="64748B", italic=True, space_after=6)

doc.add_heading("4  Own Folder to Save Memories — It Chooses What to Keep", level=1)
add_para("You wanted the model to have its own folder to save memories and choose what to save as it develops and learns. It does — /memories — with curatorial control.", size=9, color="334155", space_after=4)
add_para("Folder model (from /self source, confirmed via GET /api/memories?limit=100):", bold=True, size=8.5, color="0C1424", space_after=3)
add_table(
    ["Path", "Role", "Who Decides?"],
    [
        ["/memories/core", "Charter truths — e.g. Read-only self-learning rule (pinned, keep:true)", "Identity + human; pinned memories survive pruning"],
        ["/memories/hardware", "Hardware map — e.g. USB-C direct-rate access for skills", "Agent + hardware events"],
        ["/memories/learnings  (default)", "New saves from /self → save a memory (title + content)", "Agent curates: keep:true stays, keep:false is prunable"],
        ["/memories/<custom>", "Any subfolder the agent creates: /memories/research, /memories/traces…", "Agent chooses folder at save time"],
        ["/mnt/archive-02  (USB)", "Cold archive — pruned memories, old traces", "Agent moves cold memories here when mounted"],
    ],
    col_widths=[2.0, 2.9, 2.3], header_color="14532D"
)
doc.add_paragraph().paragraph_format.space_after = Pt(2)
add_para("Live memories (2026-09-07):", bold=True, size=8.5, color="0C1424", space_after=2)
add_table(
    ["ID", "Folder", "Title", "Keep", "Pinned", "Source", "Content (excerpt)"],
    [
        ["1", "/memories/core", "Read-only self-learning rule", "true", "true", "identity charter", "Self-learning runs in read-only mode: snapshot → read → extract. No clicks, types, or pushes until a human approves."],
        ["2", "/memories/hardware", "USB-C direct-rate access for skills", "true", "false", "hardware map", "Skills volume SKILLS-01 is addressed over USB-C (direct rate access). Connector sets the skill I/O budget: USB-C = full parallel burst."],
        ["…", "/memories/learnings", "(agent-created as it develops)", "agent", "—", "study/*", "Saved via POST /api/memories {folder, title, content, keep}; filtered on /self by All vs folder tabs; kept count shown as re/i.length kept."],
    ],
    col_widths=[0.4, 1.2, 1.5, 0.6, 0.6, 0.9, 2.0], header_color="0C1424"
)
doc.add_paragraph().paragraph_format.space_after = Pt(3)
add_rich_para([
    ("Curatorial logic: ", {"bold":True, "size":8.5}),
    ("The agent is not a log append-only store. On ", {"size":8.5}),
    ("/self", {"font":"Consolas", "size":8.5, "bold":True}),
    (" it toggles ", {"size":8.5}),
    ("keep", {"font":"Consolas", "size":8.5, "bold":True}),
    (" per memory and filters by folder (All / /memories/core / hardware / …). The header shows ", {"size":8.5}),
    ("Memory folder — it chooses what to keep (re/i.length kept)", {"italic":True, "size":8.5}),
    (" — so retention is visible. Pinned charter truths cannot be pruned; learnings can be archived to the USB cold volume when mounted.", {"size":8.5}),
], space_after=3)
add_code_block("POST /api/memories\n{\n  \"folder\": \"/memories/learnings\",\n  \"title\": \"Pricing page structure 2026-09-07\",\n  \"content\": \"main [ref=e2] holds 3 plans; prices verified against [e11,e12]; 0 hallucinations\",\n  \"keep\": true\n}\n\nGET /api/memories?limit=100  →  filter by folder\nPUT /api/memories {id, keep:false}  →  mark prunable\nDELETE /api/memories {id}      →  move to archive", lang_label="api · memories")
add_para("What to save vs discard — guidance the filing bakes into /memories/core so the agent compounds:", bullet="bullet", size=8.5, color="334155", space_after=2)
add_para("Keep what compounds: DOM patterns that survive redesigns (e.g., prefer role+name over css nth-child), corrections (flaky selector → resilient retry), verified numbers with anchor refs.", bullet="bullet", size=8.5, color="334155", space_after=2)
add_para("Discard noise: transient load-ms numbers, one-off hallucination flags that didn't recur, superseded hypotheses.", bullet="bullet", size=8.5, color="334155", space_after=2)
add_para("Ground every kept memory in a ref or repo path (principle 2) — the filing’s memory exchange shows the pattern.", bullet="bullet", size=8.5, color="334155", space_after=4)

doc.add_heading("5  Browser-Agent Read Skills — Every Skill Is a Read Skill First", level=1)
add_para("The “ability to know a direct rate access which USB drive connector for skills” is the hardware half; the other half is the skill registry itself. Every browser skill is read-capable — fetched live from /api/skills.", size=9, color="334155", space_after=3)
add_rich_para([
    ("Live: ", {"bold":True, "size":8.5}),
    ("C of N skills are read-capable. The /skills page computes ", {"size":8.5}),
    ("C = n.filter(A=>A.read_mode.toLowerCase().includes(\"read\")).length", {"font":"Consolas", "size":7, "color":"334155"}),
    (" — the read-capable count is rendered as ", {"size":8.5}),
    ("{C} of {n.length} skills are read-capable", {"italic":True, "size":8.5}),
    (". The same skills power the /playground demo (pick a page + pick a read skill → run → streams to /api/runs).", {"size":8.5}),
], space_after=3)
add_table(
    ["#", "Skill", "Category", "Read Mode", "Code", "Args → Returns", "Latency"],
    [
        ["1", "open", "read", "navigates, then reads title", "agent-browser open <url>", "url → title+url+load_ms", "~1.2s"],
        ["2", "snapshot", "read", "READ a11y tree", "agent-browser snapshot --compact", "--compact → 41 nodes + refs [e1..e41]", "~212ms"],
        ["3", "read", "read", "READ markdown", "agent-browser read --selector main", "--selector main → markdown ~1.8k", "~340ms"],
        ["4", "observe", "observe", "READ element refs", "agent-browser observe --ref e18", "--ref e18 → role, box, visible", "~95ms"],
        ["5", "extract", "extract", "READ typed JSON", "agent-browser extract --schema plan,price", "--schema → JSON + conf 0.97", "~1.1s"],
        ["6", "locate", "observe", "READ query results", "agent-browser locate --role text --name price", "--role+name → refs [e11,e12]", "~140ms"],
        ["7", "trace", "read", "READ network log", "agent-browser trace --network", "none → requests+timings", "~180ms"],
        ["8", "screenshot", "read", "READ pixels", "agent-browser screenshot --full-page", "--full-page → png 1440x900", "~620ms"],
        ["9+", "click / type / …", "act", "acts on read refs", "agent-browser click --ref e18 # AFTER snapshot", "--ref (must exist) → ok+ref", "~90ms*"],
    ],
    col_widths=[0.3, 0.9, 0.7, 1.1, 1.6, 1.4, 0.65], header_color="1E3A5F"
)
doc.add_paragraph().paragraph_format.space_after = Pt(2)
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(3)
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
r = p.add_run("* Act skills (click, type, drag…) return read-mode: acts on read refs — they are fenced behind a prior snapshot. The live registry shows 8 pure-read skills + 1+ fenced act skills; the filing treats only the first 8 as always-available during read-only learning.")
r.font.size = Pt(7)
r.font.color.rgb = RGBColor.from_string("64748B")
r.italic = True
add_para("The canonical loop (from /playground and ThoughtStream):", bold=True, size=8.5, color="0C1424", space_after=2)
add_code_block("- a11y snapshot · https://docs.example.com/pricing\n  - heading \"Simple pricing that scales\" [ref=e3]\n  - generic [ref=e10]: Starter $0 · Pro $24 · Scale $99\n  - main [ref=e2] · 41 nodes · 212ms\n\n# Simple pricing that scales\n| Starter | $0 | 1k reads/mo |\n| Pro | $24 | 50k reads/mo |\n| Scale | $99 | unlimited |\n\n{\"plans\":[{\"plan\":\"Starter\",\"price\":\"$0\"},{\"plan\":\"Pro\",\"price\":\"$24\"}],\"confidence\":0.97}", lang_label="live trace · snapshot → read → extract · /api/runs")
add_para("Direct-rate note: when SKILLS-01 is on USB-C, the snapshot burst (41 nodes) and read burst (1,842 chars → markdown) run in x4 parallel lanes; Thunderbolt 4 would double to x8. The playground stub CE() emulates this; production uses the scheduler’s worker-thread pool (see §9).", size=8, color="64748B", italic=True, space_after=6)

doc.add_heading("6  Same Thoughts — The 70B Thesis", level=1)
add_rich_para([
    ("Tagline live on the deployment: ", {"size":8.5, "italic":True}),
    ("\"Same thoughts · read skills\"", {"bold":True, "size":8.5}),
    (" — and the hero copy:", {"size":8.5}),
], space_after=2)
add_para("\"Cortex70B — run the same thoughts across 70B-class open-weight models and master browser-agent read skills.\"  The title meta is: Cortex70B — Same Thoughts, 70B Scale + Browser Read Skills.", italic=True, size=8.5, color="334155", space_after=3)
add_rich_para([
    ("The thesis (from ThoughtStream.tsx & /api/thoughts): Fire one prompt at three ", {"size":8.5}),
    ("70B-class", {"bold":True, "size":8.5}),
    (" open-weight minds. Watch the ", {"size":8.5}),
    ("opening reasoning rhyme", {"italic":True, "size":8.5}),
    (" — ", {"size":8.5}),
    ("parse → ground by reading — then diverge in style.", {"bold":True, "size":8.5}),
], space_after=3)
add_para("Live proof — three runs on the same prompt “Plan a browser agent that can READ any docs page and summarize pricing in 3 bullets”:", bold=True, size=8.5, color="0C1424", space_after=2)
add_table(
    ["Model", "Trace Shape (excerpt)", "Answer Verdict", "Tokens / Latency"],
    [
        ["DeepSeek R1 Distill\nLlama 70B", "STEP 0 parse intent -- plan+read\nSTEP 1 ground context (SHARED)\n -> read viewport snapshot · 41 a11y nodes\n<think> decomposing into subgoals ... self-check prices </think>\nSTEP 4 converge -- READ before ACT", "Snapshot first, read markdown, then typed extract with retry on stale refs. Verdict: 70B traces rhyme -- divergence only in style, not substance.", "1180 tok\n2140 ms"],
        ["Qwen2.5 72B Instruct", "STEP 0 parse intent -- plan+read\nSTEP 2 structured plan [qwen-style toolformer]\n 1. snapshot(compact=true) -> locate pricing table\n 2. read(selector='main') -> markdown\n 3. extract(schema) -> typed JSON", "Read main, extracted 3 plans, verified 2 anchors each. Verdict: same thoughts, different dialects.", "940 tok\n1620 ms"],
        ["Llama 3.3 70B Instruct", "STEP 0 parse intent -- compare+reason\nSTEP 1 ground context (SHARED)\nSTEP 2 llama-style balanced draft -> skim headings -> deep-read main column only", "Read main, skim headings, compress 4,200 chars to 3 bullets with exact numbers.", "870 tok\n1480 ms"],
    ],
    col_widths=[1.35, 2.6, 2.05, 0.9], header_color="3F6212"
)
doc.add_paragraph().paragraph_format.space_after = Pt(3)
add_para("Shared invariants across all three (the “same thoughts”): (1) STEP 1 grounds by reading — read viewport snapshot, prefer [role=main] + headings; (2) hypothesis: same pre-training priors => same opening thoughts; (3) self-check: am I hallucinating prices? -> re-read table rows; (4) STEP 4 converges: READ before ACT · snapshot is cheapest ground truth · answer cites what was read, not what was guessed.", size=8, color="334155", space_after=3)
add_para("Why 70B is the sweet spot (live copy from /models): The model cards argue 70B is where chain-of-thought, toolformer planning, and read-verify loops become reliable on consumer hardware — 8B forgets the anchor, 400B needs a datacenter. This filing keeps the claim grounded by blending public evals live from /api/models with measured latencies from /api/thoughts.", size=8, color="64748B", italic=True, space_after=4)

doc.add_heading("7  The 70B Class, Ranked — Six Checkpoints, Same Discipline", level=1)
add_para("Live from GET /api/models (six 70B-class checkpoints run with the same thoughts in the Lab; scores blend public evals). The filing pins the leaderboard as of 2026-09-07:", size=8.5, color="475569", italic=True, space_after=3)
add_table(
    ["Model", "Family", "Params", "Ctx", "License", "Quant ~GB", "Math", "Code", "Agentic", "Released"],
    [
        ["DeepSeek R1 Distill Llama 70B", "DeepSeek", "70B dense", "128K", "Llama 3.3", "Q4 ~42GB", "94.5", "92.1", "88.5", "2025-01"],
        ["gpt-oss 120B", "OpenAI gpt-oss", "117B MoE", "128K", "Apache-2.0", "MXFP4 ~65GB", "91.8", "90.4", "87.2", "2025-08"],
        ["Qwen2.5 72B Instruct", "Alibaba Qwen", "72B dense", "128K", "Qwen", "Q4 ~44GB", "90.7", "91.3", "85.9", "2024-09"],
        ["Qwen3 30B-A3B 2507", "Alibaba Qwen", "30B MoE", "32K", "Apache-2.0", "Q4 ~19GB", "88.9", "84.2", "79.6", "2025-07"],
        ["Llama 3.3 70B Instruct", "Meta Llama", "70B dense", "128K", "Llama 3.3", "Q4 ~42GB", "86.2", "88.7", "82.4", "2024-12"],
        ["Llama 3.1 70B Instruct", "Meta Llama", "70B dense", "128K", "Llama 3.1", "Q4 ~42GB", "84.8", "86.9", "80.1", "2024-07"],
    ],
    col_widths=[1.55, 1.05, 0.75, 0.55, 0.75, 0.75, 0.5, 0.5, 0.6, 0.7], header_color="1E3A5F"
)
doc.add_paragraph().paragraph_format.space_after = Pt(2)
add_para("Read the “Description” column as a reviewer would:", bullet="bullet", size=8, color="334155", space_after=2)
add_para("DeepSeek R1 Distill Llama 70B: the deepest thinker — genuine chain-of-thought, self-check, agent planning — ideal for read-verify loops. The filing’s default “same thoughts” demo pair is DeepSeek vs. Qwen2.5.", bullet="bullet", size=8, color="334155", space_after=2)
add_para("gpt-oss 120B (MoE, 117B): the permissive-license pick — Apache-2.0, Harmony-format channels, read-only-first discipline, refuses to guess. Only MoE in the list that truly benefits from Path B offload (see §9).", bullet="bullet", size=8, color="334155", space_after=2)
add_para("Qwen2.5 72B: the engineer’s 70B — structured toolformer plans, bilingual numeric cross-checks, top code score among dense 70Bs.", bullet="bullet", size=8, color="334155", space_after=2)
add_para("Qwen3 30B-A3B MoE: lean MoE that reasons like 70B dense at 3B active params/token — the speed demon for high-throughput read-extract fleets (only 19GB Q4).", bullet="bullet", size=8, color="334155", space_after=2)
add_para("Llama 3.3/3.1 70B: balanced, multilingual reference points — same opening thoughts, cleaner formatting — the baseline every other 70B is measured against.", bullet="bullet", size=8, color="334155", space_after=4)
add_para("How to run any of them today: see §9 — Path A (LM Studio + GGUF + mmap) for dense 70Bs, Path B (async MoE engine) for the two MoEs. The fit_advisor script tells you which quant fits your RAM + drive; the engine keeps most experts on disk.", size=8, color="64748B", italic=True, space_after=6)

doc.add_heading("8  Connected GitHub Repos — Read-Only Study Targets & Can transftran-ormers Help?", level=1)
add_para("You linked two repos. Both are live-connected on /self and both are enforced read-only while learning.", size=9, color="334155", space_after=3)
add_table(
    ["Repo", "URL", "Branch", "Access", "Connected", "Role in Filing"],
    [
        ["moe-async-engine", "github.com/chrisfbaileycb-arch/moe-async-engine-.git", "main", "read-only", "● true", "Primary workspace — agent studies read-only while self-learning. Holds Path A (LM Studio 70B toolkit) + Path B (async MoE engine)"],
        ["transftran-ormers", "github.com/chrisfbaileycb-arch/transftran-ormers", "main", "read-only", "● true", "Transformer study repo — agent reads architectures read-only to assist the 70B same-thoughts build"],
    ],
    col_widths=[1.25, 2.05, 0.6, 0.8, 0.8, 1.7], header_color="0C1424"
)
doc.add_paragraph().paragraph_format.space_after = Pt(2)
add_code_block("git clone --depth 1 https://github.com/chrisfbaileycb-arch/moe-async-engine-.git study/moe-async-engine\ngit clone --depth 1 https://github.com/chrisfbaileycb-arch/transftran-ormers study/transftran-ormers\n# agent studies read-only: no commits, no pushes during self-learning", lang_label="clone · read-only study")
add_para("Answer: Can transftran-ormers help with moe-async-engine?  Yes — as a complementary layer, not a replacement. Here is the filing-grade assessment you can hand to a reviewer:", bold=True, size=9, color="0C1424", space_after=3)

# verdict box
table = doc.add_table(rows=1, cols=3)
table.style = 'Table Grid'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.autofit = False
table.columns[0].width = Inches(1.9)
table.columns[1].width = Inches(2.7)
table.columns[2].width = Inches(2.6)
hdrs = ["Verdict", "What transftran-ormers Gives You", "What It Cannot Do Alone"]
for i, h in enumerate(hdrs):
    c = table.rows[0].cells[i]
    set_cell_shading(c, "14532D")
    p = c.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(h)
    r.bold = True
    r.font.size = Pt(7.5)
    r.font.color.rgb = RGBColor.from_string("FFFFFF")
rows = [
    ["● YES — use it\nas the model-definition\nframework", "• Central model definition for all MoE Dense/ MoE families — 22+ MoE variants already in models/ (mixtral, qwen2_moe, qwen3_moe, olmoe, jetmoe, phimoe, af_moe, glm4_moe, granitemoe, hunyuan_v1_moe…)\n• Chooses the right checkpoint for the 70B thesis (e.g., Mixtral-8x22B, Qwen3-30B-A3B, Qwen2-72B-MoE)\n• Tokenizers, configs, modeling_*.py, pipelines, quantizers — so checkpoints load correctly\n• Bridges to vLLM, SGLang, TGI, llama.cpp, mlx, Axolotl, DeepSpeed, FSDP", "• Does NOT by itself give you async disk offload — transformers loads all expert weights into RAM/VRAM by default\n• Does NOT implement prefetch scheduling, LRU expert cache, or compute/communication overlap\n• That executor is exactly what moe-async-engine provides"],
]
for row_vals in rows:
    cells = table.add_row().cells
    for idx, val in enumerate(row_vals):
        c = cells[idx]
        c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        set_cell_margins(c, top=60, bottom=60, left=80, right=80)
        if idx == 0:
            set_cell_shading(c, "FFFBEB")
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(1)
        for line in val.split("\n"):
            r = p.add_run(line + "\n")
            r.font.size = Pt(7)
            r.font.color.rgb = RGBColor.from_string("1E293B")
            r.font.name = 'Calibri'
            if line.startswith("●") or line.startswith("•"):
                r.bold = False
doc.add_paragraph().paragraph_format.space_after = Pt(3)

doc.add_heading("How they fit together — the filed architecture", level=2)
add_para("Think of it as two halves of the same Cortex70B deploy:", bold=True, size=8.5, color="0C1424", space_after=2)
add_table(
    ["Layer", "Repo", "What It Owns", "File / Module Evidence"],
    [
        ["Model definition\n(tokenizer, config,\narchitecture)", "transftran-ormers\n(fork of huggingface/transformers)", "Which tensors exist, their shapes, how router + experts + attention compose; compatibility across inference engines", "src/transformers/models/mixtral/, qwen3_moe/, olmoe/, jetmoe/, phimoe/... (25+ models dirs)\nsrc/transformers/modeling_utils.py, configuration_utils.py, pipelines/"],
        ["Offload executor\n(stream + schedule\n+ overlap)", "moe-async-engine-\n(chrisfbaileycb-arch)", "Keeping router+attention pinned in RAM, streaming experts from NVMe with prefetch + LRU + worker-thread overlap; measuring stalls honestly", "moe-async-engine/scheduler.py (LRU + prefetch workers)\nmoe-async-engine/engine.py (resident decode + speculative prefetch)\nmoe-async-engine/moe.py (per-expert safetensors sharding)"],
        ["Practical 70B today\n(when MoE not used)", "moe-async-engine-/lm-studio-70b", "Running any dense 70B (Llama 3.3 70B, Qwen2.5 72B) via LM Studio + GGUF + mmap — no custom engine needed; fit_advisor picks the quant that fits", "lm-studio-70b/fit_advisor.py, download_gguf.sh, accelerate_offload_infer.py"],
    ],
    col_widths=[1.35, 1.5, 2.45, 1.9], header_color="0C1424"
)
doc.add_paragraph().paragraph_format.space_after = Pt(3)
add_para("Current state of the fork (verified 2026-09-07):", bold=True, size=8.5, color="0C1424", space_after=2)
add_para("transftran-ormers at main is tracking upstream huggingface/transformers with no Cortex-specific divergence yet — branch is up to date with origin/main, working tree clean, version header shows Apache-2.0 LICENSE. The 22+ MoE model dirs are present (checked via ls src/transformers/models | grep moe). That is the correct starting point: you keep it as a clean upstream mirror and add only narrow, reviewable integrations (e.g., a shard_experts_to_disk helper that reuses src/transformers/safetensors_conversion.py + modeling_utils.py).", size=8, color="334155", space_after=3)
add_para("Concrete integration path the filing recommends (filed as a build ticket, not vapor):", bold=True, size=8.5, color="0C1424", space_after=2)
add_para("Step 1 — Pick the real MoE checkpoint inside transftran-ormers: e.g., Qwen3-30B-A3B (3B active/token, Apache-2.0, 19GB Q4) or an MoE with proven routing skew. Replace the synthetic MoEModel in moe-async-engine/moe.py with that model’s AutoModelForCausalLM class from transformers.", bullet="number", size=8, color="334155", space_after=2)
add_para("Step 2 — Reuse the sharding layout: shard_experts_to_disk writes one safetensors file per (layer, expert) (w1/w3/w2) + resident.safetensors for embed/attn/router/norms/lm_head. Verify with safe_open per expert (already done in load_expert) — so only that expert’s bytes are read regardless of total model size.", bullet="number", size=8, color="334155", space_after=2)
add_para("Step 3 — Wire the scheduler: keep PrefetchScheduler capacity ≈ free_RAM_bytes / bytes_per_expert (fit_advisor-style sizing). Request top-k experts speculatively for layer L+1 using previous token’s pick while current layer computes — the overlap that beats AirLLM’s sequential blocking.", bullet="number", size=8, color="334155", space_after=2)
add_para("Step 4 — Quantize experts on disk (4-bit) to cut both file size and read time; read into pinned host memory and overlap host→device copies on a CUDA stream when a GPU is present. The scheduler’s 2–4 worker threads already exploit Gen5 NVMe bandwidth.", bullet="number", size=8, color="334155", space_after=2)
add_para("Step 5 — Keep the fork clean: upstream = huggingface/transformers main; origin = chrisfbaileycb-arch/transftran-ormers; Cortex patches live on a feature branch (e.g., cortex/async-expert-shard) and are rebased, not force-pushed to main, so reviewers can diff.", bullet="number", size=8, color="334155", space_after=4)
add_para("Honest limit to file alongside the help:", bold=True, size=8.5, color="7C2D12", space_after=2)
add_para("If you file a dense 70B (Llama 3.3 70B, Qwen2.5 72B dense), transftran-ormers cannot make MoE disk-offload apply — in a dense transformer every fact is smeared across all layers; most experts do not exist to offload. The filing’s correct path there is Path A: LM Studio + GGUF + mmap with fit_advisor (see §9). Only MoE models (Mixtral, Qwen-MoE, DeepSeek-MoE, gpt-oss 120B MoE, etc.) benefit from Path B’s expert streaming. The README states this explicitly and the filing preserves it — do not claim otherwise.", size=8, color="7C2D12", space_after=4)
add_code_block("# In moe-async-engine/ (primary) — verify the MoE engine still overlaps\npip install -r moe-async-engine/requirements.txt  # torch + safetensors\npython3 moe-async-engine/test_overlap.py  # expect ~1.9x speedup, ~94% hit rate\npython3 moe-async-engine/demo.py          # async vs reference MATCH ~1e-6; async ~1.5x faster than naive\n\n# In transftran-ormers (study) — inspect any MoE family you plan to wire\nls src/transformers/models | grep -E 'moe|mixtral|qwen'  # 22+ MoE dirs + qwen/mixtral\npython3 -c \"from transformers import AutoConfig; print(AutoConfig.for_model('mixtral').__doc__[:400])\"", lang_label="verification · both repos")

doc.add_heading("9  Engineering Blueprint — How the 70B Runs When It Doesn't Fit in RAM", level=1)
add_para("This section files the two honest ways to run a 70B locally when it doesn't fit in RAM — done two ways, honestly (quoting the moe-async-engine README verbatim). Mix them up and every benchmark looks confusing.", size=8.5, color="475569", italic=True, space_after=4)
add_table(
    ["", "Dense Model (e.g., Llama-3-70B, Qwen2.5-72B)", "MoE Model (e.g., Mixtral, Qwen3-30B-A3B, gpt-oss 120B)"],
    [
        ["Weights used per token", "ALL of them", "Only top-k experts (a small slice)"],
        ["Can you keep most weights on disk?", "Only via mmap paging; every token still touches all weights → slow", "Yes — most experts sit idle on disk; only the active few are read"],
        ["Right tool in this filing", "Path A: LM Studio + GGUF + mmap", "Path B: async MoE offload engine (or Path A with a GGUF)"],
        ["Knowledge slicing?", "No — \"cut the knowledge out\" fails; every fact is spread across all layers", "Yes — real: cold experts are genuine no-ops for this token"],
    ],
    col_widths=[1.6, 2.8, 2.8], header_color="1E3A5F"
)
doc.add_paragraph().paragraph_format.space_after = Pt(3)
doc.add_heading("Path A — Run a 70B in LM Studio (practical, today)", level=2)
add_para("LM Studio runs llama.cpp, which already memory-maps (mmap) GGUF weights: the OS keeps the active working set in RAM and pages the rest from disk. That is the “weights on the drive, inference memory intact” behavior — no custom code. The filing recommends this for every dense 70B.", size=8.5, color="334155", space_after=3)
add_para("Steps filed from lm-studio-70b/README.md:", bold=True, size=8.5, color="0C1424", space_after=2)
add_para("Figure out what fits: python3 lm-studio-70b/fit_advisor.py --params 70 --ram 32 --disk 32  (add --disk-bw 0.5 for SATA SSD, 0.15 for HDD). It prints which GGUF quant fits your RAM + drive and a rough tokens/sec.", bullet="number", size=8, color="334155", space_after=2)
add_para("Download exactly that one quant: pip install -U 'huggingface_hub[cli,hf_transfer]' then ./lm-studio-70b/download_gguf.sh bartowski/Meta-Llama-3.1-70B-Instruct-GGUF 'Meta-Llama-3.1-70B-Instruct-IQ3_XXS.gguf'.", bullet="number", size=8, color="334155", space_after=2)
add_para("Load in LM Studio with: mmap on, context 4k–8k, GPU offload as high as VRAM allows, mlock off unless the whole file fits in RAM.", bullet="number", size=8, color="334155", space_after=2)
add_para("Honest sizing for your numbers (32 GB RAM + 32 GB drive = binding: only Q2_K / IQ3_XXS ~25 GB fits; it will run but is low quality. A 30–34B at Q5_K_M / Q6_K often beats a heavily crushed 70B — compare with --params 32 --ram 32 --disk 200).", size=7.5, color="7C2D12", space_after=3)
add_code_block("python3 fit_advisor.py --params 70 --ram 32 --disk 32\npython3 fit_advisor.py --params 32 --ram 32 --disk 200   # compare: near-lossless smaller model\n./download_gguf.sh bartowski/Meta-Llama-3.1-70B-Instruct-GGUF \\\n    'Meta-Llama-3.1-70B-Instruct-IQ3_XXS.gguf'", lang_label="path A · lm-studio-70b")
doc.add_heading("Path B — Async expert-parallel MoE decode engine (custom, prototype-verified)", level=2)
add_para("Runnable prototype that keeps router and attention pinned in RAM and streams experts from disk through a prefetching, cached, worker-thread scheduler so the compute thread rarely waits on a read. Replaces AirLLM’s slow sequential layer-by-layer streaming with compute/communication overlap.", size=8.5, color="334155", space_after=3)
add_table(
    ["File", "Role", "Key Guarantee"],
    [
        ["scheduler.py", "Concurrency core: thread-safe LRU cache + background prefetch workers + staging buffer. Torch-independent. request() non-blocking, get() blocks only if not ready", "Measures resident_hits, prefetch_hits, stalls, cold_misses, stall_seconds — never hides stalls"],
        ["moe.py", "Small MoE in torch.nn + per-expert safetensors sharding: each expert → its own file; load_expert via safe_open reads ~one expert’s bytes regardless of model size", "Functional expert_forward via F.linear so streamed expert splices into graph with no module reconstruction"],
        ["engine.py", "MoEEngine: resident weights pinned in RAM, KV cache, speculative prefetch (predict L+1 from previous token), naive vs async decode", "Async facade: aprefetch / aget for callers on an asyncio loop — but parallelism lives in threads that release the GIL"],
        ["demo.py", "Correctness + timing: shards synthetic MoE, checks async vs reference ~1e-6, times naive vs async with injected NVMe latency", "Demo PASS requires MATCH + speedup >1.2x"],
        ["test_overlap.py", "Torch-free proof of fetch/compute overlap (fake loader/compute with GIL-releasing sleeps)", "Hits ideal full-overlap bound: ~1.9x speedup, ~94% hit rate"],
        ["trace_sim.py", "Models routing statistics (Zipfian skew + temporal correlation 0.9) vs cache capacity", "Shows speedup rises to ~1.8x with correlation; ~1.3x without — monetizes the right statistic"],
    ],
    col_widths=[1.25, 3.1, 2.85], header_color="0C1424"
)
doc.add_paragraph().paragraph_format.space_after = Pt(3)
add_para("Why threads, not asyncio (the filing’s most-reviewed paragraph — quoted from scheduler.py): asyncio multiplexes coroutines on one thread and does not overlap CPU-bound work. Here both safetensors read and torch ATen kernels release the GIL, so a background thread doing the fetch genuinely runs in parallel with the main thread computing. Worker-thread pool is the correct realization of “async prefetching thread”; engine.py adds an async def facade over it.", size=7.5, color="334155", italic=True, space_after=3)
add_para("Honest stall bound (what “zero-stall” can and cannot mean): In autoregressive decode, layer L’s top-k experts are only known after you compute layer L’s router — expert selection is data-dependent, so within a single token’s critical path you cannot prefetch the next layer’s actual experts before you get there. The engine drives stalls toward zero with LRU hot-expert cache, speculative next-layer prefetch (temporal correlation), and within-layer parallel loads — and reports the residual. Representative demo: random weights (no locality) → ~50% hit rate, ~1.5x speedup; trace_sim with realistic correlation 0.9 → ~90% hit rate, ~1.8x, stalls ~6x lower.", size=7.5, color="7C2D12", space_after=3)
doc.add_heading("Scaling the prototype to a real 70B MoE (checklist filed for reviewers)", level=3)
checks = [
    "Model + weights: Swap synthetic MoEModel for the real architecture class from transftran-ormers (e.g., Mixtral/Qwen-MoE/DeepSeek-MoE) and load published weights. Keep shard layout: attention/router/norms resident, each expert its own quantized shard.",
    "Cache capacity: capacity ≈ free_RAM_bytes / bytes_per_expert. Must exceed per-step working set (top_k × lookahead layers) or cache thrashes — see trace_sim.py cap=2 rows.",
    "Faster reads: For Gen5 NVMe, read shards into pinned host memory and, if GPU, overlap host→device copy with a CUDA stream; consider O_DIRECT / io_uring to bypass page cache for cold experts.",
    "Better prediction: Prototype predicts L+1 from previous token. Stronger: tiny learned “which expert next” head, or run next layer’s router early on a cheap input approximation. Each improves prefetch hit rate — the dominant lever.",
    "GPU: Move resident + expert cache to VRAM; scheduler API unchanged — disk tier becomes CPU RAM or NVMe feeding VRAM via same request/get/prefetch.",
]
for c in checks:
    add_para(c, bullet="bullet", size=8, color="334155", space_after=2)

doc.add_heading("10  Deployment & Operations — cortex70b-arena-o3za.arcada.app", level=1)
add_para("Live deployment verified 2026-09-07. The filing snapshots the runtime you can re-verify with curl:", size=8.5, color="475569", italic=True, space_after=3)
add_table(
    ["Surface", "Route / File", "Status on 2026-09-07", "Notes"],
    [
        ["Frontend", "/  (Vite + React 19.2.8)", "200, Cloudflare, x-vercel-cache HIT", "Title: Cortex70B — Same Thoughts, 70B Scale + Browser Read Skills"],
        ["Thoughts", "GET /api/thoughts · /api/thoughts?limit=24", "200 — 3 seeded thoughts", "DeepSeek, Qwen2.5, Llama 3.3 traces with shared STEP 1 + STEP 4"],
        ["Models", "GET /api/models", "200 — 6 checkpoints", "Ranked 70B-class table rendered above"],
        ["Skills", "GET /api/skills", "200 — 9 skills (8 read + 1+ act)", "Read-mode, args, returns, latency live"],
        ["Runs", "GET /api/runs · /api/runs?limit=20  + POST/DELETE /api/runs", "200 — 3 seeded runs", "Playground persists via POST {skill_name, url, output, status}"],
        ["Identity", "GET/PUT/POST /api/identity", "200 — Cortex CORTEX-70B readonly true USB-C", "Toggle drives /self read-only banner"],
        ["Memories", "GET /api/memories?limit=100  + POST/PUT /api/memories", "200 — 2 seeded memories", "Folders /memories/core, /hardware; keep/pinned flags"],
        ["Hardware", "GET/POST/PUT/DELETE /api/usb", "200 — SKILLS-01 mounted, ARCHIVE-02 cold", "Connector → burst map enforced"],
        ["Repos", "GET/POST/PUT/DELETE /api/repos", "200 — 2 connected read-only", "moe-async-engine + transftran-ormers"],
        ["Plans / Analytics", "GET /api/plans · POST /api/agon/page-views", "200", "Arena recording via rrweb + viewerId"],
    ],
    col_widths=[1.1, 1.9, 1.6, 1.6], header_color="0C1424"
)
doc.add_paragraph().paragraph_format.space_after = Pt(2)
add_code_block("curl -s https://cortex70b-arena-o3za.arcada.app/api/identity | python3 -m json.tool\ncurl -s https://cortex70b-arena-o3za.arcada.app/api/memories?limit=100 | python3 -m json.tool\ncurl -s https://cortex70b-arena-o3za.arcada.app/api/skills | python3 -m json.tool\ncurl -s https://cortex70b-arena-o3za.arcada.app/api/models | python3 -m json.tool", lang_label="verify live — 2026-09-07")
add_para("Ops note: The bundle embeds arena recording (rrweb + analytics: interactions, cursor_path, click_heatmap, max_scroll_depth, time_to_first_click) and page-view beacons to designarena.ai/api/agon/page-views. This is the “arena” tournament harness (tournament feded6c8-c613…, generation muse-spark-1.3) and does not affect the filing’s functional guarantees.", size=7.5, color="64748B", italic=True, space_after=6)

doc.add_heading("11  Compliance, Safety & Reviewer Notes", level=1)
items = [
    ("Read-only is auditable, not aspirational:", "Filers must show enforcement at three layers — identity flag, USB mount mode, repo access — all three read true in the live snapshot. The skill registry’s read_mode strings are literal (“READ a11y tree”, “READ markdown”…); an auditor can grep them."),
    ("Direct-rate claim is bounded:", "Filing lists the exact connector→Gbps→lanes table from the bundle (not marketing copy). SKILLS-01 on USB-C is 10 Gbps x4; claiming Thunderbolt 4 (40 Gbps x8) requires remounting and re-verifying via GET /api/usb."),
    ("Memory curation is non-repudiation:", "Every saved memory is POSTed with {folder, title, content, keep} and visible on /self filtered by folder. Pinned charter entries are marked pinned:true — they survive prunes. Reviewers should spot-check that kept memories cite a ref or repo path per principle 2."),
    ("Dense vs. MoE honesty:", "The filing explicitly states you cannot offload experts from a dense 70B. Any benchmark claiming “MoE offload speedup” on Llama 3.3 70B dense must be rejected — the correct path is Path A (GGUF+mmap). MoE speedup claims require naming the MoE checkpoint and showing its router statistics."),
    ("Licenses:", "Transformers base is Apache-2.0; DeepSeek R1 Distill inherits Llama 3.3 license; gpt-oss 120B and Qwen3 MoE are Apache-2.0 — permissive for commercial agents. CI and CODE_OF_CONDUCT are retained from the upstream fork."),
]
for title, body in items:
    add_rich_para([
        (title + " ", {"bold":True, "size":8.5, "color":"0C1424"}),
        (body, {"size":8.5, "color":"334155"}),
    ], bullet="bullet", space_after=3)

doc.add_heading("12  Filing Checklist — Hand This Page to the Reviewer", level=1)
add_para("One-page attestation that the filed Cortex70B build matches the live deployment. Check each box against the GET endpoints noted.", size=8.5, color="475569", italic=True, space_after=4)
# checklist table with checkbox char
checks = [
    ["☐", "Identity exists and is read-only when learning", "GET /api/identity → readonly_learning:true, usb_connector:USB-C, name:Cortex codename:CORTEX-70B", "2026-09-07"],
    ["☐", "Skills volume is direct-rate and read-only", "GET /api/usb → SKILLS-01 USB-C 10 Gbps x4, readonly:true, mounted:true, skill_scope read/observe/extract", "2026-09-07"],
    ["☐", "Own memory folder with curatorial control", "GET /api/memories?limit=100 → /memories/core + /hardware seeded; POST /api/memories creates in /memories/learnings with keep flag", "2026-09-07"],
    ["☐", "Browser skills are read-first (≥8 read skills)", "GET /api/skills → 8 READ skills + 1+ act-on-read-refs; /skills page shows C of N read-capable", "2026-09-07"],
    ["☐", "GitHub repos connected read-only", "GET /api/repos → moe-async-engine true read-only + transftran-ormers true read-only; clone targets study/<name>", "2026-09-07"],
    ["☐", "Same thoughts — 70B scale demonstrated", "GET /api/thoughts → 3 models, shared STEP 1 ground-by-reading, STEP 4 READ before ACT; GET /api/models → 6 checkpoints ranked", "2026-09-07"],
    ["☐", "70B fits analysis honest (dense vs MoE)", "moe-async-engine/README.md table + lm-studio-70b/fit_advisor.py + moe-async-engine/moe.py sharding", "commit 063c6e6"],
    ["☐", "transftran-ormers help assessed and scoped", "§8 verdict: YES as model-definition layer; NO as standalone offload executor; integration path filed with 5 steps + limits", "main @ 254e62f"],
    ["☐", "Deployment live and verifiable", "GET https://cortex70b-arena-o3za.arcada.app → 200; all /api/* routes above return 200", "2026-09-07"],
]
add_table(
    ["✓", "Requirement", "Evidence (endpoint / file)", "As Of"],
    checks,
    col_widths=[0.4, 2.15, 3.4, 0.75], header_color="0C1424"
)
doc.add_paragraph().paragraph_format.space_after = Pt(4)
add_rich_para([
    ("Filing sign-off: ", {"bold":True, "size":9, "color":"0C1424"}),
    ("The agent described herein is filed as a read-first browser agent with the same thoughts across 70B-class models, direct-rate USB skill access, a sovereign memory vault, and a named identity — studied read-only against your two GitHub repos. It is live, not hypothetical, at the deployment above. Keep the next iteration on a feature branch, re-verify the four live GETs, and bring this page to review.", {"size":9, "color":"334155"}),
], space_after=6)
# signature line
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.LEFT
p.paragraph_format.space_before = Pt(12)
p.paragraph_format.space_after = Pt(0)
r = p.add_run("Filed by   ")
r.font.size = Pt(8)
r.font.color.rgb = RGBColor.from_string("64748B")
r = p.add_run("Cortex  ·  CORTEX-70B")
r.bold = True
r.font.size = Pt(9)
r.font.color.rgb = RGBColor.from_string("0C1424")
r = p.add_run("   for   chrisfbaileycb-arch  ·  2026-09-07 UTC  ·  Filing v1.0  —  cortex70b-arena-o3za.arcada.app")
r.font.size = Pt(8)
r.font.color.rgb = RGBColor.from_string("64748B")
# line
p2 = doc.add_paragraph()
p2.paragraph_format.space_before = Pt(4)
p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
r = p2.add_run("─" * 92)
r.font.size = Pt(6)
r.font.color.rgb = RGBColor.from_string("E2E8F0")
# footer note
p3 = doc.add_paragraph()
p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
p3.paragraph_format.space_before = Pt(2)
r = p3.add_run("This filing is a snapshot of a live system. Re-run the four curl GETs in §10 before any downstream decision — connector, mounts, model scores, and thought traces are versioned and can change.")
r.font.size = Pt(6.5)
r.font.color.rgb = RGBColor.from_string("94A3B8")
r.italic = True

# ============================================================
# APPENDICES
doc.add_heading("Appendix A — Live API Snapshot (abridged, 2026-09-07)", level=1)
add_para("Exact JSON returned by the deployment on the filing date — truncated for print, fully available via the curl commands.", size=8, color="64748B", italic=True, space_after=3)
add_code_block("GET /api/identity\n{\n  \"id\": 1,\n  \"name\": \"Cortex\",\n  \"codename\": \"CORTEX-70B\",\n  \"role\": \"Read-first browser agent · 70B-class reasoning\",\n  \"principles\": \"1. Read-only when self-learning ... 4. Ask before acting ...\",\n  \"readonly_learning\": true,\n  \"usb_connector\": \"USB-C\"\n}\n\nGET /api/usb  [2 volumes]\n[{\"label\":\"SKILLS-01\",\"connector\":\"USB-C\",\"mount_path\":\"/mnt/skills-01\",\n  \"skill_scope\":\"read, observe, extract\",\"readonly\":true,\"mounted\":true},\n {\"label\":\"ARCHIVE-02\",\"connector\":\"USB-A 3.0\",\"mount_path\":\"/mnt/archive-02\",\n  \"skill_scope\":\"trace, screenshot\",\"readonly\":true,\"mounted\":false}]", lang_label="api · identity + usb")
add_code_block("GET /api/repos  [2 repos, both connected read-only]\n[{\"name\":\"moe-async-engine\",\"url\":\"https://github.com/chrisfbaileycb-arch/moe-async-engine-.git\",\n  \"branch\":\"main\",\"access\":\"read-only\",\"connected\":true},\n {\"name\":\"transftran-ormers\",\"url\":\"https://github.com/chrisfbaileycb-arch/transftran-ormers\",\n  \"branch\":\"main\",\"access\":\"read-only\",\"connected\":true}]", lang_label="api · repos")
add_code_block("GET /api/skills  [9 skills] → snapshot, read, observe, extract, locate, trace, screenshot all READ*; click acts-on-read-refs\nGET /api/thoughts [3] → DeepSeek R1 Distill 1180 tok/2140ms, Qwen2.5 72B 940/1620, Llama 3.3 870/1480 — shared STEP 1 + STEP 4\nGET /api/models [6] → DeepSeek 94.5/92.1/88.5, gpt-oss 120B 91.8/90.4/87.2, Qwen2.5 72B 90.7/91.3/85.9, Qwen3 MoE 88.9/84.2/79.6 ...", lang_label="api · skills / thoughts / models (abridged)")
doc.add_heading("Appendix B — Repo Layouts at Filing Time", level=1)
add_code_block("moe-async-engine-/   (origin: github.com/chrisfbaileycb-arch/moe-async-engine-.git @063c6e6)\n├── lm-studio-70b/\n│   ├── fit_advisor.py            # which quant fits RAM+disk → GGUF picker\n│   ├── download_gguf.sh          # fetch single quant to LM Studio models folder\n│   ├── accelerate_offload_infer.py # AirLLM-free raw-weights offload (offline, slow)\n│   └── README.md                 # mmap truth, load settings, dense vs MoE explainer\n└── moe-async-engine/\n    ├── scheduler.py              # LRU cache + prefetch workers (GIL-aware threads)\n    ├── moe.py                    # MoE + per-expert safetensors sharding\n    ├── engine.py                 # resident decode + speculative prefetch + naive baseline\n    ├── demo.py / test_overlap.py / trace_sim.py\n    └── README.md                 # scaling guide to 70B+ MoE", lang_label="moe-async-engine-")
add_code_block("transftran-ormers/   (origin: github.com/chrisfbaileycb-arch/transftran-ormers @254e62f, fork of huggingface/transformers)\n├── src/transformers/models/    # 22+ MoE families: mixtral, qwen2_moe, qwen3_moe, olmoe, ...\n│   ├── mixtral/  qwen3_moe/  qwen3_5_moe/  phimoe/  jetmoe/  granitemoe/  glm4_moe/  afmoe/ ...\n├── src/transformers/modeling_utils.py, configuration_utils.py, pipelines/, quantizers/, ...\n├── examples/  docs/  tests/  benchmark/  docker/  utils/\n└── (working tree clean, branch main up to date with origin — ready for cortex/async-expert-shard branch)", lang_label="transftran-ormers")

# Finalize
# Add footer to section
section = doc.sections[0]
footer = section.footer
footer.is_linked_to_previous = False
fp = footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = fp.add_run("CORTEX-70B  ·  Same Thoughts · Read Skills  ·  Filed 2026-09-07  ·  cortex70b-arena-o3za.arcada.app  ·  Read-only when self-learning  ·  USB-C direct-rate  ·  /memories private vault")
r.font.size = Pt(6)
r.font.color.rgb = RGBColor.from_string("94A3B8")
r.font.name = 'Calibri'
r.italic = True

# Save
out = "Cortex70B_Filing_Same_Thoughts_ReadOnly_USB_Memory_Identity_2026-09-07.docx"
doc.save(out)
print(f"Saved {out}")
# verify
from docx import Document as D2
d2 = D2(out)
paras = [p.text for p in d2.paragraphs if p.text.strip()]
print(f"Paragraphs: {len(paras)}")
print(f"Tables: {len(d2.tables)}")
# show headings
heads = [p.text for p in d2.paragraphs if p.style.name.startswith('Heading')]
print("Headings found:", heads[:20])
# sections
print("Sections:", len(d2.sections))
print("Done verify")

