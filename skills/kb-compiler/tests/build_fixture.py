#!/usr/bin/env python3
"""
build_fixture.py — Builds a synthetic knowledge base with exactly one instance
of every problem check_master_list.py is supposed to detect, plus clean cases
that must stay clean.

The point is that "the checker works" stops being a belief and becomes
something you run. Every case below is deliberate; the expected result for
each one is declared in run_tests.py.

Usage:
    python3 build_fixture.py <output_dir>
"""

import datetime
import os
import sys
import unicodedata

import docx
import openpyxl
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

TITLE_COLOR = RGBColor(0x1F, 0x38, 0x64)
CALLOUT_TEXT_COLOR = RGBColor(0x7B, 0x5C, 0x00)
BASE_DATE = datetime.datetime(2026, 8, 1)


def add_footer(d, label="Fixture Team"):
    """Conforming footer: '<label>  |  Knowledge Base<TAB>Page <n> of <total>'."""
    p = d.sections[0].footer.paragraphs[0]
    p.text = ""
    p.add_run("%s  |  Knowledge Base\tPage " % label)
    for instr in (" PAGE ", " NUMPAGES "):
        r = p.add_run()
        begin = OxmlElement("w:fldChar"); begin.set(qn("w:fldCharType"), "begin")
        it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = instr
        sep = OxmlElement("w:fldChar"); sep.set(qn("w:fldCharType"), "separate")
        txt = OxmlElement("w:t"); txt.text = "1"
        end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end")
        for el in (begin, it, sep, txt, end):
            r._r.append(el)
        if instr == " PAGE ":
            p.add_run(" of ")


def new_doc(margin_inches=0.75, footer=True, landscape=False):
    d = docx.Document()
    s = d.sections[0]
    s.top_margin = s.bottom_margin = Inches(margin_inches)
    s.left_margin = s.right_margin = Inches(margin_inches)
    if landscape:
        from docx.enum.section import WD_ORIENT
        s.orientation = WD_ORIENT.LANDSCAPE
        s.page_width, s.page_height = s.page_height, s.page_width
    if footer:
        add_footer(d)
    return d


def add_title(d, text, bold=True, size_pt=22, color=TITLE_COLOR, split_runs=False):
    p = d.add_paragraph()
    chunks = [text[: len(text) // 2], text[len(text) // 2 :]] if split_runs else [text]
    for i, chunk in enumerate(chunks):
        r = p.add_run(chunk)
        # When split_runs is on, only the first run gets the correct formatting —
        # that's the "title broken across runs" case.
        if not split_runs or i == 0:
            r.bold = bold
            r.font.size = Pt(size_pt)
            r.font.color.rgb = color
    return p


def add_callout(d, text="⚠ Aviso", shading="FFF3CD", color=CALLOUT_TEXT_COLOR):
    p = d.add_paragraph()
    r = p.add_run(text)
    if color is not None:
        r.font.color.rgb = color
    if shading is not None:
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:fill"), shading)
        p._p.get_or_add_pPr().append(shd)
    return p


def add_body_sections(d, lists=True):
    """Emits the required Section 3 headings that come before the history."""
    d.add_heading("Purpose", 1)
    d.add_heading("Prerequisites", 1)
    if lists:
        d.add_paragraph("A prerequisite", style="List Bullet")
    d.add_heading("Step by Step", 1)
    if lists:
        d.add_paragraph("Do the thing", style="List Number")


def add_history(d, rows, legacy=False, heading=True):
    if heading:
        d.add_heading("Version History", 1)
    cols = (
        ["No.", "Revision Date", "Revision", "Reviewer"]
        if legacy
        else ["Author", "Version", "Date", "Change Description"]
    )
    t = d.add_table(rows=1 + len(rows), cols=4)
    for i, v in enumerate(cols):
        t.rows[0].cells[i].text = v
    for r, row in enumerate(rows, 1):
        for i, v in enumerate(row):
            t.rows[r].cells[i].text = v
    return t


def build(out):
    ident = os.path.join(out, "01 - Identity and Access")
    infra = os.path.join(out, "03 - Infrastructure")
    legacy_dir = os.path.join(out, "Legacy")
    attach = os.path.join(infra, "KB-017 - Autopilot - Files", "GetAutoPilot")
    for path in (ident, infra, legacy_dir, attach):
        os.makedirs(path, exist_ok=True)

    hist = [["Lucas Souza", "1.0", "Set/2026", "Criação"]]

    # --- CLEAN: must produce zero problems ---
    d = new_doc()
    d.sections[0].header.paragraphs[0].text = "KB-009 - Configurar MFA\tv1.0"
    add_title(d, "KB-009 - Configurar MFA")
    d.add_heading("Purpose", 1)
    d.add_heading("Prerequisites", 1)
    d.add_paragraph("A prerequisite", style="List Bullet")
    d.add_heading("Step by Step", 1)
    d.add_paragraph("Do the thing", style="List Number")
    add_callout(d)
    d.add_heading("Verification", 1)
    d.add_paragraph("It worked", style="List Bullet")
    add_history(d, hist)
    d.save(os.path.join(ident, "KB-009 - Configurar MFA.docx"))

    # --- name divergence: body title differs from the other three sources ---
    d = new_doc()
    d.sections[0].header.paragraphs[0].text = "KB-010 - Troca de Senha\tv1.0"
    add_title(d, "KB-010 - Reset de Senha")
    add_body_sections(d)
    add_history(d, hist)
    d.save(os.path.join(ident, "KB-010 - Troca de Senha.docx"))

    # --- header missing tab, name otherwise consistent ---
    # Must report exactly ONE problem (missing tab), never a second one as a
    # bogus name divergence.
    d = new_doc()
    d.sections[0].header.paragraphs[0].text = "KB-011 - Configurar YubiKeyv1.0"
    add_title(d, "KB-011 - Configurar YubiKey")
    add_body_sections(d)
    add_history(d, hist)
    d.save(os.path.join(ident, "KB-011 - Configurar YubiKey.docx"))

    # --- version divergence between header/history/spreadsheet ---
    d = new_doc()
    d.sections[0].header.paragraphs[0].text = "KB-017 - Autopilot\tv1.0"
    add_title(d, "KB-017 - Autopilot")
    add_body_sections(d)
    add_history(d, [["Lucas Souza", "2.1", "Set/2026", "Ajuste"]])
    d.save(os.path.join(infra, "KB-017 - Autopilot.docx"))

    # --- legacy-format version history table ---
    d = new_doc()
    d.sections[0].header.paragraphs[0].text = "KB-020 - Movimentar Workstation\tv1.0"
    add_title(d, "KB-020 - Movimentar Workstation")
    add_body_sections(d)
    add_history(d, [["1", "01/02/2025", "A", "Fulano"]], legacy=True)
    d.save(os.path.join(infra, "KB-020 - Movimentar Workstation.docx"))

    # --- NFD-encoded accented file name (must NOT be a broken reference) ---
    d = new_doc()
    d.sections[0].header.paragraphs[0].text = "KB-021 - Software Instalação\tv1.0"
    add_title(d, "KB-021 - Software Instalação")
    add_body_sections(d)
    add_history(d, hist)
    d.save(
        os.path.join(infra, unicodedata.normalize("NFD", "KB-021 - Software Instalação.docx"))
    )

    # --- orphan: real .docx with no master-index row ---
    d = new_doc()
    d.sections[0].header.paragraphs[0].text = "KB-099 - Nao Indexado\tv1.0"
    add_title(d, "KB-099 - Nao Indexado")
    add_body_sections(d)
    add_history(d, hist)
    d.save(os.path.join(infra, "KB-099 - Nao Indexado.docx"))

    # --- legacy folder document, code "—" ---
    d = new_doc()
    d.sections[0].header.paragraphs[0].text = "Superseded Process\tv1.0"
    add_title(d, "Superseded Process")
    add_body_sections(d)
    add_history(d, hist)
    d.save(os.path.join(legacy_dir, "Superseded Process.docx"))

    # --- FORMATTING: wrong margin, wrong title size, wrong title color ---
    d = new_doc(margin_inches=1.0)
    d.sections[0].header.paragraphs[0].text = "KB-030 - Formatacao Errada\tv1.0"
    add_title(d, "KB-030 - Formatacao Errada", size_pt=14, color=RGBColor(0xFF, 0x00, 0x00))
    add_body_sections(d)
    add_history(d, hist)
    d.save(os.path.join(infra, "KB-030 - Formatacao Errada.docx"))

    # --- FORMATTING: callout with wrong shading and wrong text color ---
    d = new_doc()
    d.sections[0].header.paragraphs[0].text = "KB-031 - Callout Errado\tv1.0"
    add_title(d, "KB-031 - Callout Errado")
    add_body_sections(d)
    add_callout(d, shading="CCCCCC", color=RGBColor(0x00, 0x00, 0x00))
    add_history(d, hist)
    d.save(os.path.join(infra, "KB-031 - Callout Errado.docx"))

    # --- FORMATTING: title split across runs, second run unformatted ---
    d = new_doc()
    d.sections[0].header.paragraphs[0].text = "KB-032 - Titulo Quebrado\tv1.0"
    add_title(d, "KB-032 - Titulo Quebrado", split_runs=True)
    add_body_sections(d)
    add_history(d, hist)
    d.save(os.path.join(infra, "KB-032 - Titulo Quebrado.docx"))

    # --- CONSERVATIVE MODE: formatting inherited from the style, not the run.
    # Visually correct; must land in "could not verify"... except it doesn't:
    # the cascade resolves it, so this must produce ZERO violations.
    d = new_doc()
    style = d.styles["Normal"]
    style.font.size = Pt(22)
    style.font.bold = True
    style.font.color.rgb = TITLE_COLOR
    d.sections[0].header.paragraphs[0].text = "KB-033 - Herdado do Estilo\tv1.0"
    d.add_paragraph("KB-033 - Herdado do Estilo")
    add_body_sections(d)
    add_history(d, hist)
    d.save(os.path.join(infra, "KB-033 - Herdado do Estilo.docx"))

    # --- BUG: empty first paragraph swallowed the title comparison ---
    # The body title diverges from the other three sources, but a blank
    # paragraph on top used to make it read as "" and drop out of the check.
    d = new_doc()
    d.sections[0].header.paragraphs[0].text = "KB-034 - Nome Do Header\tv1.0"
    d.add_paragraph("")
    add_title(d, "KB-034 - Nome Totalmente Diferente")
    d.add_heading("Purpose", 1)
    d.add_heading("Prerequisites", 1)
    d.add_heading("Step by Step", 1)
    add_history(d, hist)
    d.save(os.path.join(infra, "KB-034 - Nome Do Header.docx"))

    # --- folder/tab mismatch: row lives on the Identity tab, file sits in
    # the Infrastructure folder. Invisible before, because the File column
    # only stores the bare name and trusts the tab to imply the folder.
    d = new_doc()
    d.sections[0].header.paragraphs[0].text = "KB-035 - Pasta Errada\tv1.0"
    add_title(d, "KB-035 - Pasta Errada")
    d.add_heading("Purpose", 1)
    d.add_heading("Prerequisites", 1)
    d.add_heading("Step by Step", 1)
    add_history(d, hist)
    d.save(os.path.join(infra, "KB-035 - Pasta Errada.docx"))

    # --- Section 3 structure: missing "Pré-requisitos" AND Version History
    # is not the last section. Two separate structural issues.
    d = new_doc()
    d.sections[0].header.paragraphs[0].text = "KB-036 - Estrutura Quebrada\tv1.0"
    add_title(d, "KB-036 - Estrutura Quebrada")
    d.add_heading("Purpose", 1)
    d.add_heading("Step by Step", 1)
    add_history(d, hist)
    d.add_heading("Closing Notes", 1)   # extra section AFTER the history
    d.save(os.path.join(infra, "KB-036 - Estrutura Quebrada.docx"))

    # --- NEGATIVE test: a document name that legitimately ends in something
    # version-shaped ("... Office v2"), with no tab in the header. Must NOT be
    # reported as a glued header, and must NOT invent a phantom version "2".
    d = new_doc()
    d.sections[0].header.paragraphs[0].text = "KB-037 - Migracao Office v2"
    add_title(d, "KB-037 - Migracao Office v2")
    d.add_heading("Purpose", 1)
    d.add_heading("Prerequisites", 1)
    d.add_heading("Step by Step", 1)
    add_history(d, hist)
    d.save(os.path.join(infra, "KB-037 - Migracao Office v2.docx"))

    # --- extra section that is NOT a violation: an allowed optional heading
    # plus a custom one, with every required section present and in order.
    d = new_doc()
    d.sections[0].header.paragraphs[0].text = "KB-038 - Secao Extra Permitida\tv1.0"
    add_title(d, "KB-038 - Secao Extra Permitida")
    d.add_heading("Purpose", 1)
    d.add_heading("Roles & Responsibilities", 1)
    d.add_heading("Prerequisites", 1)
    d.add_heading("Step by Step", 1)
    d.add_heading("Notes", 1)
    d.add_heading("Verification", 1)
    add_history(d, hist)
    d.save(os.path.join(infra, "KB-038 - Secao Extra Permitida.docx"))

    # --- Section 4: landscape orientation and no footer at all ---
    d = new_doc(footer=False, landscape=True)
    d.sections[0].header.paragraphs[0].text = "KB-040 - Paisagem Sem Rodape\tv1.0"
    add_title(d, "KB-040 - Paisagem Sem Rodape")
    add_body_sections(d)
    add_history(d, hist)
    d.save(os.path.join(infra, "KB-040 - Paisagem Sem Rodape.docx"))

    # --- Section 4: en dash in the header instead of a plain hyphen ---
    d = new_doc()
    d.sections[0].header.paragraphs[0].text = "KB-041 \u2013 Travessao No Header\tv1.0"
    add_title(d, "KB-041 - Travessao No Header")
    add_body_sections(d)
    add_history(d, hist)
    d.save(os.path.join(infra, "KB-041 - Travessao No Header.docx"))

    # --- Section 4: prerequisites and steps as plain paragraphs, not lists ---
    d = new_doc()
    d.sections[0].header.paragraphs[0].text = "KB-042 - Sem Listas\tv1.0"
    add_title(d, "KB-042 - Sem Listas")
    d.add_heading("Purpose", 1)
    d.add_heading("Prerequisites", 1)
    d.add_paragraph("Access to the VPN.")
    d.add_heading("Step by Step", 1)
    d.add_paragraph("Open the console and click Save.")
    add_history(d, hist)
    d.save(os.path.join(infra, "KB-042 - Sem Listas.docx"))

    # --- Section 4: author name in all caps ---
    d = new_doc()
    d.sections[0].header.paragraphs[0].text = "KB-043 - Autor Caixa Alta\tv1.0"
    add_title(d, "KB-043 - Autor Caixa Alta")
    add_body_sections(d)
    add_history(d, [["LUCAS SOUZA", "1.0", "Sep/2026", "Created"]])
    d.save(os.path.join(infra, "KB-043 - Autor Caixa Alta.docx"))

    # --- Section 4: a different footer label from the rest of the base ---
    d = new_doc(footer=False)
    add_footer(d, label="Another Team")
    d.sections[0].header.paragraphs[0].text = "KB-044 - Rodape Divergente\tv1.0"
    add_title(d, "KB-044 - Rodape Divergente")
    add_body_sections(d)
    add_history(d, hist)
    d.save(os.path.join(infra, "KB-044 - Rodape Divergente.docx"))

    # --- duplicate basename across folders: the row's tab must resolve WHICH
    # file it refers to. Both copies carry different content, so picking the
    # wrong one would surface as phantom name and version divergences; the
    # unindexed twin must then show up as an orphan.
    for folder, header, title in (
        (ident, "KB-045 - Guia Duplicado\tv1.0", "KB-045 - Guia Duplicado"),
        (infra, "KB-045 - Guia Duplicado Outro\tv9.9", "KB-045 - Guia Duplicado Outro"),
    ):
        d = new_doc()
        d.sections[0].header.paragraphs[0].text = header
        add_title(d, title)
        add_body_sections(d)
        add_history(d, hist)
        d.save(os.path.join(folder, "KB-045 - Guia Duplicado.docx"))

    # --- multilingual recognition: a fully conforming document whose section
    # headings and history columns are in another language (Portuguese here).
    # It must produce ZERO findings — if language recognition ever regressed,
    # this document would start reporting missing sections.
    d = new_doc()
    d.sections[0].header.paragraphs[0].text = "KB-039 - Documento em Portugues\tv1.0"
    add_title(d, "KB-039 - Documento em Portugues")
    d.add_heading("Objetivo", 1)
    d.add_heading("Pré-requisitos", 1)
    d.add_paragraph("Um pré-requisito", style="List Bullet")
    d.add_heading("Passo a Passo", 1)
    d.add_paragraph("Faça isto", style="List Number")
    d.add_heading("Verificação", 1)
    d.add_paragraph("Confirme aquilo", style="List Bullet")
    d.add_heading("Histórico de Versões", 1)
    t = d.add_table(rows=2, cols=4)
    for i, v in enumerate(["Autor", "Versão", "Data", "Descrição da Alteração"]):
        t.rows[0].cells[i].text = v
    for i, v in enumerate(["Lucas Souza", "1.0", "Set/2026", "Criação"]):
        t.rows[1].cells[i].text = v
    d.save(os.path.join(infra, "KB-039 - Documento em Portugues.docx"))

    # --- attachments: must never be flagged as orphans ---
    with open(os.path.join(attach, "Get-WindowsAutoPilotInfo.ps1"), "w") as f:
        f.write("# attachment")
    with open(os.path.join(attach, "GetAutoPilot.cmd"), "w") as f:
        f.write("@echo off")

    # --- master-index spreadsheet ---
    wb = openpyxl.Workbook()
    ov = wb.active
    ov.title = "Visão Geral"
    ov.cell(row=1, column=1, value="Knowledge Base — Fixture")

    data = {
        "Identity and Access": [
            ("KB-009", "Configurar MFA", "KB-009 - Configurar MFA.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            # "Última Atualização" as TEXT, to exercise parse_date_value()
            ("KB-010", "Troca de Senha", "KB-010 - Troca de Senha.docx", "Active", "1.0", BASE_DATE, "15/09/2026", "Lucas Souza"),
            ("KB-011", "Configurar YubiKey", "KB-011 - Configurar YubiKey.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            # folder/tab mismatch: file physically lives in Infraestrutura
            ("KB-045", "Guia Duplicado", "KB-045 - Guia Duplicado.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            ("KB-035", "Pasta Errada", "KB-035 - Pasta Errada.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            # broken reference: file does not exist
            ("KB-050", "Arquivo Fantasma", "KB-050 - Nao Existe.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
        ],
        "Infrastructure": [
            ("KB-017", "Autopilot", "KB-017 - Autopilot.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            ("KB-020", "Movimentar Workstation", "KB-020 - Movimentar Workstation.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            ("KB-021", "Software Instalação", unicodedata.normalize("NFC", "KB-021 - Software Instalação.docx"), "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            ("KB-030", "Formatacao Errada", "KB-030 - Formatacao Errada.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            ("KB-031", "Callout Errado", "KB-031 - Callout Errado.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            ("KB-032", "Titulo Quebrado", "KB-032 - Titulo Quebrado.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            ("KB-033", "Herdado do Estilo", "KB-033 - Herdado do Estilo.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            # duplicate code: KB-009 already used in the other sheet
            ("KB-034", "Nome Do Header", "KB-034 - Nome Do Header.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            ("KB-036", "Estrutura Quebrada", "KB-036 - Estrutura Quebrada.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            ("KB-037", "Migracao Office v2", "KB-037 - Migracao Office v2.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            ("KB-038", "Secao Extra Permitida", "KB-038 - Secao Extra Permitida.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            ("KB-040", "Paisagem Sem Rodape", "KB-040 - Paisagem Sem Rodape.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            ("KB-041", "Travessao No Header", "KB-041 - Travessao No Header.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            ("KB-042", "Sem Listas", "KB-042 - Sem Listas.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            ("KB-043", "Autor Caixa Alta", "KB-043 - Autor Caixa Alta.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            ("KB-044", "Rodape Divergente", "KB-044 - Rodape Divergente.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            ("KB-039", "Documento em Portugues", "KB-039 - Documento em Portugues.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
            # duplicate code: KB-009 already used in the other sheet
            ("KB-009", "Codigo Duplicado", "KB-009 - Configurar MFA.docx", "Active", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
        ],
        "Legacy": [
            ("—", "Superseded Process", "Legacy/Superseded Process.docx", "Legacy", "1.0", BASE_DATE, BASE_DATE, "Lucas Souza"),
        ],
    }

    headers = [
        "Code", "Document", "File", "Status",
        "Version", "Creation Date", "Last Updated", "Owner",
    ]
    for cat, rows in data.items():
        ws = wb.create_sheet(cat)
        ws.cell(row=1, column=1, value=f"Knowledge Base — {cat}")
        for i, h in enumerate(headers, 1):
            ws.cell(row=2, column=i, value=h)
        for j, row in enumerate(rows, start=3):
            for i, v in enumerate(row, 1):
                ws.cell(row=j, column=i, value=v)

    wb.save(os.path.join(out, "Lista Mestra - Fixture.xlsx"))
    return out


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    target = build(sys.argv[1])
    print(f"Fixture built at: {target}")
