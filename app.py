"""
PreencherDocumentos - Morais Engenharia e Construção
Geração automática de Declaração ART (Word) e Memorial (Excel) → PDF
Python 3.11 | Tkinter | python-docx | openpyxl | comtypes
"""

# PROTEÇÃO ANTI-LOOP — deve ser a PRIMEIRA coisa no arquivo
import multiprocessing
multiprocessing.freeze_support()

import os
import sys
import shutil
import zipfile
import threading
import datetime
import re
import copy
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ──────────────────────────────────────────────
# DADOS FIXOS DOS ENGENHEIROS
# ──────────────────────────────────────────────
ENGENHEIROS = {
    "FELIPE GUILHERME BERÇAN": {
        "cpf": "147.849.107-86",
        "crea": "1022722034D-GO",
        "assinatura": "FELIPE.png",
    },
    "CAIO ARAUJO BRAGA": {
        "cpf": "011.309.411-67",
        "crea": "CREA-GO",
        "assinatura": "CAIO.jpeg",
    },
    "JOÃO VITOR CABRAL DE MORAIS": {
        "cpf": "038.144.411-25",
        "crea": "CREA-GO",
        "assinatura": "JOÃO VITOR.jpg",
    },
    "JULIO CESAR GOMES DE MORAIS FILHO": {
        "cpf": "033.865.821-17",
        "crea": "CREA-GO",
        "assinatura": "JULIO CESAR.png",
    },
    "PAULA FLEURY DE MORAIS": {
        "cpf": "033.813.881-18",
        "crea": "CREA-GO",
        "assinatura": "PAULA.png",
    },
    "ISAAC NATAN SANTOS": {
        "cpf": "701.117.261-07",
        "crea": "CREA-GO",
        "assinatura": "ISAAC.png",
    },
}

# Texto substituto quando sistema de esgoto = SIM
TEXTO_ESGOTO_SIM = (
    "Quanto as instalações hidrossanitárias, tal sistema deve obedecer as "
    "premissas das normas NBR 5626:2020 e NBR 8160:1999."
)

# Parágrafos do Word que formam a região verde (esgoto) — índices 0-based
PARAGRAFOS_ESGOTO = [15, 16, 17, 18, 19, 20, 21, 22, 23]

# Checkboxes de esgoto no Excel (drawing XML shape names)
# AM70 = SIM  |  AP70 = NÃO
SHAPE_ESGOTO_SIM = "QO012,12.L0C0;L0C-34^"
SHAPE_ESGOTO_NAO = "QO012,22.L0C0;L0C-37^"

# ──────────────────────────────────────────────
# PALETA DE CORES (padrão Morais Engenharia)
# ──────────────────────────────────────────────
COR = {
    "bg":        "#1e2a3a",
    "bg_log":    "#131c26",
    "campo":     "#2a3f55",
    "botao":     "#2e86de",
    "barra":     "#4cd964",
    "texto":     "#ffffff",
    "subtexto":  "#90adc4",
    "log":       "#7ec8a0",
    "erro":      "#ff6b6b",
    "aviso":     "#ffd93d",
}


# ══════════════════════════════════════════════
# UTILITÁRIOS
# ══════════════════════════════════════════════

def resource_path(rel: str) -> str:
    """Resolve caminho correto dentro ou fora do .exe PyInstaller."""
    base = getattr(sys, "_MEIPASS", Path(__file__).parent)
    return str(Path(base) / rel)


def formatar_data_hoje() -> str:
    hoje = datetime.date.today()
    meses = ["janeiro","fevereiro","março","abril","maio","junho",
             "julho","agosto","setembro","outubro","novembro","dezembro"]
    return f"{hoje.day} de {meses[hoje.month-1]} de {hoje.year}"


def normalizar_endereco(logradouro: str, quadra_lote: str) -> str:
    """
    Produz string de endereço para nome de arquivo:
    remove RUA/AVENIDA/AV., substitui Quadra→QD e Lote→LT
    """
    addr = logradouro.strip()
    for prefix in ["AVENIDA ", "AVENUE ", "AV. ", "AV ", "RUA ", "R. ", "R "]:
        if addr.upper().startswith(prefix):
            addr = addr[len(prefix):]
            break
    addr = addr.strip()

    ql = quadra_lote.strip()
    ql = re.sub(r'\bQUADRA\b', 'QD', ql, flags=re.IGNORECASE)
    ql = re.sub(r'\bQD\.\b', 'QD', ql, flags=re.IGNORECASE)
    ql = re.sub(r'\bLOTE\b', 'LT', ql, flags=re.IGNORECASE)
    ql = re.sub(r'\bLT\.\b', 'LT', ql, flags=re.IGNORECASE)

    return f"{addr} {ql}".strip()


def nome_arquivo(tipo: str, num_casa: int, logradouro: str, quadra_lote: str) -> str:
    """
    Gera nome de arquivo sem extensão.
    tipo: 'DECLARAÇÃO ART' ou 'MEMORIAL'
    """
    end = normalizar_endereco(logradouro, quadra_lote)
    return f"{tipo} CS {num_casa} - {end}"


# ══════════════════════════════════════════════
# PROCESSAMENTO — WORD (Declaração ART)
# ══════════════════════════════════════════════

def _substituir_placeholder_para(para, placeholder: str, valor: str):
    """
    Substitui {placeholder} dentro de um parágrafo, consolidando runs
    fragmentados e convertendo cor para preto.
    """
    from docx.oxml.ns import qn
    from docx.shared import RGBColor

    full = "".join(r.text for r in para.runs)
    if placeholder not in full:
        return False

    novo = full.replace(placeholder, valor)

    # Preservar formatação do primeiro run, limpar demais
    if not para.runs:
        return False

    # Guardar formatação base do run que continha o placeholder
    base_run = None
    for r in para.runs:
        if placeholder in r.text or (base_run is None):
            base_run = r
            break

    # Apagar todos os runs
    for r in para.runs:
        r.text = ""

    # Reescrever no primeiro run
    first = para.runs[0] if para.runs else para.add_run()
    first.text = novo

    # Copiar formatação base se disponível
    if base_run:
        try:
            first.bold = base_run.bold
            first.italic = base_run.italic
            first.underline = base_run.underline
            first.font.size = base_run.font.size
            first.font.name = base_run.font.name
        except Exception:
            pass

    # Forçar cor preta
    first.font.color.rgb = RGBColor(0, 0, 0)
    return True


def _substituir_todos_runs(para, placeholder: str, valor: str):
    """
    Abordagem alternativa: varre run a run e remonta quando o placeholder
    está fragmentado entre múltiplos runs.
    """
    from docx.shared import RGBColor

    # Reconstruir texto completo
    texto = "".join(r.text for r in para.runs)
    if placeholder not in texto:
        return

    novo_texto = texto.replace(placeholder, valor)

    # Limpar todos os runs
    for run in para.runs:
        run.text = ""

    # Escrever no primeiro run disponível
    if para.runs:
        para.runs[0].text = novo_texto
        para.runs[0].font.color.rgb = RGBColor(0, 0, 0)


def preencher_word(
    template_path: str,
    saida_path: str,
    dados: dict,
    esgoto_sim: bool,
    log=None,
):
    """
    Preenche o template Word com os dados fornecidos.
    dados: dict com chaves {1}..{11}, {ENGENHEIRO SELECIONADO}, {dia/mes/ano}
    """
    from docx import Document
    from docx.shared import RGBColor, Inches, Pt
    from docx.oxml.ns import qn
    import lxml.etree as etree

    def _log(msg):
        if log:
            log(msg)

    _log("Carregando template Word...")
    doc = Document(template_path)

    # ── Mapeamento de substituições de texto ──
    substituicoes = {
        "{1}":                       dados.get("art", ""),
        "{2}":                       dados.get("crea", ""),
        "{5}":                       dados.get("logradouro", ""),
        "{6}":                       dados.get("quadra_lote", ""),
        "{7}":                       dados.get("bairro", ""),
        "{9}":                       dados.get("complemento", ""),
        "{10}":                      dados.get("cidade", ""),
        "{11}":                      dados.get("uf", ""),
        "{ENGENHEIRO SELECIONADO}":  dados.get("engenheiro_nome", ""),
        "{dia/mes/ano}":             formatar_data_hoje(),
    }

    _log("Substituindo campos de texto...")
    for i, para in enumerate(doc.paragraphs):
        # Região esgoto: se esgoto=SIM, substituir bloco verde pelo texto novo
        if i == PARAGRAFOS_ESGOTO[0] and esgoto_sim:
            # Apagar todos os parágrafos verdes (15..23) e inserir texto novo no 15
            _log("Aplicando substituição de texto de esgoto (SIM)...")
            # Primeiro parágrafo recebe o novo texto
            for r in doc.paragraphs[i].runs:
                r.text = ""
            if doc.paragraphs[i].runs:
                doc.paragraphs[i].runs[0].text = TEXTO_ESGOTO_SIM
                doc.paragraphs[i].runs[0].font.color.rgb = RGBColor(0, 0, 0)
                doc.paragraphs[i].runs[0].bold = True
            # Demais parágrafos da região verde: limpar
            for pi in PARAGRAFOS_ESGOTO[1:]:
                for r in doc.paragraphs[pi].runs:
                    r.text = ""
            continue

        if i in PARAGRAFOS_ESGOTO[1:] and esgoto_sim:
            continue  # já limpou acima

        # Substituições normais
        for placeholder, valor in substituicoes.items():
            _substituir_todos_runs(para, placeholder, valor)

    _log("Inserindo assinatura no Word...")
    _inserir_assinatura_word(doc, dados.get("assinatura_path", ""), log)

    _log(f"Salvando Word em: {saida_path}")
    doc.save(saida_path)
    _log("Word salvo com sucesso.")


def _inserir_assinatura_word(doc, img_path: str, log=None):
    """
    Insere imagem de assinatura no parágrafo 35 (entre parágrafos 34 e 36).
    O parágrafo 36 contém '____________________________'.
    Insere imagem como inline no parágrafo vazio antes da linha.
    """
    from docx.shared import Inches
    import os

    if not img_path or not os.path.exists(img_path):
        if log:
            log(f"⚠ Assinatura não encontrada: {img_path}")
        return

    # Parágrafo alvo: o que está antes da linha (______)
    # Linha está no índice 36; inserimos na 35 (vazio)
    try:
        target_para = doc.paragraphs[35]
        run = target_para.add_run()
        run.add_picture(img_path, width=Inches(1.5))
        if log:
            log("Assinatura inserida no Word.")
    except Exception as e:
        if log:
            log(f"⚠ Erro ao inserir assinatura Word: {e}")


# ══════════════════════════════════════════════
# PROCESSAMENTO — EXCEL (Memorial)
# ══════════════════════════════════════════════

def preencher_excel(
    template_path: str,
    saida_path: str,
    dados: dict,
    esgoto_sim: bool,
    assinatura_path: str,
    log=None,
):
    """
    Preenche o Memorial Excel com os dados fornecidos.
    Manipula checkboxes diretamente no XML do drawing.
    """
    import zipfile
    import shutil
    from xml.etree import ElementTree as ET
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    from openpyxl.drawing.image import Image as XLImage

    def _log(msg):
        if log:
            log(msg)

    _log("Copiando template Excel...")
    shutil.copy2(template_path, saida_path)

    # ── Passo 1: Preencher células com python via openpyxl ──
    _log("Preenchendo células do Excel...")
    wb = load_workbook(saida_path)
    ws = wb["ElemConstrutivos"]

    mapa_celulas = {
        "G40":  dados.get("contratante", ""),
        "G43":  dados.get("engenheiro_nome", ""),
        "AH43": dados.get("crea", ""),
        "AP43": "GO",
        "AR43": dados.get("cpf", ""),
        "G47":  dados.get("logradouro", ""),
        "AJ47": dados.get("quadra_lote", ""),
        "G49":  dados.get("bairro", ""),
        "V49":  dados.get("cep", ""),
        "AA49": dados.get("cidade", ""),
        "AU49": dados.get("uf", ""),
        "H53":  dados.get("engenheiro_nome", ""),
        "Y54":  dados.get("art", ""),
        "H75":  f"GOIÂNIA, {formatar_data_hoje()}",
        "AE77": dados.get("engenheiro_nome", ""),
        "AE78": dados.get("cpf", ""),
        "AE79": dados.get("crea", ""),
    }

    from openpyxl.styles import Font
    for coord, valor in mapa_celulas.items():
        cell = ws[coord]
        cell.value = valor
        # Preservar fonte existente, apenas mudar valor
        if cell.font:
            cell.font = Font(
                name=cell.font.name,
                size=cell.font.size,
                bold=cell.font.bold,
                italic=cell.font.italic,
                color="000000",
            )

    wb.save(saida_path)
    _log("Células preenchidas.")

    # ── Passo 2: Manipular checkboxes de esgoto no XML ──
    _log("Configurando checkboxes de esgoto no XML...")
    _ajustar_checkbox_esgoto(saida_path, esgoto_sim, log)

    # ── Passo 3: Inserir assinatura como imagem ──
    _log("Inserindo assinatura no Excel...")
    _inserir_assinatura_excel(saida_path, assinatura_path, log)

    _log("Excel preenchido com sucesso.")


def _ajustar_checkbox_esgoto(xlsx_path: str, esgoto_sim: bool, log=None):
    """
    Modifica o drawing XML para marcar/desmarcar os checkboxes de esgoto.
    SIM marcado = solidFill preto | NÃO marcado = noFill
    """
    import zipfile, shutil, os
    from xml.etree import ElementTree as ET

    NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    NS_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"

    tmp_path = xlsx_path + ".tmp"
    shutil.copy2(xlsx_path, tmp_path)

    with zipfile.ZipFile(tmp_path, "r") as zin, \
         zipfile.ZipFile(xlsx_path, "w", zipfile.ZIP_DEFLATED) as zout:

        for item in zin.infolist():
            data = zin.read(item.filename)

            if item.filename == "xl/drawings/drawing1.xml":
                ET.register_namespace("xdr", NS_XDR)
                ET.register_namespace("a",   NS_A)

                root = ET.fromstring(data)

                for anchor in root.findall(f"{{{NS_XDR}}}twoCellAnchor"):
                    sp = anchor.find(f"{{{NS_XDR}}}sp")
                    if sp is None:
                        continue
                    cNvPr = sp.find(f".//{{{NS_XDR}}}cNvPr")
                    if cNvPr is None:
                        continue
                    name = cNvPr.get("name", "")

                    if name not in (SHAPE_ESGOTO_SIM, SHAPE_ESGOTO_NAO):
                        continue

                    spPr = sp.find(f"{{{NS_XDR}}}spPr")
                    if spPr is None:
                        continue

                    # Remover fill existente
                    for tag in [f"{{{NS_A}}}solidFill", f"{{{NS_A}}}noFill"]:
                        el = spPr.find(tag)
                        if el is not None:
                            spPr.remove(el)

                    # Inserir fill correto
                    if (name == SHAPE_ESGOTO_SIM and esgoto_sim) or \
                       (name == SHAPE_ESGOTO_NAO and not esgoto_sim):
                        # Marcar: solidFill preto
                        solid = ET.SubElement(spPr, f"{{{NS_A}}}solidFill")
                        clr   = ET.SubElement(solid, f"{{{NS_A}}}srgbClr")
                        clr.set("val", "000000")
                    else:
                        # Desmarcar: noFill
                        ET.SubElement(spPr, f"{{{NS_A}}}noFill")

                data = ET.tostring(root, encoding="UTF-8", xml_declaration=True)

            zout.writestr(item, data)

    os.remove(tmp_path)
    if log:
        log(f"  Checkbox esgoto → {'SIM marcado' if esgoto_sim else 'NÃO marcado'}")


def _inserir_assinatura_excel(xlsx_path: str, img_path: str, log=None):
    """
    Insere imagem de assinatura na região AE-AH rows 73-76 (acima de Nome/CPF/CREA).
    Usa openpyxl Image após reabrir o arquivo.
    """
    import os
    from openpyxl import load_workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.utils import get_column_letter

    if not img_path or not os.path.exists(img_path):
        if log:
            log(f"⚠ Assinatura não encontrada: {img_path}")
        return

    try:
        wb = load_workbook(xlsx_path)
        ws = wb["ElemConstrutivos"]
        img = XLImage(img_path)
        img.width  = 120
        img.height = 45
        img.anchor = "AE73"
        ws.add_image(img)
        wb.save(xlsx_path)
        if log:
            log("Assinatura inserida no Excel.")
    except Exception as e:
        if log:
            log(f"⚠ Erro ao inserir assinatura Excel: {e}")


# ══════════════════════════════════════════════
# EXPORTAÇÃO PARA PDF via COM (Office)
# ══════════════════════════════════════════════

def exportar_word_pdf(docx_path: str, pdf_path: str, log=None):
    """Exporta .docx para .pdf via Word COM automation."""
    try:
        import comtypes.client
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(docx_path))
        doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)  # 17 = wdFormatPDF
        doc.Close()
        word.Quit()
        if log:
            log(f"PDF Word gerado: {Path(pdf_path).name}")
    except Exception as e:
        if log:
            log(f"✗ Erro ao exportar Word PDF: {e}")
        raise


def exportar_excel_pdf(xlsx_path: str, pdf_path: str, log=None):
    """Exporta .xlsx para .pdf via Excel COM automation (1 página)."""
    try:
        import comtypes.client
        excel = comtypes.client.CreateObject("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))
        ws = wb.Worksheets("ElemConstrutivos")

        # Forçar escala para caber em 1 página
        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.FitToPagesTall = 1

        ws.ExportAsFixedFormat(
            0,                          # 0 = xlTypePDF
            os.path.abspath(pdf_path),
            1,                          # 1 = xlQualityStandard
            True,                       # IncludeDocProperties
            False,                      # IgnorePrintAreas
        )
        wb.Close(False)
        excel.Quit()
        if log:
            log(f"PDF Excel gerado: {Path(pdf_path).name}")
    except Exception as e:
        if log:
            log(f"✗ Erro ao exportar Excel PDF: {e}")
        raise


# ══════════════════════════════════════════════
# ORQUESTRADOR PRINCIPAL
# ══════════════════════════════════════════════

def processar(params: dict, step_cb=None, log=None):
    """
    Gera os documentos para cada casa.
    params: dict com todos os dados do formulário.
    step_cb(pct, desc): callback de progresso.
    log(msg): callback de log.
    """
    def _step(pct, desc):
        if step_cb:
            step_cb(pct, desc)

    def _log(msg):
        if log:
            log(msg)

    word_template   = params["word_template"]
    excel_template  = params["excel_template"]
    saida_dir       = params["saida_dir"]
    assinatura_path = params["assinatura_path"]
    esgoto_sim      = params["esgoto_sim"]
    casas           = params["casas"]   # lista de dicts: {num, logradouro}
    dados_base      = params["dados"]   # campos extraídos da ART + engenheiro

    total_steps = len(casas) * 4  # word + excel + pdf_word + pdf_excel
    step_atual  = 0

    for casa in casas:
        num      = casa["num"]
        logr     = casa["logradouro"]
        quadlote = dados_base["quadra_lote"]

        dados_casa = {**dados_base, "logradouro_casa": logr}
        dados_para_doc = {**dados_base}
        if logr != dados_base.get("logradouro", ""):
            dados_para_doc["logradouro"] = logr

        _log(f"\n{'='*40}")
        _log(f"Processando CASA {num} — {logr}")
        _log(f"{'='*40}")

        # ── Nomes de arquivo ──
        nome_decl  = nome_arquivo("DECLARAÇÃO ART", num, logr, quadlote)
        nome_mem   = nome_arquivo("MEMORIAL", num, logr, quadlote)

        docx_out  = os.path.join(saida_dir, nome_decl + ".docx")
        xlsx_out  = os.path.join(saida_dir, nome_mem  + ".xlsx")
        pdf_decl  = os.path.join(saida_dir, nome_decl + ".pdf")
        pdf_mem   = os.path.join(saida_dir, nome_mem  + ".pdf")

        # ── Word ──
        step_atual += 1
        pct = int(step_atual / total_steps * 100)
        _step(pct, f"Casa {num}: preenchendo Declaração ART...")
        preencher_word(
            word_template, docx_out,
            dados_para_doc, esgoto_sim, log=_log
        )

        # ── Excel ──
        step_atual += 1
        pct = int(step_atual / total_steps * 100)
        _step(pct, f"Casa {num}: preenchendo Memorial...")
        preencher_excel(
            excel_template, xlsx_out,
            dados_para_doc, esgoto_sim, assinatura_path, log=_log
        )

        # ── PDF Word ──
        step_atual += 1
        pct = int(step_atual / total_steps * 100)
        _step(pct, f"Casa {num}: exportando Declaração ART para PDF...")
        exportar_word_pdf(docx_out, pdf_decl, log=_log)

        # ── PDF Excel ──
        step_atual += 1
        pct = int(step_atual / total_steps * 100)
        _step(pct, f"Casa {num}: exportando Memorial para PDF...")
        exportar_excel_pdf(xlsx_out, pdf_mem, log=_log)

        _log(f"✓ Casa {num} concluída.")

    _step(100, "Processamento concluído!")
    _log("\n✓ Todos os documentos foram gerados com sucesso.")


# ══════════════════════════════════════════════
# INTERFACE TKINTER
# ══════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Morais Engenharia — Preenchimento de Documentos")
        self.configure(bg=COR["bg"])
        self.resizable(False, False)

        self._assinatura_dir = ""  # pasta onde estão as imagens de assinatura
        self._campos_ruas = []     # widgets dinâmicos de rua por casa

        self._build_ui()
        self._centralizar()

    # ──────────────────────────────────────────
    # BUILD UI
    # ──────────────────────────────────────────

    def _build_ui(self):
        PAD = 14

        # ── Cabeçalho ──
        hdr = tk.Frame(self, bg="#162030", pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text="MORAIS ENGENHARIA", font=("Segoe UI", 15, "bold"),
                 bg="#162030", fg=COR["texto"]).pack()
        tk.Label(hdr, text="Preenchimento Automático de Documentos",
                 font=("Segoe UI", 9), bg="#162030", fg=COR["subtexto"]).pack()

        # ── Corpo principal ──
        corpo = tk.Frame(self, bg=COR["bg"])
        corpo.pack(fill="both", padx=PAD, pady=PAD)

        col_esq = tk.Frame(corpo, bg=COR["bg"])
        col_dir = tk.Frame(corpo, bg=COR["bg"])
        col_esq.pack(side="left", fill="both", padx=(0, 8))
        col_dir.pack(side="left", fill="both")

        # ══ COLUNA ESQUERDA ══

        # Arquivos
        self._secao(col_esq, "ARQUIVOS")
        self.var_word     = self._campo_arquivo(col_esq, "Template Word (.docx):", "docx")
        self.var_excel    = self._campo_arquivo(col_esq, "Memorial Excel (.xls/.xlsx):", "excel")
        self.var_assindir = self._campo_arquivo(col_esq, "Pasta de assinaturas:", "dir")
        self.var_saida    = self._campo_arquivo(col_esq, "Pasta de saída:", "dir")

        # Engenheiro
        self._secao(col_esq, "ENGENHEIRO RESPONSÁVEL")
        self.var_eng = tk.StringVar()
        combo = ttk.Combobox(col_esq, textvariable=self.var_eng,
                             values=list(ENGENHEIROS.keys()),
                             state="readonly", width=44)
        combo.pack(fill="x", pady=(0, 6))
        combo.bind("<<ComboboxSelected>>", self._on_eng_select)

        self.lbl_cpf  = self._info_label(col_esq, "CPF: —")
        self.lbl_crea = self._info_label(col_esq, "CREA: —")

        # Dados da ART
        self._secao(col_esq, "DADOS DA ART")
        campos_art = [
            ("Número da ART {1}:",        "art"),
            ("Número de Registro CREA {2}:", "crea"),
            ("Contratante {4}:",           "contratante"),
            ("Logradouro da Obra {5}:",    "logradouro"),
            ("Quadra e Lote {6}:",         "quadra_lote"),
            ("Bairro {7}:",                "bairro"),
            ("Complemento {9}:",           "complemento"),
            ("CEP {8}:",                   "cep"),
            ("Cidade {10}:",               "cidade"),
            ("UF {11}:",                   "uf"),
        ]
        self.vars_art = {}
        for label, key in campos_art:
            self.vars_art[key] = self._campo_texto(col_esq, label)

        # ══ COLUNA DIREITA ══

        # Opções
        self._secao(col_dir, "OPÇÕES")

        self.var_esgoto = tk.BooleanVar(value=False)
        self._checkbox(col_dir, "Sistema público de esgoto (SIM)", self.var_esgoto)

        self.var_esquina = tk.BooleanVar(value=False)
        self._checkbox(col_dir, "Lote de esquina", self.var_esquina,
                       command=self._on_esquina_toggle)

        # Qtd casas
        frm_qtd = tk.Frame(col_dir, bg=COR["bg"])
        frm_qtd.pack(fill="x", pady=(4, 0))
        tk.Label(frm_qtd, text="Quantidade de casas:",
                 bg=COR["bg"], fg=COR["subtexto"],
                 font=("Segoe UI", 9)).pack(side="left")
        self.var_qtd_casas = tk.IntVar(value=1)
        spn = tk.Spinbox(frm_qtd, from_=1, to=20, width=5,
                         textvariable=self.var_qtd_casas,
                         bg=COR["campo"], fg=COR["texto"],
                         insertbackground=COR["texto"],
                         command=self._on_qtd_casas_change)
        spn.pack(side="left", padx=6)
        spn.bind("<FocusOut>", lambda e: self._on_qtd_casas_change())

        # Painel dinâmico de ruas por casa
        self._secao(col_dir, "RUAS POR CASA (esquina c/ ruas diferentes)")
        self.frm_ruas = tk.Frame(col_dir, bg=COR["bg"])
        self.frm_ruas.pack(fill="x")
        self.lbl_ruas_hint = tk.Label(
            self.frm_ruas,
            text="(Ativo apenas quando 'Lote de esquina' +\n casas em ruas diferentes)",
            bg=COR["bg"], fg=COR["subtexto"], font=("Segoe UI", 8)
        )
        self.lbl_ruas_hint.pack()

        # Esquina — mesma rua?
        self.var_mesma_rua = tk.BooleanVar(value=True)
        self.frm_esquina_opt = tk.Frame(col_dir, bg=COR["bg"])
        self.frm_esquina_opt.pack(fill="x", pady=(2, 0))
        self.rb_mesma = tk.Radiobutton(
            self.frm_esquina_opt, text="Todas as casas na mesma rua",
            variable=self.var_mesma_rua, value=True,
            bg=COR["bg"], fg=COR["texto"], selectcolor=COR["campo"],
            command=self._on_mesma_rua_toggle
        )
        self.rb_dif = tk.Radiobutton(
            self.frm_esquina_opt, text="Casas em ruas diferentes",
            variable=self.var_mesma_rua, value=False,
            bg=COR["bg"], fg=COR["texto"], selectcolor=COR["campo"],
            command=self._on_mesma_rua_toggle
        )
        self.rb_mesma.pack(anchor="w")
        self.rb_dif.pack(anchor="w")
        self.frm_esquina_opt.pack_forget()  # oculto até marcar esquina

        # Log
        self._secao(col_dir, "LOG DE EXECUÇÃO")
        self.txt_log = tk.Text(
            col_dir, height=10, width=52,
            bg=COR["bg_log"], fg=COR["log"],
            font=("Consolas", 8), relief="flat",
            state="disabled"
        )
        self.txt_log.pack(fill="x")

        # Barra de progresso
        self._secao(col_dir, "PROGRESSO")
        self.lbl_prog = tk.Label(col_dir, text="Aguardando...",
                                 bg=COR["bg"], fg=COR["subtexto"],
                                 font=("Segoe UI", 8))
        self.lbl_prog.pack(anchor="w")
        self.pb = ttk.Progressbar(col_dir, length=380, mode="determinate")
        self.pb.pack(fill="x", pady=(2, 8))

        style = ttk.Style()
        style.theme_use("default")
        style.configure("green.Horizontal.TProgressbar",
                        troughcolor=COR["campo"], background=COR["barra"])
        self.pb.configure(style="green.Horizontal.TProgressbar")

        # Botão
        self.btn = tk.Button(
            col_dir, text="⚡  GERAR DOCUMENTOS",
            font=("Segoe UI", 11, "bold"),
            bg=COR["botao"], fg=COR["texto"],
            activebackground="#1a6ab5", activeforeground=COR["texto"],
            relief="flat", pady=10, cursor="hand2",
            command=self._iniciar
        )
        self.btn.pack(fill="x", pady=(4, 0))

    # ──────────────────────────────────────────
    # HELPERS DE UI
    # ──────────────────────────────────────────

    def _secao(self, parent, titulo):
        tk.Label(parent, text=titulo,
                 bg=COR["bg"], fg=COR["subtexto"],
                 font=("Segoe UI", 8, "bold")).pack(anchor="w", pady=(10, 2))
        tk.Frame(parent, bg=COR["campo"], height=1).pack(fill="x", pady=(0, 4))

    def _info_label(self, parent, texto):
        lbl = tk.Label(parent, text=texto, bg=COR["bg"], fg=COR["subtexto"],
                       font=("Segoe UI", 8))
        lbl.pack(anchor="w")
        return lbl

    def _checkbox(self, parent, texto, var, command=None):
        tk.Checkbutton(
            parent, text=texto, variable=var,
            bg=COR["bg"], fg=COR["texto"],
            selectcolor=COR["campo"], activebackground=COR["bg"],
            font=("Segoe UI", 9),
            command=command
        ).pack(anchor="w", pady=2)

    def _campo_texto(self, parent, label, default=""):
        tk.Label(parent, text=label, bg=COR["bg"], fg=COR["subtexto"],
                 font=("Segoe UI", 8)).pack(anchor="w")
        var = tk.StringVar(value=default)
        tk.Entry(parent, textvariable=var, width=46,
                 bg=COR["campo"], fg=COR["texto"],
                 insertbackground=COR["texto"], relief="flat",
                 font=("Segoe UI", 9)).pack(fill="x", pady=(0, 4))
        return var

    def _campo_arquivo(self, parent, label, tipo):
        tk.Label(parent, text=label, bg=COR["bg"], fg=COR["subtexto"],
                 font=("Segoe UI", 8)).pack(anchor="w")
        frm = tk.Frame(parent, bg=COR["bg"])
        frm.pack(fill="x", pady=(0, 4))
        var = tk.StringVar()
        tk.Entry(frm, textvariable=var, width=36,
                 bg=COR["campo"], fg=COR["texto"],
                 insertbackground=COR["texto"], relief="flat",
                 font=("Segoe UI", 9)).pack(side="left", fill="x", expand=True)

        def _browse():
            if tipo == "dir":
                p = filedialog.askdirectory()
            elif tipo == "docx":
                p = filedialog.askopenfilename(
                    filetypes=[("Word", "*.docx"), ("Todos", "*.*")])
            else:  # excel
                p = filedialog.askopenfilename(
                    filetypes=[("Excel", "*.xls *.xlsx"), ("Todos", "*.*")])
            if p:
                var.set(p)
                if tipo == "dir" and label.startswith("Pasta de assinatura"):
                    self._assinatura_dir = p

        tk.Button(frm, text="...", bg=COR["campo"], fg=COR["texto"],
                  relief="flat", font=("Segoe UI", 9),
                  command=_browse).pack(side="left", padx=(4, 0))
        return var

    def _centralizar(self):
        self.update_idletasks()
        w = self.winfo_reqwidth()
        h = self.winfo_reqheight()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    # ──────────────────────────────────────────
    # EVENTOS
    # ──────────────────────────────────────────

    def _on_eng_select(self, event=None):
        eng = self.var_eng.get()
        if eng in ENGENHEIROS:
            info = ENGENHEIROS[eng]
            self.lbl_cpf.config(text=f"CPF: {info['cpf']}")
            self.lbl_crea.config(text=f"CREA: {info['crea']}")
            # Auto-preencher CREA no campo ART se ainda vazio
            if not self.vars_art["crea"].get():
                self.vars_art["crea"].set(info["crea"])

    def _on_esquina_toggle(self):
        if self.var_esquina.get():
            self.frm_esquina_opt.pack(fill="x", pady=(2, 0))
        else:
            self.frm_esquina_opt.pack_forget()
        self._on_mesma_rua_toggle()

    def _on_mesma_rua_toggle(self):
        # Mostrar/ocultar campos de rua
        ativa = self.var_esquina.get() and not self.var_mesma_rua.get()
        if ativa:
            self._rebuild_campos_ruas()
        else:
            self._limpar_campos_ruas()

    def _on_qtd_casas_change(self):
        if self.var_esquina.get() and not self.var_mesma_rua.get():
            self._rebuild_campos_ruas()

    def _limpar_campos_ruas(self):
        for w in self.frm_ruas.winfo_children():
            w.destroy()
        self.lbl_ruas_hint = tk.Label(
            self.frm_ruas,
            text="(Ativo apenas quando 'Lote de esquina' +\n casas em ruas diferentes)",
            bg=COR["bg"], fg=COR["subtexto"], font=("Segoe UI", 8)
        )
        self.lbl_ruas_hint.pack()
        self._campos_ruas = []

    def _rebuild_campos_ruas(self):
        for w in self.frm_ruas.winfo_children():
            w.destroy()
        self._campos_ruas = []
        try:
            n = int(self.var_qtd_casas.get())
        except Exception:
            n = 1
        for i in range(1, n + 1):
            tk.Label(self.frm_ruas, text=f"CASA {i}:",
                     bg=COR["bg"], fg=COR["subtexto"],
                     font=("Segoe UI", 8)).pack(anchor="w")
            var = tk.StringVar()
            tk.Entry(self.frm_ruas, textvariable=var, width=46,
                     bg=COR["campo"], fg=COR["texto"],
                     insertbackground=COR["texto"], relief="flat",
                     font=("Segoe UI", 9)).pack(fill="x", pady=(0, 3))
            self._campos_ruas.append(var)

    # ──────────────────────────────────────────
    # LOG
    # ──────────────────────────────────────────

    def _log(self, msg: str):
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")

    def _step(self, pct: int, desc: str):
        self.pb["value"] = pct
        self.lbl_prog.config(text=desc)

    # ──────────────────────────────────────────
    # VALIDAÇÃO
    # ──────────────────────────────────────────

    def _validar(self) -> bool:
        erros = []

        if not self.var_word.get() or not Path(self.var_word.get()).exists():
            erros.append("Template Word não encontrado.")
        if not self.var_excel.get() or not Path(self.var_excel.get()).exists():
            erros.append("Memorial Excel não encontrado.")
        if not self.var_saida.get() or not Path(self.var_saida.get()).exists():
            erros.append("Pasta de saída não encontrada.")
        if not self.var_eng.get():
            erros.append("Selecione um engenheiro.")

        obrigatorios = ["art", "logradouro", "quadra_lote", "bairro", "cidade"]
        nomes = {"art":"Número da ART","logradouro":"Logradouro",
                 "quadra_lote":"Quadra/Lote","bairro":"Bairro","cidade":"Cidade"}
        for k in obrigatorios:
            if not self.vars_art[k].get().strip():
                erros.append(f"Campo obrigatório: {nomes.get(k, k)}")

        if self.var_esquina.get() and not self.var_mesma_rua.get():
            for i, var in enumerate(self._campos_ruas):
                if not var.get().strip():
                    erros.append(f"Informe a rua da CASA {i+1}.")

        if erros:
            messagebox.showerror("Campos inválidos", "\n".join(erros))
            return False
        return True

    # ──────────────────────────────────────────
    # MONTAR LISTA DE CASAS
    # ──────────────────────────────────────────

    def _montar_casas(self) -> list:
        n = int(self.var_qtd_casas.get())
        logr_base = self.vars_art["logradouro"].get().strip()
        casas = []

        if self.var_esquina.get() and not self.var_mesma_rua.get():
            # Cada casa tem rua própria
            for i, var in enumerate(self._campos_ruas):
                casas.append({"num": i + 1, "logradouro": var.get().strip()})
        else:
            # Todas as casas usam o logradouro da ART
            for i in range(n):
                casas.append({"num": i + 1, "logradouro": logr_base})

        return casas

    # ──────────────────────────────────────────
    # INICIAR PROCESSAMENTO
    # ──────────────────────────────────────────

    def _iniciar(self):
        if not self._validar():
            return

        eng      = self.var_eng.get()
        eng_info = ENGENHEIROS[eng]

        # Montar caminho da assinatura
        assin_dir  = self.var_assindir.get().strip()
        assin_file = eng_info["assinatura"]
        assin_path = os.path.join(assin_dir, assin_file) if assin_dir else ""

        dados = {
            "art":            self.vars_art["art"].get().strip(),
            "crea":           self.vars_art["crea"].get().strip() or eng_info["crea"],
            "contratante":    self.vars_art["contratante"].get().strip(),
            "logradouro":     self.vars_art["logradouro"].get().strip(),
            "quadra_lote":    self.vars_art["quadra_lote"].get().strip(),
            "bairro":         self.vars_art["bairro"].get().strip(),
            "complemento":    self.vars_art["complemento"].get().strip(),
            "cep":            self.vars_art["cep"].get().strip(),
            "cidade":         self.vars_art["cidade"].get().strip(),
            "uf":             self.vars_art["uf"].get().strip() or "GO",
            "engenheiro_nome": eng,
            "cpf":            eng_info["cpf"],
            "assinatura_path": assin_path,
        }

        params = {
            "word_template":  self.var_word.get(),
            "excel_template": self.var_excel.get(),
            "saida_dir":      self.var_saida.get(),
            "assinatura_path": assin_path,
            "esgoto_sim":     self.var_esgoto.get(),
            "casas":          self._montar_casas(),
            "dados":          dados,
        }

        # Limpar log e barra
        self.txt_log.config(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.config(state="disabled")
        self.pb["value"] = 0

        self.btn.config(state="disabled")

        def _run():
            try:
                processar(
                    params,
                    step_cb=lambda p, d: self.after(0, self._step, p, d),
                    log=lambda m: self.after(0, self._log, m),
                )
                self.after(0, self._done, True)
            except Exception as e:
                self.after(0, self._log, f"\n✗ ERRO: {e}")
                self.after(0, self._done, False)

        threading.Thread(target=_run, daemon=True).start()

    def _done(self, ok: bool):
        self.btn.config(state="normal")
        if ok:
            messagebox.showinfo(
                "Concluído",
                f"Documentos gerados com sucesso!\n\n"
                f"Pasta: {self.var_saida.get()}"
            )
        else:
            messagebox.showerror(
                "Erro",
                "Ocorreu um erro durante o processamento.\n"
                "Verifique o log para detalhes."
            )


# ══════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════

if __name__ == "__main__":
    App().mainloop()
