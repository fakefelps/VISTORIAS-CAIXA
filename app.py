"""
BERÇAN PROJETOS — Preenchimento Automático de Documentos
Declaração ART (Word) + Memorial (Excel) → PDF
v3 — templates embutidos, dois Word (fossa/esgoto), logo, assets internos
"""
import multiprocessing
multiprocessing.freeze_support()

import os, sys, shutil, zipfile, threading, datetime, re, traceback
from pathlib import Path
from io import BytesIO

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ══════════════════════════════════════════════
# PATHS
# ══════════════════════════════════════════════

def resource_path(rel: str) -> str:
    base = getattr(sys, "_MEIPASS", Path(__file__).parent)
    return str(Path(base) / rel)

def asset(nome: str) -> str:
    return resource_path(os.path.join("assets", nome))

# ══════════════════════════════════════════════
# ENGENHEIROS
# ══════════════════════════════════════════════

ENGENHEIROS = {
    "FELIPE GUILHERME BERÇAN":          {"cpf":"147.849.107-86",  "crea":"1022722034D-GO", "assinatura":"FELIPE.png"},
    "CAIO ARAUJO BRAGA":                {"cpf":"011.309.411-67",  "crea":"CREA-GO",         "assinatura":"CAIO.png"},
    "JOÃO VITOR CABRAL DE MORAIS":      {"cpf":"038.144.411-25",  "crea":"CREA-GO",         "assinatura":"JOAO_VITOR.jpg"},
    "JULIO CESAR GOMES DE MORAIS FILHO":{"cpf":"033.865.821-17",  "crea":"CREA-GO",         "assinatura":"JULIO_CESAR.png"},
    "PAULA FLEURY DE MORAIS":           {"cpf":"033.813.881-18",  "crea":"CREA-GO",         "assinatura":"PAULA.png"},
    "ISAAC NATAN SANTOS":               {"cpf":"701.117.261-07",  "crea":"CREA-GO",         "assinatura":"ISAAC.png"},
}

# ══════════════════════════════════════════════
# CONSTANTES DOS TEMPLATES
# ══════════════════════════════════════════════

# Templates Word embutidos
TEMPLATE_FOSSA  = "TEMPLATE_FOSSA.docx"
TEMPLATE_ESGOTO = "TEMPLATE_ESGOTO.docx"

# Estrutura do template FOSSA
FOSSA_LINHA_ASS = 36   # parágrafo com ____
FOSSA_DATA_IDX  = 30
FOSSA_ENG_IDX   = 38
FOSSA_CREA_IDX  = 39

# Estrutura do template ESGOTO
ESGOTO_LINHA_ASS = 41
ESGOTO_DATA_IDX  = 35
ESGOTO_ENG_IDX   = 43
ESGOTO_CREA_IDX  = 44

# Shapes dos checkboxes de esgoto no Excel
SHAPE_ESGOTO_SIM = "QO012,12.L0C0;L0C-34^"
SHAPE_ESGOTO_NAO = "QO012,22.L0C0;L0C-37^"

# ══════════════════════════════════════════════
# PALETA
# ══════════════════════════════════════════════

COR = {
    "bg":      "#1a1f2e", "bg2":    "#232b3e",
    "bg_log":  "#0f1520", "campo":  "#2a3550",
    "botao":   "#2563eb", "barra":  "#22c55e",
    "texto":   "#f1f5f9", "subtexto":"#7c8fa8",
    "log":     "#4ade80", "divisor": "#2d3748",
    "roxo":    "#7c3aed",
}

# ══════════════════════════════════════════════
# UTILS
# ══════════════════════════════════════════════

def formatar_data_hoje() -> str:
    h = datetime.date.today()
    meses = ["janeiro","fevereiro","março","abril","maio","junho",
             "julho","agosto","setembro","outubro","novembro","dezembro"]
    return f"{h.day} de {meses[h.month-1]} de {h.year}"

def criar_pasta_saida() -> str:
    pasta = Path.home() / "Downloads" / "Bercan Projetos" / datetime.date.today().strftime("%Y-%m-%d")
    pasta.mkdir(parents=True, exist_ok=True)
    return str(pasta)

def normalizar_end(logr: str, ql: str) -> str:
    a = logr.strip()
    for p in ["AVENIDA ","AV. ","AV ","RUA ","R. "]:
        if a.upper().startswith(p): a = a[len(p):]; break
    ql2 = re.sub(r'\bQUADRA\b','QD',ql,flags=re.I)
    ql2 = re.sub(r'\bLOTE\b','LT',ql2,flags=re.I)
    return f"{a.strip()} {ql2.strip()}".strip()

def nome_arq(tipo, num, logr, ql):
    return f"{tipo} CS {num} - {normalizar_end(logr, ql)}"

# ══════════════════════════════════════════════
# WORD
# ══════════════════════════════════════════════

def _preto(para):
    from docx.shared import RGBColor
    for r in para.runs:
        try: r.font.color.rgb = RGBColor(0,0,0)
        except Exception: pass

def _sub(para, ph, val):
    from docx.shared import RGBColor
    txt = "".join(r.text for r in para.runs)
    if ph not in txt: return
    novo = txt.replace(ph, val)
    for r in para.runs: r.text = ""
    if para.runs:
        para.runs[0].text = novo
        try: para.runs[0].font.color.rgb = RGBColor(0,0,0)
        except Exception: pass

def preencher_word(esgoto_sim, saida_path, dados, num_casa, log=None):
    """
    Escolhe o template correto (fossa ou esgoto), preenche placeholders,
    força preto em todos os runs, insere assinatura flutuante.
    """
    from docx import Document
    from docx.shared import RGBColor

    def L(m):
        if log: log(m)

    # Escolher template
    tpl_nome = TEMPLATE_ESGOTO if esgoto_sim else TEMPLATE_FOSSA
    tpl_path = asset(tpl_nome)
    linha_ass = ESGOTO_LINHA_ASS if esgoto_sim else FOSSA_LINHA_ASS

    L(f"Template: {'ESGOTO' if esgoto_sim else 'FOSSA'}")
    doc = Document(tpl_path)

    subs = {
        "{1}":                      dados.get("art",""),
        "{2}":                      dados.get("crea",""),
        "{5}":                      dados.get("logradouro",""),
        "{6}":                      dados.get("quadra_lote",""),
        "{7}":                      dados.get("bairro",""),
        "{9}":                      f"CASA {num_casa}",
        "{10}":                     dados.get("cidade",""),
        "{11}":                     dados.get("uf",""),
        "{ENGENHEIRO SELECIONADO}": dados.get("engenheiro_nome",""),
        "{dia/mes/ano}":            formatar_data_hoje(),
    }

    L("Preenchendo e forçando preto...")
    for i, para in enumerate(doc.paragraphs):
        for ph, val in subs.items():
            _sub(para, ph, val)
        _preto(para)

    L("Inserindo assinatura...")
    _assinatura_word(doc, dados.get("assinatura_path",""), linha_ass, log)

    L(f"Salvando: {Path(saida_path).name}")
    doc.save(saida_path)
    L("Word salvo.")

def _assinatura_word(doc, img_path, linha_ass_idx, log=None):
    """
    Insere assinatura como imagem flutuante behind-text no parágrafo
    anterior à linha ___, sem deslocar nenhum conteúdo.
    """
    from docx.shared import Inches
    from docx.oxml.ns import qn
    import lxml.etree as etree, copy

    def L(m):
        if log: log(m)

    if not img_path or not os.path.exists(img_path):
        L(f"⚠ Assinatura não encontrada: {img_path}"); return

    # Parágrafo de inserção = um antes da linha ___
    idx = max(0, linha_ass_idx - 1)
    try:
        target = doc.paragraphs[idx]
        for r in target.runs: r.text = ""

        run = target.add_run()
        run.add_picture(img_path, width=Inches(1.8))

        drawing = run._r.find(qn('w:drawing'))
        if drawing is None: L("Assinatura inline."); return

        inline = drawing.find(qn('wp:inline'))
        if inline is None: L("Assinatura inline (sem conversão)."); return

        # Copiar elemento graphic
        graphic_el = None
        for child in inline:
            if 'graphic' in child.tag:
                graphic_el = copy.deepcopy(child); break

        W = int(Inches(1.8).emu)
        H = int(Inches(0.55).emu)
        anchor_xml = (
            f'<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
            f'distT="0" distB="0" distL="114300" distR="114300" '
            f'simplePos="0" relativeHeight="251658240" behindDoc="1" '
            f'locked="0" layoutInCell="1" allowOverlap="1">'
            f'<wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="column"><wp:posOffset>457200</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>-400000</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{W}" cy="{H}"/>'
            f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:wrapNone/>'
            f'<wp:docPr id="99" name="Assinatura"/>'
            f'<wp:cNvGraphicFramePr/>'
            f'</wp:anchor>'
        )
        anchor = etree.fromstring(anchor_xml)
        if graphic_el is not None:
            anchor.append(graphic_el)
        drawing.remove(inline)
        drawing.append(anchor)
        L("Assinatura inserida (flutuante behind text).")
    except Exception as e:
        L(f"⚠ Assinatura fallback: {e}")
        try:
            t = doc.paragraphs[idx]
            for r in t.runs: r.text = ""
            t.add_run().add_picture(img_path, width=Inches(1.5))
            L("Assinatura inline (fallback).")
        except Exception as e2:
            L(f"⚠ Erro assinatura: {e2}")

def exportar_word_pdf(docx_path, pdf_path, log=None):
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = None
    try:
        doc = word.Documents.Open(os.path.abspath(docx_path))
        doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
        doc.Close(); word.Quit()
        if log: log(f"✓ PDF Word: {Path(pdf_path).name}")
    except Exception:
        try:
            if doc: doc.Close()
            word.Quit()
        except Exception: pass
        raise

# ══════════════════════════════════════════════
# EXCEL
# ══════════════════════════════════════════════

def _excel_preencher(template_path, xlsx_saida, dados, num_casa, modo_mapeado, log):
    import win32com.client
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False; xl.DisplayAlerts = False
    wb = None
    try:
        wb = xl.Workbooks.Open(os.path.abspath(template_path))
        ws = wb.Worksheets("ElemConstrutivos")
        log("Excel aberto.")

        if modo_mapeado:
            ql = dados.get("quadra_lote","")
            mapa = {
                "G40":  dados.get("contratante",""),
                "G43":  dados.get("engenheiro_nome",""),
                "AH43": dados.get("crea",""),
                "AP43": "GO",
                "AR43": dados.get("cpf",""),
                "G47":  dados.get("logradouro",""),
                "AJ47": f"{ql}   CASA {num_casa}",
                "G49":  dados.get("bairro",""),
                "V49":  dados.get("cep",""),
                "AA49": dados.get("cidade",""),
                "AU49": dados.get("uf","GO"),
                "H53":  dados.get("engenheiro_nome",""),
                "Y54":  dados.get("art",""),
                "H75":  f"GOIÂNIA, {formatar_data_hoje()}",
                "AE77": dados.get("engenheiro_nome",""),
                "AE78": dados.get("cpf",""),
                "AE79": dados.get("crea",""),
            }
            log("Preenchendo células (mapeado)...")
            for coord, val in mapa.items():
                ws.Range(coord).Value = val
        else:
            # Modo auto: detecta células com fundo azul claro CAIXA
            ql = dados.get("quadra_lote","")
            campos = [
                dados.get("contratante",""),
                dados.get("engenheiro_nome",""),
                dados.get("crea",""),
                "GO",
                dados.get("cpf",""),
                dados.get("logradouro",""),
                f"{ql} CASA {num_casa}",
                dados.get("bairro",""),
                dados.get("cep",""),
                dados.get("cidade",""),
                dados.get("uf","GO"),
            ]
            idx = 0
            log("Preenchendo células (auto-detecção azul)...")
            for row in ws.UsedRange.Rows:
                for cell in row.Cells:
                    c = cell.Interior.Color
                    r,g,b = c&0xFF,(c>>8)&0xFF,(c>>16)&0xFF
                    if b>170 and g>190 and r>160 and b>=g and idx<len(campos):
                        if not str(cell.Value or "").strip():
                            cell.Value = campos[idx]; idx += 1
            log(f"Auto-preenchimento: {idx} células.")

        log(f"Salvando .xlsx: {Path(xlsx_saida).name}")
        wb.SaveAs(os.path.abspath(xlsx_saida), 51)
        wb.Close(False); xl.Quit()
        log("Excel salvo.")
    except Exception:
        try:
            if wb: wb.Close(False)
            xl.Quit()
        except Exception: pass
        raise

def _checkboxes_xml(xlsx_path, esgoto_sim, log=None):
    """
    Manipula APENAS os dois checkboxes de esgoto via XML usando lxml.
    lxml preserva todos os prefixos de namespace originais (xdr:, a:, r:)
    evitando que o Excel trate os outros shapes como inválidos.
    """
    from lxml import etree
    NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    NS_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
    tmp = xlsx_path + ".tmp"
    shutil.copy2(xlsx_path, tmp)
    try:
        with zipfile.ZipFile(tmp,"r") as zi, zipfile.ZipFile(xlsx_path,"w",zipfile.ZIP_DEFLATED) as zo:
            for item in zi.infolist():
                data = zi.read(item.filename)
                if item.filename == "xl/drawings/drawing1.xml":
                    root = etree.fromstring(data)
                    for anch in root.findall(f"{{{NS_XDR}}}twoCellAnchor"):
                        sp = anch.find(f"{{{NS_XDR}}}sp")
                        if sp is None: continue
                        cNvPr = sp.find(f".//{{{NS_XDR}}}cNvPr")
                        if cNvPr is None: continue
                        nome = cNvPr.get("name","")
                        if nome not in (SHAPE_ESGOTO_SIM, SHAPE_ESGOTO_NAO): continue
                        spPr = sp.find(f"{{{NS_XDR}}}spPr")
                        if spPr is None: continue
                        marcado = (nome==SHAPE_ESGOTO_SIM and esgoto_sim) or \
                                  (nome==SHAPE_ESGOTO_NAO and not esgoto_sim)
                        cor = "000000" if marcado else "FFFFFF"
                        # Remover fill anterior
                        for tag in [f"{{{NS_A}}}solidFill", f"{{{NS_A}}}noFill"]:
                            el = spPr.find(tag)
                            if el is not None: spPr.remove(el)
                        # Inserir solidFill com cor correta
                        solid = etree.SubElement(spPr, f"{{{NS_A}}}solidFill")
                        clr   = etree.SubElement(solid, f"{{{NS_A}}}srgbClr")
                        clr.set("val", cor)
                    # lxml preserva namespaces originais (xdr:, a:, r:)
                    data = etree.tostring(root, xml_declaration=True,
                                         encoding="UTF-8", standalone=True)
                zo.writestr(item, data)
        if log: log(f"Checkbox esgoto → {'SIM' if esgoto_sim else 'NÃO'}")
    finally:
        if os.path.exists(tmp): os.remove(tmp)

def _forcar_preto_excel(xlsx_path, log=None):
    from openpyxl import load_workbook
    from openpyxl.styles import Font
    try:
        wb = load_workbook(xlsx_path)
        ws = wb["ElemConstrutivos"]
        for coord in ["G40","G43","AH43","AP43","AR43","G47","AJ47",
                      "G49","V49","AA49","AU49","H53","Y54","H75","AE77","AE78","AE79"]:
            cell = ws[coord]
            if cell.value:
                f = cell.font
                cell.font = Font(name=f.name, size=f.size, bold=f.bold,
                                 italic=f.italic, underline=f.underline, color="000000")
        wb.save(xlsx_path)
        if log: log("Cores → preto no Excel.")
    except Exception as e:
        if log: log(f"⚠ Força-preto Excel: {e}")

def _assinatura_excel(xlsx_path, img_path, log=None):
    from openpyxl import load_workbook
    from openpyxl.drawing.image import Image as XLImage
    if not img_path or not os.path.exists(img_path):
        if log: log("⚠ Assinatura Excel não encontrada."); return
    try:
        wb = load_workbook(xlsx_path)
        ws = wb["ElemConstrutivos"]
        img = XLImage(img_path)
        img.width, img.height, img.anchor = 120, 45, "AE73"
        ws.add_image(img)
        wb.save(xlsx_path)
        if log: log("Assinatura Excel inserida.")
    except Exception as e:
        if log: log(f"⚠ Assinatura Excel: {e}")

def _pdf_excel(xlsx_path, pdf_path, log=None):
    import win32com.client
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False; xl.DisplayAlerts = False
    wb = None
    try:
        wb = xl.Workbooks.Open(os.path.abspath(xlsx_path))
        ws = wb.Worksheets("ElemConstrutivos")
        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.FitToPagesTall = 1
        ws.ExportAsFixedFormat(0, os.path.abspath(pdf_path), 1, True, False)
        wb.Close(False); xl.Quit()
        if log: log(f"✓ PDF Excel: {Path(pdf_path).name}")
    except Exception:
        try:
            if wb: wb.Close(False)
            xl.Quit()
        except Exception: pass
        raise

def preencher_excel_e_pdf(template_path, xlsx_saida, pdf_saida,
                          dados, esgoto_sim, num_casa, assinatura_path,
                          modo_mapeado=True, log=None):
    def L(m):
        if log: log(m)
    L(f"Excel {'mapeado' if modo_mapeado else 'auto'}...")
    try:
        _excel_preencher(template_path, xlsx_saida, dados, num_casa, modo_mapeado, L)
    except Exception:
        L("✗ Erro Excel:\n" + traceback.format_exc()); raise
    try: _checkboxes_xml(xlsx_saida, esgoto_sim, L)
    except Exception as e: L(f"⚠ Checkboxes: {e}")
    _forcar_preto_excel(xlsx_saida, L)
    if assinatura_path: _assinatura_excel(xlsx_saida, assinatura_path, L)
    try: _pdf_excel(xlsx_saida, pdf_saida, L)
    except Exception:
        L("✗ Erro PDF Excel:\n" + traceback.format_exc()); raise

# ══════════════════════════════════════════════
# OCR
# ══════════════════════════════════════════════

def ler_art_ocr(pdf_path, log=None):
    def L(m):
        if log: log(m)
    res = {"art":"","crea":"","contratante":"","logradouro":"",
           "quadra_lote":"","bairro":"","cep":"","cidade":"","uf":"GO"}
    try:
        import pytesseract
        from PIL import Image
        try:
            import fitz
            L("Renderizando PDF (PyMuPDF)...")
            d = fitz.open(pdf_path)
            pix = d[0].get_pixmap(matrix=fitz.Matrix(2.5,2.5))
            img = Image.open(BytesIO(pix.tobytes("png"))); d.close()
        except ImportError:
            from pdf2image import convert_from_path
            L("Renderizando PDF (pdf2image)...")
            img = convert_from_path(pdf_path, dpi=250)[0]

        L("OCR em andamento...")
        texto = pytesseract.image_to_string(img, lang="por")
        linhas = [l.strip() for l in texto.split("\n") if l.strip()]

        for l in linhas:
            m = re.search(r'\b(\d{13})\b', l)
            if m: res["art"] = m.group(1); break
        for l in linhas:
            m = re.search(r'(\d{7,10}[A-Z]-GO)', l)
            if m: res["crea"] = m.group(1); break
        for l in linhas:
            m = re.search(r'\b(\d{5}-\d{3})\b', l)
            if m: res["cep"] = m.group(1); break
        for l in linhas:
            m = re.search(r'Quadra[:\s]*(\d+)\s*Lote[:\s]*(\d+)', l, re.I)
            if m:
                res["quadra_lote"] = f"QD {m.group(1)} LT {m.group(2)}"
                logr = l.split(m.group(0))[0].strip().rstrip(',')
                if logr: res["logradouro"] = logr
                break
        for l in linhas:
            m = re.search(r'Cidade[:\s]+([A-ZÀ-Ú][^\n]+?)(?:\s*-\s*(GO|SP|MG))?$', l, re.I)
            if m:
                res["cidade"] = m.group(1).strip()
                if m.group(2): res["uf"] = m.group(2)
                break
        for l in linhas:
            m = re.search(r'Bairro[:\s]+(.+)', l, re.I)
            if m: res["bairro"] = m.group(1).strip()[:60]; break

        n = sum(1 for v in res.values() if v)
        L(f"OCR: {n}/{len(res)} campos identificados.")
    except ImportError:
        L("⚠ pytesseract não disponível.")
    except Exception as e:
        L(f"⚠ Erro OCR: {e}")
    return res

# ══════════════════════════════════════════════
# ORQUESTRADOR
# ══════════════════════════════════════════════

def processar(params, step_cb=None, log=None):
    def S(p,d):
        if step_cb: step_cb(p,d)
    def L(m):
        if log: log(m)

    excel_tpl = params["excel_template"]
    saida     = params["saida_dir"]
    assin     = params["assinatura_path"]
    esgoto    = params["esgoto_sim"]
    casas     = params["casas"]
    dados     = params["dados"]
    mapeado   = params.get("modo_mapeado", True)

    total = len(casas) * 3; atual = 0

    for casa in casas:
        num  = casa["num"]
        logr = casa["logradouro"]
        ql   = dados["quadra_lote"]
        d    = {**dados, "logradouro": logr}

        L(f"\n{'='*42}\nCASA {num} — {logr}\n{'='*42}")

        nd   = nome_arq("DECLARAÇÃO ART", num, logr, ql)
        nm   = nome_arq("MEMORIAL",       num, logr, ql)
        docx = os.path.join(saida, nd+".docx")
        xlsx = os.path.join(saida, nm+".xlsx")
        pdfd = os.path.join(saida, nd+".pdf")
        pdfm = os.path.join(saida, nm+".pdf")

        atual+=1; S(int(atual/total*100), f"Casa {num}: Word...")
        preencher_word(esgoto, docx, d, num, log=L)

        atual+=1; S(int(atual/total*100), f"Casa {num}: Excel...")
        preencher_excel_e_pdf(excel_tpl, xlsx, pdfm, d, esgoto, num, assin, mapeado, L)

        atual+=1; S(int(atual/total*100), f"Casa {num}: PDF Word...")
        exportar_word_pdf(docx, pdfd, L)

        L(f"✓ Casa {num} concluída.")

    S(100, "Concluído!")
    L("\n✓ Todos os documentos gerados.")

# ══════════════════════════════════════════════
# INTERFACE
# ══════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("BERÇAN PROJETOS — Preenchimento de Documentos")
        self.configure(bg=COR["bg"])
        self.resizable(True, True)
        self.minsize(620, 520)
        self._campos_ruas = []
        self._build_ui()
        self._centralizar()

    def _build_ui(self):
        P = 14

        # Cabeçalho
        hdr = tk.Frame(self, bg="#111827", pady=10)
        hdr.pack(fill="x", side="top")
        try:
            from PIL import Image, ImageTk
            for ext in ["LOGO.jpg","LOGO.png"]:
                lp = asset(ext)
                if os.path.exists(lp):
                    img = Image.open(lp).resize((48,48), Image.LANCZOS)
                    self._logo = ImageTk.PhotoImage(img)
                    tk.Label(hdr, image=self._logo, bg="#111827").pack(side="left", padx=14)
                    break
        except Exception: pass
        tf = tk.Frame(hdr, bg="#111827"); tf.pack(side="left")
        tk.Label(tf, text="BERÇAN PROJETOS", font=("Segoe UI",15,"bold"),
                 bg="#111827", fg=COR["texto"]).pack(anchor="w")
        tk.Label(tf, text="Preenchimento Automático de Documentos",
                 font=("Segoe UI",9), bg="#111827", fg=COR["subtexto"]).pack(anchor="w")

        # Scroll
        vsb = tk.Scrollbar(self, orient="vertical")
        vsb.pack(side="right", fill="y")
        self._cv = tk.Canvas(self, bg=COR["bg"], highlightthickness=0, yscrollcommand=vsb.set)
        self._cv.pack(side="left", fill="both", expand=True)
        vsb.config(command=self._cv.yview)
        inner = tk.Frame(self._cv, bg=COR["bg"])
        self._iid = self._cv.create_window((0,0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: self._cv.configure(scrollregion=self._cv.bbox("all")))
        self._cv.bind("<Configure>", lambda e: self._cv.itemconfig(self._iid, width=e.width))
        self._cv.bind_all("<MouseWheel>", lambda e: self._cv.yview_scroll(int(-1*(e.delta/120)),"units"))

        row = tk.Frame(inner, bg=COR["bg"])
        row.pack(fill="both", expand=True, padx=P, pady=P)
        ce = tk.Frame(row, bg=COR["bg"])
        cd = tk.Frame(row, bg=COR["bg"])
        ce.pack(side="left", fill="both", expand=True, padx=(0,10))
        cd.pack(side="left", fill="both", expand=True)

        # ── COLUNA ESQ ──
        self._s(ce, "MEMORIAL EXCEL")
        self.var_excel = self._arq(ce, "Arquivo Memorial (.xlsx):", "excel")

        self._s(ce, "MODO DO MEMORIAL")
        self.var_modo = tk.StringVar(value="mapeado")
        mf = tk.Frame(ce, bg=COR["bg"]); mf.pack(fill="x", pady=(0,6))
        for val, txt in [("mapeado","Mapeado (com {N})"),("nao_mapeado","Não mapeado (detectar azul)")]:
            tk.Radiobutton(mf, text=txt, variable=self.var_modo, value=val,
                           bg=COR["bg"], fg=COR["texto"], selectcolor=COR["campo"],
                           font=("Segoe UI",9)).pack(side="left", padx=(0,12))

        self._s(ce, "ENGENHEIRO RESPONSÁVEL")
        self.var_eng = tk.StringVar()
        cb = ttk.Combobox(ce, textvariable=self.var_eng,
                          values=list(ENGENHEIROS.keys()), state="readonly", width=42)
        cb.pack(fill="x", pady=(0,4))
        cb.bind("<<ComboboxSelected>>", self._on_eng)
        self.lbl_cpf  = self._lbl(ce, "CPF: —")
        self.lbl_crea = self._lbl(ce, "CREA: —")

        self._s(ce, "DADOS DA ART")
        of = tk.Frame(ce, bg=COR["bg"]); of.pack(fill="x", pady=(0,6))
        self.var_art_pdf = tk.StringVar()
        tk.Entry(of, textvariable=self.var_art_pdf, width=26,
                 bg=COR["campo"], fg=COR["texto"], insertbackground=COR["texto"],
                 relief="flat", font=("Segoe UI",8)).pack(side="left", fill="x", expand=True)
        tk.Button(of, text="...", bg=COR["campo"], fg=COR["texto"], relief="flat",
                  font=("Segoe UI",8), command=self._browse_art).pack(side="left", padx=(3,0))
        tk.Button(of, text="🔍 LER ART", bg=COR["roxo"], fg=COR["texto"],
                  relief="flat", font=("Segoe UI",8,"bold"), cursor="hand2",
                  command=self._ocr).pack(side="left", padx=(5,0))

        self.vars_art = {}
        for lbl, key in [
            ("Número da ART {1}:","art"),
            ("Registro CREA {2}:","crea"),
            ("Contratante {4}:","contratante"),
            ("Logradouro {5}:","logradouro"),
            ("Quadra e Lote {6}:","quadra_lote"),
            ("Bairro {7}:","bairro"),
            ("CEP {8}:","cep"),
            ("Cidade {10}:","cidade"),
            ("UF {11}:","uf"),
        ]:
            self.vars_art[key] = self._txt(ce, lbl)

        # ── COLUNA DIR ──
        self._s(cd, "OPÇÕES")
        self.var_esgoto = tk.BooleanVar(value=False)
        self._ck(cd, "Sistema público de esgoto (SIM)", self.var_esgoto)

        # Info sobre template
        tk.Label(cd, text="  ↳ Template Word selecionado automaticamente",
                 bg=COR["bg"], fg=COR["subtexto"], font=("Segoe UI",8)).pack(anchor="w")

        self.var_esquina = tk.BooleanVar(value=False)
        self._ck(cd, "Lote de esquina", self.var_esquina, command=self._on_esquina)

        qf = tk.Frame(cd, bg=COR["bg"]); qf.pack(fill="x", pady=(6,0))
        tk.Label(qf, text="Quantidade de casas:", bg=COR["bg"], fg=COR["subtexto"],
                 font=("Segoe UI",9)).pack(side="left")
        self.var_qtd = tk.IntVar(value=1)
        spn = tk.Spinbox(qf, from_=1, to=20, width=5, textvariable=self.var_qtd,
                         bg=COR["campo"], fg=COR["texto"], insertbackground=COR["texto"],
                         command=self._on_qtd)
        spn.pack(side="left", padx=6)
        spn.bind("<FocusOut>", lambda e: self._on_qtd())

        self.var_mesma = tk.BooleanVar(value=True)
        self.frm_esq_opt = tk.Frame(cd, bg=COR["bg"])
        for v,t in [(True,"Mesma rua"),(False,"Ruas diferentes")]:
            tk.Radiobutton(self.frm_esq_opt, text=t, variable=self.var_mesma, value=v,
                           bg=COR["bg"], fg=COR["texto"], selectcolor=COR["campo"],
                           command=self._on_mesma).pack(anchor="w")
        self.frm_esq_opt.pack_forget()

        self._s(cd, "RUAS POR CASA")
        self.frm_ruas = tk.Frame(cd, bg=COR["bg"]); self.frm_ruas.pack(fill="x")
        tk.Label(self.frm_ruas, text="(Ativo: lote de esquina + ruas diferentes)",
                 bg=COR["bg"], fg=COR["subtexto"], font=("Segoe UI",8)).pack()

        self._s(cd, "LOG")
        lf = tk.Frame(cd, bg=COR["bg_log"]); lf.pack(fill="x")
        self.txt_log = tk.Text(lf, height=10, width=44, bg=COR["bg_log"], fg=COR["log"],
                               font=("Consolas",8), relief="flat", state="disabled")
        sb = tk.Scrollbar(lf, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=sb.set)
        self.txt_log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self._s(cd, "PROGRESSO")
        self.lbl_prog = tk.Label(cd, text="Aguardando...", bg=COR["bg"], fg=COR["subtexto"],
                                 font=("Segoe UI",8))
        self.lbl_prog.pack(anchor="w")
        self.pb = ttk.Progressbar(cd, length=360, mode="determinate")
        self.pb.pack(fill="x", pady=(2,8))
        sty = ttk.Style(); sty.theme_use("default")
        sty.configure("bp.Horizontal.TProgressbar", troughcolor=COR["campo"], background=COR["barra"])
        self.pb.configure(style="bp.Horizontal.TProgressbar")

        self.btn = tk.Button(cd, text="⚡  GERAR DOCUMENTOS",
                             font=("Segoe UI",11,"bold"), bg=COR["botao"], fg=COR["texto"],
                             activebackground="#1d4ed8", activeforeground=COR["texto"],
                             relief="flat", pady=10, cursor="hand2", command=self._iniciar)
        self.btn.pack(fill="x", pady=(4,0))

    # helpers
    def _s(self, p, t):
        tk.Label(p, text=t, bg=COR["bg"], fg=COR["subtexto"],
                 font=("Segoe UI",8,"bold")).pack(anchor="w", pady=(10,2))
        tk.Frame(p, bg=COR["divisor"], height=1).pack(fill="x", pady=(0,4))

    def _lbl(self, p, t):
        l = tk.Label(p, text=t, bg=COR["bg"], fg=COR["subtexto"], font=("Segoe UI",8))
        l.pack(anchor="w"); return l

    def _ck(self, p, t, v, command=None):
        tk.Checkbutton(p, text=t, variable=v, bg=COR["bg"], fg=COR["texto"],
                       selectcolor=COR["campo"], activebackground=COR["bg"],
                       font=("Segoe UI",9), command=command).pack(anchor="w", pady=2)

    def _txt(self, p, lbl, default=""):
        tk.Label(p, text=lbl, bg=COR["bg"], fg=COR["subtexto"],
                 font=("Segoe UI",8)).pack(anchor="w")
        v = tk.StringVar(value=default)
        tk.Entry(p, textvariable=v, width=44, bg=COR["campo"], fg=COR["texto"],
                 insertbackground=COR["texto"], relief="flat",
                 font=("Segoe UI",9)).pack(fill="x", pady=(0,4))
        return v

    def _arq(self, p, lbl, tipo):
        tk.Label(p, text=lbl, bg=COR["bg"], fg=COR["subtexto"],
                 font=("Segoe UI",8)).pack(anchor="w")
        f = tk.Frame(p, bg=COR["bg"]); f.pack(fill="x", pady=(0,4))
        v = tk.StringVar()
        tk.Entry(f, textvariable=v, width=34, bg=COR["campo"], fg=COR["texto"],
                 insertbackground=COR["texto"], relief="flat",
                 font=("Segoe UI",9)).pack(side="left", fill="x", expand=True)
        def browse():
            if tipo=="excel":
                p2 = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls"),("Todos","*.*")])
            else:
                p2 = filedialog.askopenfilename(filetypes=[("Todos","*.*")])
            if p2: v.set(p2)
        tk.Button(f, text="...", bg=COR["campo"], fg=COR["texto"], relief="flat",
                  font=("Segoe UI",9), command=browse).pack(side="left", padx=(4,0))
        return v

    def _centralizar(self):
        self.update_idletasks()
        sw,sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w,h = min(int(sw*.90),1150), min(int(sh*.88),780)
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _on_eng(self, e=None):
        eng = self.var_eng.get()
        if eng in ENGENHEIROS:
            info = ENGENHEIROS[eng]
            self.lbl_cpf.config(text=f"CPF: {info['cpf']}")
            self.lbl_crea.config(text=f"CREA: {info['crea']}")
            if "crea" in self.vars_art and not self.vars_art["crea"].get():
                self.vars_art["crea"].set(info["crea"])

    def _on_esquina(self):
        if self.var_esquina.get(): self.frm_esq_opt.pack(fill="x", pady=(4,0))
        else: self.frm_esq_opt.pack_forget()
        self._on_mesma()

    def _on_mesma(self):
        if self.var_esquina.get() and not self.var_mesma.get(): self._rebuild_ruas()
        else: self._limpar_ruas()

    def _on_qtd(self):
        if self.var_esquina.get() and not self.var_mesma.get(): self._rebuild_ruas()

    def _limpar_ruas(self):
        for w in self.frm_ruas.winfo_children(): w.destroy()
        tk.Label(self.frm_ruas, text="(Ativo: lote de esquina + ruas diferentes)",
                 bg=COR["bg"], fg=COR["subtexto"], font=("Segoe UI",8)).pack()
        self._campos_ruas = []

    def _rebuild_ruas(self):
        for w in self.frm_ruas.winfo_children(): w.destroy()
        self._campos_ruas = []
        try: n = int(self.var_qtd.get())
        except Exception: n = 1
        for i in range(1, n+1):
            tk.Label(self.frm_ruas, text=f"CASA {i}:", bg=COR["bg"], fg=COR["subtexto"],
                     font=("Segoe UI",8)).pack(anchor="w")
            v = tk.StringVar()
            tk.Entry(self.frm_ruas, textvariable=v, width=44, bg=COR["campo"], fg=COR["texto"],
                     insertbackground=COR["texto"], relief="flat",
                     font=("Segoe UI",9)).pack(fill="x", pady=(0,3))
            self._campos_ruas.append(v)

    def _log(self, msg):
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", msg+"\n")
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")

    def _step(self, pct, desc):
        self.pb["value"] = pct
        self.lbl_prog.config(text=desc)

    def _browse_art(self):
        p = filedialog.askopenfilename(filetypes=[("PDF","*.pdf"),("Todos","*.*")])
        if p: self.var_art_pdf.set(p)

    def _ocr(self):
        pdf = self.var_art_pdf.get().strip()
        if not pdf or not Path(pdf).exists():
            messagebox.showwarning("OCR","Selecione o PDF da ART primeiro."); return
        self._log("🔍 Iniciando OCR..."); self.btn.config(state="disabled")
        def run():
            campos = ler_art_ocr(pdf, log=lambda m: self.after(0, self._log, m))
            self.after(0, self._aplicar_ocr, campos)
        threading.Thread(target=run, daemon=True).start()

    def _aplicar_ocr(self, campos):
        self.btn.config(state="normal")
        n = sum(1 for k,v in campos.items() if v and k in self.vars_art and not self.vars_art[k].get())
        for k, v in campos.items():
            if v and k in self.vars_art:
                self.vars_art[k].set(v)
        self._log(f"✓ OCR: campos preenchidos. Verifique antes de gerar.")

    def _validar(self):
        erros = []
        if not self.var_excel.get() or not Path(self.var_excel.get()).exists():
            erros.append("Memorial Excel não encontrado.")
        if not self.var_eng.get():
            erros.append("Selecione um engenheiro.")
        for k, nome in [("art","ART"),("logradouro","Logradouro"),
                        ("quadra_lote","Quadra/Lote"),("bairro","Bairro"),("cidade","Cidade")]:
            if not self.vars_art[k].get().strip():
                erros.append(f"Campo obrigatório: {nome}")
        if self.var_esquina.get() and not self.var_mesma.get():
            for i, v in enumerate(self._campos_ruas):
                if not v.get().strip(): erros.append(f"Informe rua da CASA {i+1}.")
        if erros:
            messagebox.showerror("Campos inválidos", "\n".join(erros)); return False
        return True

    def _montar_casas(self):
        n = int(self.var_qtd.get())
        lb = self.vars_art["logradouro"].get().strip()
        if self.var_esquina.get() and not self.var_mesma.get():
            return [{"num":i+1,"logradouro":v.get().strip()} for i,v in enumerate(self._campos_ruas)]
        return [{"num":i+1,"logradouro":lb} for i in range(n)]

    def _iniciar(self):
        if not self._validar(): return
        eng = self.var_eng.get()
        info = ENGENHEIROS[eng]
        assin = asset(info["assinatura"])
        if not os.path.exists(assin): assin = ""
        saida = criar_pasta_saida()
        dados = {
            "art":             self.vars_art["art"].get().strip(),
            "crea":            self.vars_art["crea"].get().strip() or info["crea"],
            "contratante":     self.vars_art["contratante"].get().strip(),
            "logradouro":      self.vars_art["logradouro"].get().strip(),
            "quadra_lote":     self.vars_art["quadra_lote"].get().strip(),
            "bairro":          self.vars_art["bairro"].get().strip(),
            "cep":             self.vars_art["cep"].get().strip(),
            "cidade":          self.vars_art["cidade"].get().strip(),
            "uf":              self.vars_art["uf"].get().strip() or "GO",
            "engenheiro_nome": eng,
            "cpf":             info["cpf"],
            "assinatura_path": assin,
        }
        params = {
            "excel_template": self.var_excel.get(),
            "saida_dir":      saida,
            "assinatura_path":assin,
            "esgoto_sim":     self.var_esgoto.get(),
            "casas":          self._montar_casas(),
            "dados":          dados,
            "modo_mapeado":   self.var_modo.get() == "mapeado",
        }
        self.txt_log.config(state="normal"); self.txt_log.delete("1.0","end")
        self.txt_log.config(state="disabled")
        self.pb["value"] = 0; self.btn.config(state="disabled")
        self._log(f"📁 Saída: {saida}")
        self._log(f"📄 Template: {'ESGOTO' if self.var_esgoto.get() else 'FOSSA'} (embutido)")

        def run():
            try:
                processar(params,
                          step_cb=lambda p,d: self.after(0,self._step,p,d),
                          log=lambda m: self.after(0,self._log,m))
                self.after(0, self._done, True, saida)
            except Exception:
                tb = traceback.format_exc()
                self.after(0, self._log, "\n✗ ERRO:\n"+tb)
                self.after(0, self._done, False, saida)
        threading.Thread(target=run, daemon=True).start()

    def _done(self, ok, saida):
        self.btn.config(state="normal")
        if ok:
            messagebox.showinfo("Concluído", f"Documentos gerados!\n\nPasta:\n{saida}")
            try: os.startfile(saida)
            except Exception: pass
        else:
            log_text = self.txt_log.get("1.0","end")
            ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            rel = os.path.join(saida, f"ERRO_{ts}.txt")
            try:
                with open(rel,"w",encoding="utf-8") as f:
                    f.write("RELATÓRIO DE ERRO — BERÇAN PROJETOS\n"+"="*50+"\n")
                    f.write(f"Data: {datetime.datetime.now()}\n")
                    f.write(f"Excel: {self.var_excel.get()}\n")
                    f.write(f"Eng: {self.var_eng.get()}\n"+"="*50+"\n\n"+log_text)
                messagebox.showerror("Erro", f"Erro no processamento.\nRelatório:\n{rel}")
            except Exception:
                messagebox.showerror("Erro","Erro. Verifique o log.")

if __name__ == "__main__":
    App().mainloop()
