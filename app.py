# -*- coding: utf-8 -*-
"""
BERÇAN PROJETOS — Preenchimento Automático de Documentos CAIXA
Versão 4.5 — Abril/2026

Correções v4 (original):
- Assinatura Memorial (Excel) via win32com com qualidade preservada (+50% tamanho)
- Assinatura Declaração (Word) +50% tamanho
- Checkboxes do Memorial via imagens PNG sobrepostas (confiável e editável)
- OCR da ART com detecção expandida + pré-processamento da imagem + regex tolerantes
- Bug do Modo Não Mapeado (float & int) corrigido
- Pasta destino renomeada: DOCUMENTOS DE VISTORIA
- Botão INTERROMPER processamento (threading.Event)

Correções v4.2:
- Assinatura Word: âncora dinâmica no parágrafo ____ (detectado automaticamente)
- Assinatura Word: posOffset -457200 EMU (0,5cm acima da linha de assinatura)
- Excel COM: xl.Visible/DisplayAlerts envolvidos em try/except (fix AttributeError)
- UI: botão LER ART e campo PDF removidos (OCR descartado)
- UF: campo mantido com padrão GO

Correções v4.1:
- CHECKBOX_ANCORA_CELULA corrigida de AR55 → AM70 (posição real no drawing XML)
- CHECKBOX_LARGURA_PT/ALTURA_PT ajustados para cobrir AM70:AP70 corretamente
- asset_checkbox() adicionada: fallback automático de extensão (.png/.jpeg/.png.jpeg)
- ASSINATURA_EXCEL_ANCORA ajustada de AE73 → AE74 (sem offset negativo frágil)
- Assinatura Word: posOffset corrigido para 0/−685800 (alinha à esquerda da coluna)
- SHAPE_ESGOTO_SIM/NAO confirmados via inspeção do drawing1.xml do template real
- UF padrão mantido como GO (OCR descartado — preenchimento manual preferido)
"""

# ============================================================
# PROTEÇÃO ANTI-LOOP PyInstaller — DEVE SER A PRIMEIRA COISA
# ============================================================
import multiprocessing
multiprocessing.freeze_support()

# ============================================================
# IMPORTS
# ============================================================
import os
import sys
import shutil
import zipfile
import threading
import datetime
import re
import traceback
from pathlib import Path
from io import BytesIO
from copy import deepcopy

import tkinter as tk

# ── SCPO: imports Selenium (usados apenas ao clicar "Preencher SCPO") ────────
import winreg as _winreg
import json as _json_scpo
import urllib.request as _urllib_scpo
import zipfile as _zipfile_scpo
from selenium import webdriver as _webdriver_scpo
from selenium.webdriver.common.by import By as _By_scpo
from selenium.webdriver.support.ui import WebDriverWait as _Wait_scpo
from selenium.webdriver.support.ui import Select as _Select_scpo
from selenium.webdriver.support import expected_conditions as _EC_scpo
from selenium.webdriver.chrome.service import Service as _Service_scpo

def buscar_cep(cep, callback_ok, callback_erro):
    """
    Consulta ViaCEP em thread separada e chama callback com o resultado.
    callback_ok(data): dict com logradouro, bairro, localidade, uf
    callback_erro(msg): string com mensagem de erro
    """
    import urllib.request, json, threading, ssl
    _ssl_ctx = ssl.create_default_context()
    _ssl_ctx.check_hostname = False
    _ssl_ctx.verify_mode = ssl.CERT_NONE
    cep_limpo = re.sub(r"[^0-9]", "", cep)
    if len(cep_limpo) != 8:
        callback_erro("CEP deve ter 8 dígitos")
        return
    def _worker():
        try:
            url = f"https://viacep.com.br/ws/{cep_limpo}/json/"
            with urllib.request.urlopen(url, timeout=5, context=_ssl_ctx) as r:
                data = json.loads(r.read().decode())
            if data.get("erro"):
                callback_erro("CEP não encontrado")
            else:
                callback_ok(data)
        except Exception as e:
            callback_erro(f"Erro na consulta: {e}")
    threading.Thread(target=_worker, daemon=True).start()

from tkinter import ttk, filedialog, messagebox

# Word
from docx import Document
from docx.shared import Inches, RGBColor, Pt
from docx.oxml.ns import qn

# Excel + imagens
from PIL import Image, ImageFilter, ImageOps

# COM (Word/Excel nativo)
import win32com.client
import pythoncom

# XML (preservação de namespaces)
from lxml import etree


# ============================================================
# CONSTANTES GLOBAIS
# ============================================================

# ----- Pasta de destino (ALTERADO v4) -----
PASTA_DESTINO = "DOCUMENTOS DE VISTORIA"

# ----- Templates Word -----
TEMPLATE_FOSSA = "TEMPLETE PARA FOSSA.docx"
TEMPLATE_ESGOTO = "TEMPLETE PARA ESGOTO.docx"
FOSSA_LINHA_ASS = 36
ESGOTO_LINHA_ASS = 41

# ----- Checkboxes como imagens PNG (NOVO v4) -----
# NOTA: os arquivos no repositório têm extensão dupla (.png.jpeg).
# A função asset_checkbox() abaixo resolve automaticamente a extensão correta.
CHECKBOX_COM_ESGOTO_IMG = "CHECKBOX_COM_ESGOTO.png"
CHECKBOX_SEM_ESGOTO_IMG = "CHECKBOX_SEM_ESGOTO.png"

# ----- Checkboxes imagem — valores calibrados 17/04/2026 -----
# 4 estados independentes: ancora + offset + tamanho
# Estado 1: Esgoto SIM
CHK1_ANCORA  = "AM70"
CHK1_OFF_X   = 10
CHK1_OFF_Y   = 3
CHK1_LARGURA = 4
CHK1_ALTURA  = 5
# Estado 2: Esgoto NÃO
CHK2_ANCORA  = "AP70"
CHK2_OFF_X   = 11
CHK2_OFF_Y   = 3
CHK2_LARGURA = 4
CHK2_ALTURA  = 5
# Estado 3: Condomínio SIM
CHK3_ANCORA  = "AM65"
CHK3_OFF_X   = 10
CHK3_OFF_Y   = 8
CHK3_LARGURA = 4
CHK3_ALTURA  = 5
# Estado 4: Condomínio Não se aplica
CHK4_ANCORA  = "AS65"
CHK4_OFF_X   = 12
CHK4_OFF_Y   = 8
CHK4_LARGURA = 4
CHK4_ALTURA  = 5
# Fallback de ancora para detecção automática
CHECKBOX_ANCORA_FALLBACK = "AM70"

# ----- Texto usado para localizar a linha do esgoto automaticamente -----
# A detecção varre a planilha procurando este fragmento (case-insensitive).
# Se o texto mudar na nova versão do memorial, atualizar apenas aqui.
TEXTO_ITEM_ESGOTO = "sistema público de coleta de esgoto sanitário"

# Checkboxes de casas geminadas — opções: "sim", "nao", "nao_se_aplica"
# Shapes confirmados via drawing1.xml:
#   Loteamentos  linha 64: SIM=QOCI,13... | NAO=QOCI,23... | NSA=QOCI,33...
#   Condomínios  linha 65: SIM=QOCN,13... | NAO=QOCN,23... | NSA=QOCN,33...
GEMINADAS_LOTEAMENTOS = "nao_se_aplica"
GEMINADAS_CONDOMINIOS = "nao_se_aplica"


# ----- Assinatura Word (DECLARAÇÃO) — +50% v4 -----
# Antes: Inches(1.8). Agora: Inches(2.7) — aumento proporcional de 50%.
ASSINATURA_WORD_LARGURA = Inches(2.7)

# ----- Assinatura Excel (MEMORIAL) — calibrado 17/04/2026 -----
ASSINATURA_EXCEL_ANCORA      = "AE72"
ASSINATURA_EXCEL_OFFSET_X_PT = 10
ASSINATURA_EXCEL_OFFSET_Y_PT = -5
ASSINATURA_EXCEL_LARGURA_PT  = 170
ASSINATURA_EXCEL_ALTURA_PT   = 55

# ----- Engenheiros cadastrados -----
ENGENHEIROS = {
    "FELIPE GUILHERME BERÇAN": {
        "cpf": "147.849.107-86",
        "crea": "1022722034D-GO",
        "assinatura": "FELIPE.png",
    },
    "CAIO ARAUJO BRAGA": {
        "cpf": "011.309.411-67",
        "crea": "CREA-GO",
        "assinatura": "CAIO.png",
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

# ----- Paleta de cores da UI -----
COR_FUNDO = "#1e2a3a"
COR_LOG_FUNDO = "#131c26"
COR_CAMPO = "#2a3f55"
COR_BOTAO = "#2e86de"
COR_BOTAO_STOP = "#e74c3c"
COR_PROGRESSO = "#4cd964"
COR_TEXTO = "#ffffff"
COR_TEXTO_SEC = "#90adc4"
COR_LOG_TEXTO = "#7ec8a0"


# ============================================================
# PERSISTÊNCIA — arquivo de configuração do usuário
# ============================================================
ARQUIVO_CONFIG = Path.home() / ".bercan_config.json"

def _config_carregar() -> dict:
    """Lê o arquivo de config; retorna dict vazio se não existir ou estiver corrompido."""
    try:
        if ARQUIVO_CONFIG.exists():
            import json
            return json.loads(ARQUIVO_CONFIG.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}

def _config_salvar(dados: dict):
    """Salva o dict de configuração em disco (merge com o que já existe)."""
    try:
        import json
        atual = _config_carregar()
        atual.update(dados)
        ARQUIVO_CONFIG.write_text(
            json.dumps(atual, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )
    except Exception:
        pass


# ============================================================
# HELPERS — CAMINHOS
# ============================================================

def resource_path(rel: str) -> str:
    """
    Resolve caminho de asset, funciona tanto em .py quanto em .exe do PyInstaller.
    O PyInstaller extrai arquivos para _MEIPASS em runtime.
    """
    base = getattr(sys, "_MEIPASS", str(Path(__file__).parent))
    return str(Path(base) / rel)


def asset(nome: str) -> str:
    """Retorna o caminho completo de um arquivo em assets/."""
    return resource_path(os.path.join("assets", nome))


def asset_checkbox(nome_base: str) -> str:
    """
    Retorna o caminho do arquivo de checkbox, tolerando extensão dupla.
    O GitHub às vezes sobe arquivos como 'CHECKBOX_X.png.jpeg'.
    Testa extensões na ordem: .png → .jpeg → .png.jpeg → .jpg
    """
    candidatos = [
        asset(nome_base),                          # ex: CHECKBOX_COM_ESGOTO.png
        asset(nome_base + ".jpeg"),                # ex: CHECKBOX_COM_ESGOTO.png.jpeg
        asset(nome_base.replace(".png", ".jpeg")), # ex: CHECKBOX_COM_ESGOTO.jpeg
        asset(nome_base.replace(".png", ".jpg")),  # ex: CHECKBOX_COM_ESGOTO.jpg
    ]
    for c in candidatos:
        if os.path.exists(c):
            return c
    # Retorna o original mesmo sem existir (vai gerar erro descritivo depois)
    return asset(nome_base)



def _detectar_posicao_esgoto(xlsx_path, log=None):
    """
    Detecta automaticamente a posição dos checkboxes de esgoto no memorial.

    Estratégia:
      1. Varre ElemConstrutivos procurando a célula com TEXTO_ITEM_ESGOTO.
      2. Na linha encontrada, identifica colunas "Sim" e "Não".
      3. No drawing XML, encontra os shapes ancorados nessa linha/coluna.

    Retorna dict:
      { "linha", "ancora_sim", "ancora_nao", "shape_sim", "shape_nao" }
    Ou None se não encontrar (app usa CHECKBOX_ANCORA_FALLBACK).
    """
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter

    try:
        wb = load_workbook(xlsx_path, data_only=False, read_only=True)
        sheet_name = "ElemConstrutivos" if "ElemConstrutivos" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]

        # Passo 1: encontrar linha pelo texto
        linha_esgoto = None
        col_sim = None
        col_nao = None

        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if TEXTO_ITEM_ESGOTO in cell.value.lower():
                        linha_esgoto = cell.row
                        break
            if linha_esgoto:
                break

        if not linha_esgoto:
            if log:
                log(f"  \u26a0 Texto nao encontrado no memorial \u2014 usando fallback")
            wb.close()
            return None

        # Passo 2: encontrar colunas SIM e NAO na linha detectada
        for cell in ws[linha_esgoto]:
            if cell.value and isinstance(cell.value, str):
                v = cell.value.strip().upper()
                if v == "SIM" and col_sim is None:
                    col_sim = cell.column
                elif v in ("N\u00c3O", "NAO") and col_nao is None:
                    col_nao = cell.column

        wb.close()

        if not col_sim:
            if log:
                log("  \u26a0 Coluna SIM nao encontrada \u2014 usando fallback")
            return None

        ancora_sim = f"{get_column_letter(col_sim)}{linha_esgoto}"
        ancora_nao = f"{get_column_letter(col_nao)}{linha_esgoto}" if col_nao else ancora_sim

        if log:
            log(f"  \u2713 Esgoto detectado: linha {linha_esgoto} | SIM={ancora_sim} NAO={ancora_nao}")

        # Passo 3: localizar shapes no drawing XML
        shape_sim = None
        shape_nao = None

        with zipfile.ZipFile(xlsx_path) as z:
            drawings = [f for f in z.namelist()
                        if f.startswith("xl/drawings/drawing") and f.endswith(".xml")]
            for drw in drawings:
                root = etree.fromstring(z.read(drw))
                ns_xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                ns_map = {"xdr": ns_xdr}

                for anchor in root.findall("xdr:twoCellAnchor", ns_map):
                    sp = anchor.find("xdr:sp", ns_map)
                    if sp is None:
                        continue
                    frm = anchor.find("xdr:from", ns_map)
                    if frm is None:
                        continue
                    row_xml = int(frm.find("xdr:row", ns_map).text) + 1
                    col_xml = int(frm.find("xdr:col", ns_map).text) + 1
                    if row_xml != linha_esgoto:
                        continue
                    cnv = sp.find(f".//{{{ns_xdr}}}cNvPr")
                    nome = cnv.get("name", "") if cnv is not None else ""
                    if col_xml == col_sim:
                        shape_sim = nome
                    elif col_nao and col_xml == col_nao:
                        shape_nao = nome

        if log:
            log(f"  \u2713 Shapes: SIM='{shape_sim}' NAO='{shape_nao}'")

        return {
            "linha": linha_esgoto,
            "ancora_sim": ancora_sim,
            "ancora_nao": ancora_nao,
            "shape_sim": shape_sim,
            "shape_nao": shape_nao,
        }

    except Exception as e:
        if log:
            log(f"  \u26a0 Deteccao automatica falhou ({e}) \u2014 usando fallback")
        return None


def formatar_data_hoje() -> str:
    """Retorna data atual no formato DD/MM/AAAA."""
    return datetime.date.today().strftime("%d/%m/%Y")


def formatar_data_extenso() -> str:
    """Retorna data no formato '16 de abril de 2026'."""
    meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho",
             "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    hoje = datetime.date.today()
    return f"{hoje.day} de {meses[hoje.month - 1]} de {hoje.year}"


# ============================================================
# HELPERS — WORD (python-docx)
# ============================================================

def _preto_run(run):
    """Força a cor de um run para preto (RGB 0,0,0)."""
    run.font.color.rgb = RGBColor(0, 0, 0)


def _preto_paragrafo(para):
    """Força todos os runs do parágrafo para preto."""
    for run in para.runs:
        _preto_run(run)


def _sub_paragrafo(para, placeholder, valor):
    """
    Substitui placeholder em um parágrafo, consolidando runs fragmentados.
    O Word quebra texto em múltiplos runs por formatação (ex: '{6}' pode estar em
    runs separados: '{', '6', '}'). Reconstruímos o texto, substituímos e
    reescrevemos tudo no primeiro run, zerando os demais.
    """
    if placeholder not in para.text:
        return False

    # Texto completo do parágrafo
    texto_completo = "".join(r.text for r in para.runs)
    if placeholder not in texto_completo:
        return False

    texto_novo = texto_completo.replace(placeholder, str(valor))

    # Reescreve tudo no primeiro run (preservando formatação dele)
    if para.runs:
        para.runs[0].text = texto_novo
        _preto_run(para.runs[0])
        # Zera os demais runs
        for run in para.runs[1:]:
            run.text = ""
    return True


def _detectar_paragrafo_assinatura(doc):
    """
    Detecta o índice do parágrafo de assinatura no template Word.
    Estratégia: procura o parágrafo que contém apenas underscores (____),
    que é a linha de assinatura em ambos os templates (FOSSA e ESGOTO).
    Fallback: usa o penúltimo parágrafo antes de "RT:".
    Retorna o índice do parágrafo encontrado.
    """
    paras = doc.paragraphs
    # Passo 1: procurar linha de underscores
    for i, p in enumerate(paras):
        txt = p.text.strip()
        if txt and all(c in ('_', ' ') for c in txt) and len(txt) >= 5:
            return i

    # Passo 2: procurar parágrafo com "RT:" e voltar 2 posições
    for i, p in enumerate(paras):
        if p.text.strip().startswith("RT:"):
            return max(0, i - 2)

    # Fallback: penúltimo parágrafo
    return max(0, len(paras) - 2)


def _inserir_assinatura_word(doc, img_path, linha_idx=None, log=None):
    """
    Insere a assinatura como imagem FLUTUANTE (behind text) no Word.
    Ancora dinamicamente no parágrafo com ____ (linha de assinatura).
    linha_idx ignorado — mantido apenas por compatibilidade.
    """
    if not os.path.exists(img_path):
        if log:
            log(f"⚠ Assinatura não encontrada: {img_path}")
        return

    idx = _detectar_paragrafo_assinatura(doc)
    if log:
        log(f"  • Assinatura ancorada no parágrafo [{idx}]: {repr(doc.paragraphs[idx].text[:40])}")

    target = doc.paragraphs[idx]
    run = target.add_run()
    run.add_picture(img_path, width=ASSINATURA_WORD_LARGURA)

    # Converter inline → anchor (flutuante behind text)
    drawing = run._r.find(qn("w:drawing"))
    if drawing is None:
        return

    inline = drawing.find(qn("wp:inline"))
    if inline is None:
        return  # já está como anchor

    # Copiar o elemento <a:graphic> do inline
    graphic_elems = [c for c in inline if "graphic" in c.tag]
    if not graphic_elems:
        return
    graphic_el = graphic_elems[0]

    # Extent (dimensões) do inline
    extent_el = inline.find(qn("wp:extent"))
    cx = extent_el.get("cx") if extent_el is not None else "1800000"
    cy = extent_el.get("cy") if extent_el is not None else "600000"

    anchor_xml = f'''<wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        distT="0" distB="0" distL="0" distR="0"
        simplePos="0" relativeHeight="251658240"
        behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="column">
            <wp:posOffset>0</wp:posOffset>
        </wp:positionH>
        <wp:positionV relativeFrom="paragraph">
            <wp:posOffset>-457200</wp:posOffset>
        </wp:positionV>
        <wp:extent cx="{cx}" cy="{cy}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:wrapNone/>
        <wp:docPr id="1" name="Assinatura"/>
        <wp:cNvGraphicFramePr/>
    </wp:anchor>'''

    anchor = etree.fromstring(anchor_xml)
    anchor.append(deepcopy(graphic_el))
    drawing.remove(inline)
    drawing.append(anchor)


def preencher_word(esgoto_sim, saida_path, dados, num_casa, log=None):
    """Preenche o template Word (FOSSA ou ESGOTO) e salva como .docx."""
    tpl_nome = TEMPLATE_ESGOTO if esgoto_sim else TEMPLATE_FOSSA
    tpl_path = asset(tpl_nome)
    linha_ass = ESGOTO_LINHA_ASS if esgoto_sim else FOSSA_LINHA_ASS

    if not os.path.exists(tpl_path):
        raise FileNotFoundError(f"Template Word não encontrado: {tpl_path}")

    doc = Document(tpl_path)

    subs = {
        "{1}": dados.get("art", ""),
        "{2}": dados.get("crea", ""),
        "{5}": dados.get("logradouro", ""),
        "{6}": dados.get("quadra_lote", ""),
        "{7}": dados.get("bairro", ""),
        "{9}": f"CASA {num_casa}",
        "{10}": dados.get("cidade", ""),
        "{11}": dados.get("uf", ""),
        "{ENGENHEIRO SELECIONADO}": dados.get("engenheiro_nome", ""),
        "{dia/mes/ano}": formatar_data_hoje(),
    }

    # Substituir em parágrafos
    for para in doc.paragraphs:
        for ph, val in subs.items():
            _sub_paragrafo(para, ph, val)
        _preto_paragrafo(para)

    # Substituir em tabelas também (caso os placeholders estejam em células)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for ph, val in subs.items():
                        _sub_paragrafo(para, ph, val)
                    _preto_paragrafo(para)

    # Inserir assinatura
    _inserir_assinatura_word(doc, dados["assinatura_path"], linha_ass, log)

    doc.save(saida_path)
    if log:
        log(f"  ✓ Word gerado: {os.path.basename(saida_path)}")


# ============================================================
# HELPERS — EXCEL (win32com)
# ============================================================

def _safe_rgb_to_hex(color_obj):
    """
    Converte objeto de cor do openpyxl para hex string, tolerante a float/tint.
    Corrige o bug 'unsupported operand type(s) for &: float and int'.
    """
    if color_obj is None:
        return None
    try:
        rgb = color_obj.rgb
        if rgb is None:
            return None
        # rgb pode vir como string 'FF1E2A3A' ou como int/float
        if isinstance(rgb, (int, float)):
            # Converter float para int antes de qualquer operação bitwise
            rgb_int = int(rgb) & 0xFFFFFF
            return f"{rgb_int:06X}"
        if isinstance(rgb, str):
            # Remove canal alpha se presente (primeiros 2 chars)
            s = rgb.upper()
            if len(s) == 8:
                s = s[2:]
            return s
        return None
    except Exception:
        return None


def _celula_tem_fundo_azul_caixa(cell):
    """
    Detecta se uma célula tem fundo azul (padrão CAIXA) no Modo Não Mapeado.
    Usa conversão segura que não falha com cores em formato float/tint.
    """
    try:
        fill = cell.fill
        if fill is None or fill.fgColor is None:
            return False
        hex_cor = _safe_rgb_to_hex(fill.fgColor)
        if not hex_cor:
            return False
        # Azul CAIXA é tipicamente #DCE6F1 ou variações claras
        # Heurística: R < G < B e diferença B-R > 15
        try:
            r = int(hex_cor[0:2], 16)
            g = int(hex_cor[2:4], 16)
            b = int(hex_cor[4:6], 16)
        except ValueError:
            return False
        return (b > r + 10) and (b > 200) and (r < 240)
    except Exception:
        return False


def _detectar_celulas_azuis_openpyxl(xlsx_path, sheet_name="ElemConstrutivos"):
    """
    Percorre o .xlsx com openpyxl e retorna lista de coordenadas com fundo azul.
    Usado no Modo Não Mapeado.
    """
    from openpyxl import load_workbook
    wb = load_workbook(xlsx_path, data_only=False)
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    coords = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None and _celula_tem_fundo_azul_caixa(cell):
                coords.append(cell.coordinate)
    wb.close()
    return coords


def _obter_fator_dpi() -> float:
    """
    Retorna fator de correção DPI relativo a 125% (escala de calibração).
    Se o PC estiver em 125% (mesmo que o de calibração) → fator 1.0 → sem ajuste.
    Se estiver em 150% → fator 1.2 → offsets aumentam proporcionalmente.
    Se estiver em 100% → fator 0.8 → offsets diminuem.
    """
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        hdc = ctypes.windll.user32.GetDC(0)
        dpi = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)  # LOGPIXELSX
        ctypes.windll.user32.ReleaseDC(0, hdc)
        # Normaliza em relação a 120 DPI (= 125% no Windows)
        return dpi / 120.0
    except Exception:
        return 1.0  # fallback — sem ajuste


def _inserir_imagem_excel_win32(ws, img_path, ancora_celula,
                                 offset_x_pt=0, offset_y_pt=0,
                                 largura_pt=100, altura_pt=50):
    """
    Insere uma imagem no Excel via win32com Shapes.AddPicture com controle
    preciso de posição e tamanho em pontos.

    Os offsets são calibrados em 100% DPI. Em PCs com escala diferente
    (125%, 150%), o Excel escala cell.Left/Top automaticamente mas os offsets
    precisam ser ajustados pelo mesmo fator para manter o alinhamento.
    """
    if not os.path.exists(img_path):
        raise FileNotFoundError(f"Imagem não encontrada: {img_path}")

    # Fator DPI: 1.0 em 100%, 1.25 em 125%, 1.5 em 150%
    # cell.Left/Top já vêm escalados pelo Excel — os offsets precisam acompanhar
    dpi = _obter_fator_dpi()

    cell = ws.Range(ancora_celula)
    left = cell.Left + (offset_x_pt * dpi)
    top  = cell.Top  + (offset_y_pt * dpi)

    # Shapes.AddPicture(Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
    shape = ws.Shapes.AddPicture(
        os.path.abspath(img_path),
        False,   # LinkToFile = False (imagem embutida)
        True,    # SaveWithDocument = True
        left,
        top,
        largura_pt * dpi,
        altura_pt  * dpi,
    )
    return shape


def _marcar_checkboxes_nativos(xlsx_path, esgoto_sim, log=None,
                               shape_sim=None, shape_nao=None):
    """
    Marca os checkboxes nativos do Excel manipulando diretamente o XML dos shapes.
    Usa lxml (preserva namespaces) para evitar corrupção do arquivo.

    shape_sim / shape_nao: nomes detectados automaticamente por
    _detectar_posicao_esgoto(). Se None, busca qualquer shape na linha
    do esgoto que não seja imagem (pic).

    Retorna True se conseguiu marcar com sucesso, False caso contrário.
    """
    import tempfile

    try:
        # .xlsx é um ZIP — vamos abrir, modificar drawing1.xml, fechar
        tmp_fd, tmp_path = tempfile.mkstemp(suffix=".xlsx")
        os.close(tmp_fd)
        shutil.copy2(xlsx_path, tmp_path)

        shapes_encontrados = 0

        with zipfile.ZipFile(tmp_path, "r") as zi, \
             zipfile.ZipFile(xlsx_path, "w", zipfile.ZIP_DEFLATED) as zo:

            for item in zi.infolist():
                data = zi.read(item.filename)

                # Procuramos em todos os drawing*.xml (pode haver mais de um)
                if item.filename.startswith("xl/drawings/drawing") and \
                   item.filename.endswith(".xml"):
                    try:
                        root = etree.fromstring(data)

                        # Mapear namespaces para busca
                        nsmap = {
                            'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
                            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                        }

                        # Percorrer todos os shapes
                        for sp in root.iter('{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}sp'):
                            # Nome do shape em xdr:nvSpPr/xdr:cNvPr/@name
                            cnv_pr = sp.find('.//xdr:nvSpPr/xdr:cNvPr', nsmap)
                            if cnv_pr is None:
                                continue
                            nome = cnv_pr.get('name', '')

                            # Determinar qual cor aplicar
                            # Usa nomes detectados automaticamente (shape_sim/shape_nao).
                            # Fallback: compara com padrão hardcoded se não foram detectados.
                            cor = None
                            _sim = shape_sim or "QO012,12.L0C0;L0C-34^"
                            _nao = shape_nao or "QO012,22.L0C0;L0C-37^"
                            if nome == _sim:
                                cor = "000000" if esgoto_sim else "FFFFFF"
                            elif nome == _nao:
                                cor = "000000" if not esgoto_sim else "FFFFFF"
                            # Geminadas loteamentos (linha 64)
                            elif nome == "QOCI,13.L0C-32;L0C-34^":
                                cor = "000000" if GEMINADAS_LOTEAMENTOS == "sim" else "FFFFFF"
                            elif nome == "QOCI,23.L0C-35;L0C-37^":
                                cor = "000000" if GEMINADAS_LOTEAMENTOS == "nao" else "FFFFFF"
                            elif nome == "QOCI,33.L0C-38;L0C-40^":
                                cor = "000000" if GEMINADAS_LOTEAMENTOS == "nao_se_aplica" else "FFFFFF"
                            # Geminadas condomínios (linha 65)
                            elif nome == "QOCN,13.L0C-32;L0C-34^":
                                cor = "000000" if GEMINADAS_CONDOMINIOS == "sim" else "FFFFFF"
                            elif nome == "QOCN,23.L0C-35;L0C-37^":
                                cor = "000000" if GEMINADAS_CONDOMINIOS == "nao" else "FFFFFF"
                            elif nome == "QOCN,33.L0C-38;L0C-40^":
                                cor = "000000" if GEMINADAS_CONDOMINIOS == "nao_se_aplica" else "FFFFFF"

                            if cor is None:
                                continue

                            shapes_encontrados += 1

                            # Buscar o solidFill dentro do shape e modificar a cor
                            # Estrutura: xdr:sp/xdr:spPr/a:solidFill/a:srgbClr
                            sp_pr = sp.find('.//xdr:spPr', nsmap)
                            if sp_pr is None:
                                continue

                            # Remover fills existentes para aplicar novo limpo
                            for fill_type in ['a:solidFill', 'a:noFill', 'a:gradFill']:
                                existing = sp_pr.find(fill_type, nsmap)
                                if existing is not None:
                                    sp_pr.remove(existing)

                            # Criar novo solidFill com a cor
                            solid_fill = etree.SubElement(
                                sp_pr,
                                '{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill'
                            )
                            srgb_clr = etree.SubElement(
                                solid_fill,
                                '{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr'
                            )
                            srgb_clr.set('val', cor)

                            # Reordenar: solidFill precisa vir em posição específica
                            # dentro de spPr. Movemos para logo após xfrm se existir,
                            # senão no início de spPr.
                            sp_pr.remove(solid_fill)
                            xfrm = sp_pr.find('a:xfrm', nsmap)
                            if xfrm is not None:
                                xfrm.addnext(solid_fill)
                            else:
                                sp_pr.insert(0, solid_fill)

                        # Serializar com lxml (PRESERVA prefixos de namespace)
                        data = etree.tostring(
                            root,
                            xml_declaration=True,
                            encoding="UTF-8",
                            standalone=True,
                        )
                    except Exception as e:
                        if log:
                            log(f"  ⚠ Erro ao processar {item.filename}: {e}")

                zo.writestr(item, data)

        os.unlink(tmp_path)

        if shapes_encontrados == 0:
            if log:
                log("  ⚠ Shapes de checkbox NÃO encontrados no XML")
            return False

        if log:
            log(f"  ✓ {shapes_encontrados} shape(s) de checkbox modificado(s) nativamente")
        return True

    except Exception as e:
        if log:
            log(f"  ✗ Falha no método nativo: {e}")
        return False


def _fechar_excel(xl, wb):
    """Fecha wb e xl com segurança — sempre mata EXCEL.EXE."""
    if wb is not None:
        try: wb.Close(SaveChanges=False)
        except Exception: pass
    if xl is not None:
        try: xl.Quit()
        except Exception: pass
        try:
            import subprocess
            subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"],
                           capture_output=True, creationflags=0x08000000)
        except Exception: pass


def _excel_preencher(template_path, xlsx_saida, dados, num_casa,
                     modo_mapeado, esgoto_sim, modo_checkbox="auto", log=None):
    """
    Preenche o Memorial Excel via win32com (nativo, preserva tudo).

    Args:
        modo_mapeado: True = células fixas; False = detectar azul
        esgoto_sim: True = sistema com esgoto público
        modo_checkbox: "nativo" | "imagem" | "auto"
            - nativo: manipula shapes existentes via XML (preserva visual original)
            - imagem: sobrepõe PNG na posição (fallback confiável)
            - auto: tenta nativo, se falhar usa imagem
    """
    pythoncom.CoInitialize()
    xl = None
    wb = None
    try:
        xl = win32com.client.Dispatch("Excel.Application")
        try: xl.Visible = False
        except Exception: pass
        try: xl.DisplayAlerts = False
        except Exception: pass
        try: xl.ScreenUpdating = False
        except Exception: pass

        wb = xl.Workbooks.Open(os.path.abspath(template_path))

        # Localizar a aba correta
        try:
            ws = wb.Worksheets("ElemConstrutivos")
        except Exception:
            ws = wb.Worksheets(1)
            if log:
                log(f"  ⚠ Aba 'ElemConstrutivos' não encontrada, usando: {ws.Name}")

        if modo_mapeado:
            mapa = {
                "G40": dados.get("contratante", ""),
                "G43": dados.get("engenheiro_nome", ""),
                "AH43": dados.get("crea", ""),
                "AP43": "GO",
                "AR43": dados.get("cpf", ""),
                "G47": dados.get("logradouro", ""),
                "AJ47": f"{dados.get('quadra_lote', '')}   CASA {num_casa}",
                "G49": dados.get("bairro", ""),
                "V49": dados.get("cep", ""),
                "AA49": dados.get("cidade", ""),
                "AU49": dados.get("uf", ""),
                "H53": dados.get("engenheiro_nome", ""),
                "Y54": dados.get("art", ""),
                "H75": f"GOIÂNIA, {formatar_data_extenso()}",
                "AE77": dados.get("engenheiro_nome", ""),
                "AE78": dados.get("cpf", ""),
                "AE79": dados.get("crea", ""),
            }
            for coord, val in mapa.items():
                try:
                    rng = ws.Range(coord)
                    rng.Value = val
                    rng.Font.Color = 0  # preto (RGB 0,0,0)
                except Exception as e:
                    if log:
                        log(f"  ⚠ Célula {coord} falhou: {e}")


        # ===================================================================
        # CHECKBOXES: feito APÓS salvar (fora do win32com)
        # Isso porque o método nativo (XML) precisa do arquivo salvo,
        # e o método imagem é mais simples de fazer num passo separado.
        # ===================================================================

        # ===================================================================
        # INSERIR ASSINATURA DO ENGENHEIRO (via win32com — qualidade preservada)
        # ===================================================================
        ass_path = dados.get("assinatura_path", "")
        if ass_path and os.path.exists(ass_path):
            try:
                _inserir_imagem_excel_win32(
                    ws,
                    ass_path,
                    ASSINATURA_EXCEL_ANCORA,
                    ASSINATURA_EXCEL_OFFSET_X_PT,
                    ASSINATURA_EXCEL_OFFSET_Y_PT,
                    ASSINATURA_EXCEL_LARGURA_PT,
                    ASSINATURA_EXCEL_ALTURA_PT,
                )
                if log:
                    log("  ✓ Assinatura do engenheiro inserida (alta qualidade)")
            except Exception as e:
                if log:
                    log(f"  ⚠ Falha ao inserir assinatura: {e}")

        # Salvar como .xlsx (51 = xlOpenXMLWorkbook)
        wb.SaveAs(os.path.abspath(xlsx_saida), FileFormat=51)

    finally:
        _fechar_excel(xl, wb)
        pythoncom.CoUninitialize()


def _quadrado_preto_temp():
    """Gera PNG temporário de quadrado preto sólido (■)."""
    import tempfile
    from PIL import Image
    tmp = tempfile.mktemp(suffix=".png")
    Image.new("RGBA", (20, 20), (0, 0, 0, 255)).save(tmp)
    return tmp


def _aplicar_checkboxes(xlsx_path, esgoto_sim, modo_checkbox="auto", log=None):
    """
    Aplica a marcação de esgoto SIM/NÃO no Memorial.

    modo_checkbox:
        "nativo" — só tenta XML (se falhar, fica sem marcação)
        "imagem" — só sobrepõe PNG
        "auto"   — tenta nativo; se falhar, cai pra imagem

    Retorna string com o método que funcionou: "nativo", "imagem" ou "nenhum".
    """
    metodo_usado = "imagem"

    # Inserir quadrado preto calibrado para cada estado
    # Estado determinado por esgoto_sim + geminadas
    q_preto = _quadrado_preto_temp()

    pythoncom.CoInitialize()
    xl = None
    wb = None
    try:
        xl = win32com.client.Dispatch("Excel.Application")
        try: xl.Visible = False
        except Exception: pass
        try: xl.DisplayAlerts = False
        except Exception: pass
        wb = xl.Workbooks.Open(os.path.abspath(xlsx_path))
        try:
            ws = wb.Worksheets("ElemConstrutivos")
        except Exception:
            ws = wb.Worksheets(1)

        # Estado 1 ou 2: esgoto
        if esgoto_sim:
            _inserir_imagem_excel_win32(ws, q_preto, CHK1_ANCORA,
                                        CHK1_OFF_X, CHK1_OFF_Y, CHK1_LARGURA, CHK1_ALTURA)
            if log: log(f"  ✓ Checkbox esgoto SIM inserido em {CHK1_ANCORA}")
        else:
            _inserir_imagem_excel_win32(ws, q_preto, CHK2_ANCORA,
                                        CHK2_OFF_X, CHK2_OFF_Y, CHK2_LARGURA, CHK2_ALTURA)
            if log: log(f"  ✓ Checkbox esgoto NÃO inserido em {CHK2_ANCORA}")

        # Estado 3 ou 4: condomínios
        if GEMINADAS_CONDOMINIOS == "sim":
            _inserir_imagem_excel_win32(ws, q_preto, CHK3_ANCORA,
                                        CHK3_OFF_X, CHK3_OFF_Y, CHK3_LARGURA, CHK3_ALTURA)
            if log: log(f"  ✓ Checkbox condomínio SIM inserido em {CHK3_ANCORA}")
        elif GEMINADAS_CONDOMINIOS == "nao_se_aplica":
            _inserir_imagem_excel_win32(ws, q_preto, CHK4_ANCORA,
                                        CHK4_OFF_X, CHK4_OFF_Y, CHK4_LARGURA, CHK4_ALTURA)
            if log: log(f"  ✓ Checkbox condomínio NSA inserido em {CHK4_ANCORA}")

        wb.Save()
        if log: log("  ✓ Checkboxes inseridos")

    except Exception as e:
        if log: log(f"  ⚠ Falha ao inserir checkboxes: {e}")
    finally:
        _fechar_excel(xl, wb)
        pythoncom.CoUninitialize()

    return metodo_usado


def _excel_para_pdf(xlsx_path, pdf_path, log=None):
    """Exporta Excel para PDF em 1 página via win32com."""
    pythoncom.CoInitialize()
    xl = None
    wb = None
    try:
        xl = win32com.client.Dispatch("Excel.Application")
        try: xl.Visible = False
        except Exception: pass
        try: xl.DisplayAlerts = False
        except Exception: pass
        wb = xl.Workbooks.Open(os.path.abspath(xlsx_path))

        try:
            ws = wb.Worksheets("ElemConstrutivos")
        except Exception:
            ws = wb.Worksheets(1)

        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.FitToPagesTall = 1

        # ExportAsFixedFormat(Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas)
        ws.ExportAsFixedFormat(
            0,   # xlTypePDF
            os.path.abspath(pdf_path),
            0,   # xlQualityStandard
            True,
            False,
        )
        if log:
            log(f"  ✓ PDF gerado: {os.path.basename(pdf_path)}")

    finally:
        _fechar_excel(xl, wb)
        pythoncom.CoUninitialize()


def _word_para_pdf(docx_path, pdf_path, log=None):
    """Converte Word para PDF via win32com."""
    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        try: word.Visible = False
        except Exception: pass
        try: word.DisplayAlerts = False
        except Exception: pass
        doc = word.Documents.Open(os.path.abspath(docx_path))
        # 17 = wdFormatPDF — SaveAs2 é o método correto em Word 2013+
        # Fallback para SaveAs em versões mais antigas
        try:
            doc.SaveAs2(os.path.abspath(pdf_path), FileFormat=17)
        except AttributeError:
            doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
        if log:
            log(f"  ✓ PDF gerado: {os.path.basename(pdf_path)}")
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


# ============================================================
# OCR DA ART (melhorado v4)
# ============================================================

def _detectar_tesseract(log=None):
    """
    Procura o executável Tesseract em múltiplos locais possíveis.
    Retorna o caminho se encontrado, None caso contrário.
    """
    candidatos = [
        # Embutido no .exe (prioridade máxima)
        resource_path(os.path.join("tesseract", "tesseract.exe")),
        resource_path(os.path.join("assets", "tesseract", "tesseract.exe")),
        # Instalado pelo Chocolatey (GitHub Actions)
        r"C:\ProgramData\chocolatey\lib\tesseract\tools\tesseract.exe",
        r"C:\ProgramData\chocolatey\bin\tesseract.exe",
        # Instalação padrão UB-Mannheim
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        # PATH do sistema
        "tesseract",
    ]

    for caminho in candidatos:
        if caminho == "tesseract":
            # Tentar executar direto do PATH
            try:
                import subprocess
                subprocess.run(
                    ["tesseract", "--version"],
                    capture_output=True,
                    timeout=5,
                    creationflags=0x08000000,  # CREATE_NO_WINDOW
                )
                if log:
                    log("  ✓ Tesseract detectado no PATH do sistema")
                return "tesseract"
            except Exception:
                continue
        elif os.path.exists(caminho):
            if log:
                log(f"  ✓ Tesseract detectado em: {caminho}")
            return caminho

    if log:
        log("  ✗ Tesseract NÃO encontrado em nenhum local conhecido")
    return None


def _detectar_tessdata(log=None):
    """Localiza a pasta tessdata (idiomas do Tesseract)."""
    candidatos = [
        resource_path(os.path.join("tesseract", "tessdata")),
        resource_path(os.path.join("assets", "tesseract", "tessdata")),
        r"C:\ProgramData\chocolatey\lib\tesseract\tools\tessdata",
        r"C:\Program Files\Tesseract-OCR\tessdata",
        r"C:\Program Files (x86)\Tesseract-OCR\tessdata",
    ]
    for caminho in candidatos:
        if os.path.exists(caminho):
            # Verifica se tem ao menos o por.traineddata
            if os.path.exists(os.path.join(caminho, "por.traineddata")):
                if log:
                    log(f"  ✓ tessdata encontrado em: {caminho}")
                return caminho
    if log:
        log("  ⚠ tessdata 'por.traineddata' não encontrado — OCR pode falhar")
    return None


def _preprocessar_imagem_ocr(pil_img):
    """
    Pré-processa imagem para melhorar precisão do OCR:
    1. Converte para escala de cinza
    2. Upscale 3x com LANCZOS
    3. Aumenta contraste (autocontrast)
    4. Binariza (preto/branco puro)
    5. Aplica filtro de nitidez
    """
    img = pil_img.convert("L")  # escala de cinza
    # Upscale 3x
    img = img.resize((img.width * 3, img.height * 3), Image.LANCZOS)
    # Autocontraste
    img = ImageOps.autocontrast(img, cutoff=2)
    # Nitidez
    img = img.filter(ImageFilter.SHARPEN)
    # Binarização (threshold 160 funciona bem para documentos)
    img = img.point(lambda p: 255 if p > 160 else 0)
    return img


def ler_art_ocr(pdf_path, log=None):
    """
    Lê os campos da ART via OCR.
    Retorna dict com os campos extraídos (ou {} se falhar).
    """
    try:
        import pytesseract
        import fitz  # PyMuPDF
    except ImportError as e:
        if log:
            log(f"  ✗ Bibliotecas de OCR ausentes: {e}")
        return {}

    # Detectar Tesseract
    tess_cmd = _detectar_tesseract(log=log)
    if tess_cmd is None:
        if log:
            log("  ✗ Tesseract não instalado. Instale de: https://github.com/UB-Mannheim/tesseract/releases")
        return {}
    pytesseract.pytesseract.tesseract_cmd = tess_cmd

    # Detectar tessdata
    tessdata = _detectar_tessdata(log=log)
    cfg_tessdata = f'--tessdata-dir "{tessdata}"' if tessdata else ""

    # Renderizar PDF como imagem (300 DPI equivalente)
    try:
        doc = fitz.open(pdf_path)
        if len(doc) == 0:
            if log:
                log("  ✗ PDF vazio")
            return {}
        # Matrix(3.0, 3.0) = 216 DPI se PDF é 72 DPI
        pix = doc[0].get_pixmap(matrix=fitz.Matrix(3.0, 3.0))
        img = Image.open(BytesIO(pix.tobytes("png")))
        doc.close()
    except Exception as e:
        if log:
            log(f"  ✗ Erro ao renderizar PDF: {e}")
        return {}

    # Pré-processar imagem (MELHORIA v4)
    img = _preprocessar_imagem_ocr(img)

    # Executar OCR
    try:
        # PSM 6 = assume single uniform block of text (bom para ARTs)
        config = f'{cfg_tessdata} --psm 6 --oem 3'
        texto = pytesseract.image_to_string(img, lang="por+eng", config=config)
    except Exception as e:
        if log:
            log(f"  ✗ Erro no OCR: {e}")
        return {}

    if log:
        log(f"  • OCR extraiu {len(texto)} caracteres")

    return _extrair_campos_art(texto, log=log)


def _extrair_campos_art(texto, log=None):
    """
    Extrai campos da ART a partir do texto OCR.
    Regex TOLERANTES a variações de formato (v4).
    """
    campos = {}

    # Normaliza texto — remove múltiplos espaços
    texto_norm = re.sub(r"[ \t]+", " ", texto)

    # ART — 13 dígitos (pode ter espaços/pontos no meio)
    # Ex: "ART nº 1022722034987" ou "Nº da ART: 10 22 72 20 34 987"
    m = re.search(
        r"(?:ART|N[º°\.]?\s*(?:da\s*)?ART|N[ºº]mero)\s*[:\-]?\s*"
        r"([\d\.\s\-/]{10,20})",
        texto_norm, flags=re.IGNORECASE,
    )
    if m:
        num = re.sub(r"[^\d]", "", m.group(1))
        if 10 <= len(num) <= 15:
            campos["art"] = num

    # CREA — padrão NNNNNNNN[D/TD]-UF
    # Ex: "CREA 1022722034D-GO", "CREA:1022722034/D-GO"
    m = re.search(
        r"(\d{7,12}\s*[/\-]?\s*[A-Z]{1,3}\s*[/\-]?\s*[A-Z]{2})",
        texto_norm,
    )
    if m:
        crea = re.sub(r"\s+", "", m.group(1))
        campos["crea"] = crea

    # CEP — XXXXX-XXX ou XXXXXXXX
    m = re.search(r"(\d{5})\s*[\-\.]?\s*(\d{3})", texto_norm)
    if m:
        campos["cep"] = f"{m.group(1)}-{m.group(2)}"

    # Quadra / Lote
    # Ex: "Quadra 15 Lote 10", "QD 15 LT 10", "Qd. 15, Lt. 10"
    m = re.search(
        r"(?:quadra|qd\.?|Q)\s*[:\.]?\s*(\w{1,10})"
        r"[\s,\-e]+"
        r"(?:lote|lt\.?|L)\s*[:\.]?\s*(\w{1,10})",
        texto_norm, flags=re.IGNORECASE,
    )
    if m:
        campos["quadra_lote"] = f"QD {m.group(1).upper()} LT {m.group(2).upper()}"

    # CPF — XXX.XXX.XXX-XX
    m = re.search(r"(\d{3})\.?(\d{3})\.?(\d{3})\-?(\d{2})", texto_norm)
    if m:
        campos["cpf"] = f"{m.group(1)}.{m.group(2)}.{m.group(3)}-{m.group(4)}"

    # Cidade — procurar após "cidade" ou "município"
    m = re.search(
        r"(?:cidade|munic[ií]pio)\s*[:\-]?\s*([A-ZÀ-Ú][A-ZÀ-Úa-zà-ú\s]{2,30})",
        texto_norm, flags=re.IGNORECASE,
    )
    if m:
        campos["cidade"] = m.group(1).strip().upper()

    # Bairro
    m = re.search(
        r"(?:bairro|setor)\s*[:\-]?\s*([A-ZÀ-Ú][A-ZÀ-Úa-zà-ú\s]{2,30})",
        texto_norm, flags=re.IGNORECASE,
    )
    if m:
        campos["bairro"] = m.group(1).strip().upper()

    # Logradouro — procurar padrões "Rua X", "Av. Y", "Alameda Z"
    m = re.search(
        r"((?:rua|avenida|av\.|alameda|al\.|travessa|tv\.|rodovia|rod\.)\s+"
        r"[A-ZÀ-Ú][A-ZÀ-Úa-zà-ú\s\d]{3,50})",
        texto_norm, flags=re.IGNORECASE,
    )
    if m:
        campos["logradouro"] = m.group(1).strip().upper()

    if log:
        if campos:
            log(f"  ✓ Campos extraídos: {', '.join(campos.keys())}")
        else:
            log("  ⚠ Nenhum campo extraído — verifique qualidade do PDF")

    return campos



# ============================================================
# SCPO — FUNÇÕES DE AUTOMAÇÃO
# ============================================================

def _scpo_obter_chromedriver(log_cb=print):
    """Detecta versão do Chrome e baixa ChromeDriver compatível."""
    versao_major = "147"  # fallback — atualizar se o Chrome avançar muito
    try:
        for chave in [
            r"SOFTWARE\Google\Chrome\BLBeacon",
            r"SOFTWARE\Wow6432Node\Google\Chrome\BLBeacon",
        ]:
            try:
                with _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE, chave) as k:
                    v, _ = _winreg.QueryValueEx(k, "version")
                    versao_major = v.split(".")[0]
                    log_cb(f"  Chrome detectado: v{v}")
                    break
            except Exception:
                pass
    except Exception:
        pass

    driver_dir  = Path.home() / "AppData" / "Local" / "SCPODriver"
    driver_path = driver_dir / "chromedriver.exe"
    versao_file = driver_dir / "versao.txt"
    driver_dir.mkdir(parents=True, exist_ok=True)

    if driver_path.exists() and versao_file.exists():
        cached = versao_file.read_text().strip()
        if cached.startswith(versao_major + "."):
            log_cb(f"  ChromeDriver {cached} em cache.")
            return str(driver_path)

    log_cb(f"  Baixando ChromeDriver para Chrome {versao_major}...")
    import ssl as _ssl
    _ctx = _ssl.create_default_context()
    _ctx.check_hostname = False
    _ctx.verify_mode = _ssl.CERT_NONE
    api = "https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json"
    with _urllib_scpo.urlopen(api, timeout=15, context=_ctx) as r:
        dados = _json_scpo.loads(r.read())

    url_zip = versao_exata = None
    for v in reversed(dados["versions"]):
        if v["version"].startswith(versao_major + "."):
            for d in v.get("downloads", {}).get("chromedriver", []):
                if d["platform"] == "win64":
                    url_zip = d["url"]
                    versao_exata = v["version"]
                    break
            if url_zip:
                break

    if not url_zip:
        raise Exception(f"ChromeDriver para Chrome {versao_major} não encontrado.")

    zip_path = driver_dir / "chromedriver.zip"
    opener = _urllib_scpo.build_opener(_urllib_scpo.HTTPSHandler(context=_ctx))
    with opener.open(url_zip) as resp, open(zip_path, "wb") as f:
        f.write(resp.read())
    with _zipfile_scpo.ZipFile(zip_path, "r") as z:
        for nome in z.namelist():
            if nome.endswith("chromedriver.exe"):
                with z.open(nome) as s, open(driver_path, "wb") as d:
                    d.write(s.read())
                break
    zip_path.unlink()
    versao_file.write_text(versao_exata)
    log_cb(f"  ChromeDriver {versao_exata} instalado.")
    return str(driver_path)


def _scpo_montar_nome_obra(logradouro, quadra, lote):
    return f"RESIDENCIAL {logradouro.upper()} QUADRA {quadra} LOTE {lote}"


def _scpo_montar_observacao(logradouro, quadra, lote, n_casas,
                             esquina, rua2, ruas_casas):
    casas = []
    for i in range(1, n_casas + 1):
        label = f"CASA {i}"
        if esquina and ruas_casas and i <= len(ruas_casas) and ruas_casas[i-1]:
            label += f" SITUADA NA {ruas_casas[i-1].upper()}"
        casas.append(label)
    casas_str = ", ".join(casas)
    if not esquina:
        return (f"OBRA RESIDENCIAL UNIFAMILIAR SITUADA NA {logradouro.upper()}, "
                f"QUADRA {quadra} LOTE {lote} COMPOSTA POR: {casas_str}")
    return (f"OBRA RESIDENCIAL UNIFAMILIAR SITUADA NA {logradouro.upper()} "
            f"E {rua2.upper()}, QUADRA {quadra} LOTE {lote} "
            f"COMPOSTA POR: {casas_str}")


def _scpo_data_termino(data_inicio_str):
    """Retorna data de término = 1 mês após data_inicio_str (DD/MM/AAAA)."""
    from dateutil.relativedelta import relativedelta
    dt = datetime.datetime.strptime(data_inicio_str, "%d/%m/%Y")
    return (dt + relativedelta(months=1)).strftime("%d/%m/%Y")


def _scpo_executar(dados, step_cb, log_cb, done_cb,
                   evento_captcha, fn_habilitar_captcha,
                   evento_envio, fn_habilitar_envio):
    """Thread principal de automação SCPO."""
    import time, traceback

    LOGIN_CPF     = "038.144.411-25"
    EMAIL_FIXO    = "joaovitorcabral94@gmail.com"
    TELEFONE_FIXO = "(62)99266-5923"
    EMP_PRINCIPAL = "0"
    EMP_TERCEIROS = "5"
    URL_LOGIN     = "https://scpo.mte.gov.br/"

    driver = None
    try:
        step_cb(5, "Obtendo ChromeDriver...")
        driver_path = _scpo_obter_chromedriver(log_cb)

        options = _webdriver_scpo.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")
        driver = _webdriver_scpo.Chrome(
            service=_Service_scpo(driver_path), options=options)
        wait = _Wait_scpo(driver, 30)

        # LOGIN
        step_cb(10, "Abrindo SCPO...")
        log_cb("Abrindo SCPO...")
        driver.get(URL_LOGIN)
        wait.until(_EC_scpo.presence_of_element_located((_By_scpo.ID, "txtCPF")))
        driver.find_element(_By_scpo.ID, "txtCPF").send_keys(LOGIN_CPF)
        driver.find_element(_By_scpo.ID, "PlaceHolderConteudo_txtSenha").send_keys(dados["senha"])
        log_cb("CPF e senha preenchidos. Digite o captcha no navegador.")
        step_cb(15, "Aguardando captcha...")
        fn_habilitar_captcha()
        evento_captcha.wait()
        log_cb("Captcha confirmado. Entrando...")
        step_cb(20, "Efetuando login...")
        driver.find_element(_By_scpo.ID, "PlaceHolderConteudo_btnLogin").click()
        wait.until(_EC_scpo.presence_of_element_located(
            (_By_scpo.XPATH, "//a[contains(@onclick,'subMenu01')]")))
        log_cb("✓ Login OK!")

        # NAVEGAÇÃO
        step_cb(30, "Navegando para Comunicar Obra...")
        wait.until(_EC_scpo.element_to_be_clickable(
            (_By_scpo.XPATH, "//a[contains(@onclick,'subMenu01')]"))).click()
        time.sleep(1)
        wait.until(_EC_scpo.element_to_be_clickable(
            (_By_scpo.XPATH, "//a[contains(@href,'DeclaracaoPreviaObra/Comunicar')]"))).click()

        # TELA INTERMEDIÁRIA
        step_cb(38, "Identificando empresa...")
        chk = wait.until(_EC_scpo.presence_of_element_located(
            (_By_scpo.ID, "PlaceHolderConteudo_chkObraSemCNPJ")))
        if not chk.is_selected():
            chk.click()
        time.sleep(1)
        cpf_f = wait.until(_EC_scpo.element_to_be_clickable(
            (_By_scpo.ID, "txtCPFProprietarioObra")))
        cpf_f.clear()
        cpf_f.send_keys(LOGIN_CPF)
        driver.find_element(
            _By_scpo.ID, "PlaceHolderConteudo_btnDeclararObra").click()

        # FORMULÁRIO
        step_cb(50, "Preenchendo formulário...")
        log_cb("Preenchendo formulário...")
        wait.until(_EC_scpo.presence_of_element_located(
            (_By_scpo.ID, "txtNomeObraEmpreendimento"))).send_keys(dados["nome_obra"])
        log_cb(f"  Nome: {dados['nome_obra']}")
        driver.find_element(_By_scpo.ID, "txtEmailObra").send_keys(EMAIL_FIXO)
        driver.find_element(_By_scpo.ID, "txtTelefoneObra").send_keys(TELEFONE_FIXO)

        # CEP dígito por dígito
        cep_limpo = re.sub(r"[^0-9]", "", dados["cep"])
        campo_cep = driver.find_element(_By_scpo.ID, "txtObraCEP")
        driver.execute_script("arguments[0].value = '';", campo_cep)
        driver.execute_script("arguments[0].focus();", campo_cep)
        for digito in cep_limpo:
            campo_cep.send_keys(digito)
            time.sleep(0.05)
        time.sleep(1)
        driver.find_element(
            _By_scpo.ID, "PlaceHolderConteudo_imgPesquisarCEPObra").click()
        time.sleep(4)
        log_cb(f"  CEP: {cep_limpo}")

        # Número: SN
        try:
            n = driver.find_element(_By_scpo.ID, "txtObraNumero")
            n.clear(); n.send_keys("SN")
        except Exception: pass

        # Complemento
        try:
            c = driver.find_element(_By_scpo.XPATH,
                "//input[contains(@id,'Complemento') or contains(@id,'complemento')]")
            c.clear()
            c.send_keys(f"QUADRA {dados['quadra']} LOTE {dados['lote']}")
        except Exception: pass

        # Observação
        try:
            obs = driver.find_element(_By_scpo.XPATH,
                "//textarea[contains(@id,'Observ') or contains(@id,'observ') "
                "or contains(@id,'Descri')]")
            obs.clear()
            obs.send_keys(dados["observacao"])
            log_cb(f"  Obs: {dados['observacao'][:80]}...")
        except Exception as e:
            log_cb(f"  textarea não localizada: {e}")

        # Classe CNAE 4120-4
        step_cb(65, "Selecionando CNAE e tipo...")
        try:
            sel = _Select_scpo(driver.find_element(
                _By_scpo.ID, "PlaceHolderConteudo_cboClasseCNAE"))
            for opt in sel.options:
                if "4120" in opt.text:
                    sel.select_by_visible_text(opt.text); break
            time.sleep(1)
        except Exception as e: log_cb(f"  CNAE: {e}")

        # Subclasse 00
        try:
            sel_sub = _Select_scpo(wait.until(_EC_scpo.presence_of_element_located(
                (_By_scpo.ID, "PlaceHolderConteudo_cboSubclasse"))))
            for opt in sel_sub.options:
                if opt.text.strip().startswith("00"):
                    sel_sub.select_by_visible_text(opt.text); break
        except Exception as e: log_cb(f"  Subclasse: {e}")

        # Tipo Construção — Edifício
        try:
            sel_tipo = _Select_scpo(driver.find_element(
                _By_scpo.ID, "PlaceHolderConteudo_CboTipoConstrucao"))
            for opt in sel_tipo.options:
                if "dif" in opt.text.lower():
                    sel_tipo.select_by_visible_text(opt.text); break
        except Exception as e: log_cb(f"  Tipo Construção: {e}")

        # Tipo Obra — Privada
        try:
            driver.find_element(_By_scpo.XPATH,
                "//input[@type='radio' and contains(@id,'rivada')]").click()
        except Exception as e: log_cb(f"  Privada: {e}")

        # Característica — Construção
        try:
            driver.find_element(_By_scpo.XPATH,
                "//input[@type='radio' and "
                "(@value='Construcao' or @value='Construção' "
                "or contains(@id,'onstrucao'))]").click()
        except Exception as e: log_cb(f"  Construção: {e}")

        # FGTS — Não
        try:
            driver.find_element(
                _By_scpo.ID, "PlaceHolderConteudo_rdbFinanciamentoFGTSNao").click()
        except Exception as e: log_cb(f"  FGTS: {e}")

        # Datas
        try:
            from dateutil.relativedelta import relativedelta
            data_termino = (datetime.date.today() + relativedelta(months=1)).strftime("%d/%m/%Y")
            driver.find_element(_By_scpo.ID, "txtInicio").send_keys(dados["data_inicio"])
            driver.find_element(_By_scpo.ID, "txtTermino").send_keys(data_termino)
            log_cb(f"  Datas: {dados['data_inicio']} → {data_termino}")
        except Exception as e: log_cb(f"  Datas: {e}")

        # Empregados
        try:
            driver.find_element(
                _By_scpo.ID, "txtNumeroEmpregadosEmpresaPrincipal").send_keys(EMP_PRINCIPAL)
        except Exception as e: log_cb(f"  Emp. principal: {e}")
        try:
            driver.find_element(_By_scpo.XPATH,
                "//input[contains(@id,'Terceiros') or contains(@id,'terceiros')]"
            ).send_keys(EMP_TERCEIROS)
        except Exception as e: log_cb(f"  Emp. terceiros: {e}")

        step_cb(85, "Aguardando confirmação...")
        log_cb("✓ Formulário preenchido! Verifique o navegador.")
        log_cb("Clique 'Confirmar envio SCPO' quando estiver pronto.")
        fn_habilitar_envio()
        evento_envio.wait()

        # SUBMETER
        step_cb(92, "Submetendo...")
        log_cb("Submetendo formulário...")
        try:
            driver.find_element(
                _By_scpo.ID, "PlaceHolderConteudo_btnConfirmar").click()
            time.sleep(3)
            log_cb("✓ Formulário submetido!")
        except Exception as e:
            log_cb(f"  Erro ao submeter: {e}")

        step_cb(100, "SCPO concluído!")
        done_cb(True, "SCPO preenchido com sucesso!")

    except Exception as e:
        tb = traceback.format_exc()
        log_cb(f"✗ ERRO SCPO: {e}")
        log_cb(tb)
        done_cb(False, str(e))


# ============================================================
# INTERFACE TKINTER
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("BERÇAN PROJETOS — Preenchimento de Documentos")
        self.geometry("1100x780")
        self.configure(bg=COR_FUNDO)

        # Controle de interrupção (NOVO v4)
        self.stop_event = threading.Event()
        self.processando = False

        self._criar_widgets()

        # Restaurar configurações salvas
        cfg = _config_carregar()
        if cfg.get("scpo_senha"):
            self.var_scpo_senha.set(cfg["scpo_senha"])

    # ------------------------------------------------------------------
    # Construção da UI
    # ------------------------------------------------------------------
    def _criar_widgets(self):
        # Cabeçalho
        header = tk.Frame(self, bg=COR_FUNDO)
        header.pack(fill="x", padx=20, pady=(20, 10))

        # Logo + título lado a lado
        header_inner = tk.Frame(header, bg=COR_FUNDO)
        header_inner.pack(anchor="w")

        # Logo (carrega LOGO.png dos assets)
        self._logo_img = None
        try:
            from PIL import Image, ImageTk
            logo_path = asset("LOGO.png")
            if not os.path.exists(logo_path):
                logo_path = asset("LOGO.jpg")
            if os.path.exists(logo_path):
                img = Image.open(logo_path).convert("RGBA")
                img = img.resize((52, 52), Image.LANCZOS)
                self._logo_img = ImageTk.PhotoImage(img)
                tk.Label(header_inner, image=self._logo_img,
                         bg=COR_FUNDO).pack(side="left", padx=(0, 12))
        except Exception:
            pass  # sem logo, continua normal

        texto_frame = tk.Frame(header_inner, bg=COR_FUNDO)
        texto_frame.pack(side="left")
        tk.Label(
            texto_frame, text="BERÇAN PROJETOS",
            font=("Segoe UI", 18, "bold"),
            fg=COR_TEXTO, bg=COR_FUNDO,
        ).pack(anchor="w")
        tk.Label(
            texto_frame, text="Preenchimento Automático de Documentos",
            font=("Segoe UI", 10),
            fg=COR_TEXTO_SEC, bg=COR_FUNDO,
        ).pack(anchor="w")

        # Container principal em 2 colunas
        main = tk.Frame(self, bg=COR_FUNDO)
        main.pack(fill="both", expand=True, padx=20, pady=10)

        col_esq = tk.Frame(main, bg=COR_FUNDO)
        col_esq.pack(side="left", fill="both", expand=True, padx=(0, 10))

        col_dir = tk.Frame(main, bg=COR_FUNDO)
        col_dir.pack(side="right", fill="both", expand=True, padx=(10, 0))

        # --- Coluna esquerda ---
        self._secao_label(col_esq, "MEMORIAL EXCEL")
        self.var_memorial = tk.StringVar()
        self._campo_arquivo(col_esq, self.var_memorial, "Arquivo Memorial (.xls/.xlsx)")

        self._secao_label(col_esq, "ENGENHEIRO RESPONSÁVEL")
        self.var_engenheiro = tk.StringVar()
        combo_eng = ttk.Combobox(
            col_esq, textvariable=self.var_engenheiro,
            values=list(ENGENHEIROS.keys()),
            state="readonly", font=("Segoe UI", 10),
        )
        combo_eng.pack(fill="x", pady=3)
        combo_eng.bind("<<ComboboxSelected>>", self._preencher_eng_campos)

        self.var_cpf = tk.StringVar()
        self.var_crea = tk.StringVar()
        self._campo_simples(col_esq, self.var_cpf, "CPF")
        self._campo_simples(col_esq, self.var_crea, "CREA")

        self._secao_label(col_esq, "DADOS DA ART")

        self.var_art = tk.StringVar()
        self._campo_simples(col_esq, self.var_art, "Número da ART")
        self.var_registro_crea = tk.StringVar()
        self._campo_simples(col_esq, self.var_registro_crea, "Registro CREA")

        self.var_contratante = tk.StringVar()
        self._campo_simples(col_esq, self.var_contratante, "Contratante")

        # CEP com botão de busca automática — PRIMEIRO para auto-preencher
        self.var_cep = tk.StringVar()
        # CEP com botão de busca automática
        tk.Label(col_esq, text="CEP",
                 bg=COR_FUNDO, fg=COR_TEXTO_SEC, font=("Segoe UI", 8)).pack(anchor="w")
        frame_cep = tk.Frame(col_esq, bg=COR_FUNDO)
        frame_cep.pack(fill="x", pady=(0, 3))
        tk.Entry(frame_cep, textvariable=self.var_cep,
                 bg=COR_CAMPO, fg=COR_TEXTO, insertbackground=COR_TEXTO,
                 relief="flat", font=("Segoe UI", 10),
                 width=12).pack(side="left")
        tk.Button(frame_cep, text="🔍 Buscar CEP",
                  command=self._buscar_cep,
                  bg=COR_BOTAO, fg=COR_TEXTO, relief="flat",
                  font=("Segoe UI", 9, "bold"),
                  ).pack(side="left", padx=(6, 0))
        self.lbl_cep_status = tk.Label(frame_cep, text="",
                  bg=COR_FUNDO, fg=COR_LOG_TEXTO,
                  font=("Segoe UI", 8))
        self.lbl_cep_status.pack(side="left", padx=(6, 0))

        self.var_logradouro = tk.StringVar()
        self._campo_simples(col_esq, self.var_logradouro, "Logradouro")

        self.var_quadra_lote = tk.StringVar()
        self._campo_simples(col_esq, self.var_quadra_lote, "Quadra e Lote")

        self.var_bairro = tk.StringVar()
        self._campo_simples(col_esq, self.var_bairro, "Bairro")

        self.var_cidade = tk.StringVar()
        self._campo_simples(col_esq, self.var_cidade, "Cidade")

        self.var_uf = tk.StringVar(value="GO")
        self._campo_simples(col_esq, self.var_uf, "UF")

        # ── SCPO: campos movidos para coluna direita (scroll_frame) ──────────
        self.var_scpo_data_inicio = tk.StringVar(
            value=datetime.date.today().strftime("%d/%m/%Y")
        )
        self.var_scpo_senha = tk.StringVar()

        # --- Coluna direita com Canvas+Scroll para todo o conteúdo ---
        # Botões fixos no rodapé (fora do scroll)
        frame_botoes = tk.Frame(col_dir, bg=COR_FUNDO)
        frame_botoes.pack(side="bottom", fill="x", pady=(6,0))
        self.btn_gerar = tk.Button(
            frame_botoes, text="⚡ GERAR DOCUMENTOS",
            command=self._iniciar_geracao,
            bg=COR_BOTAO, fg=COR_TEXTO, relief="flat",
            font=("Segoe UI", 12, "bold"), padx=20, pady=10,
        )
        self.btn_gerar.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.btn_stop = tk.Button(
            frame_botoes, text="⛔ INTERROMPER",
            command=self._solicitar_stop,
            bg=COR_BOTAO_STOP, fg=COR_TEXTO, relief="flat",
            font=("Segoe UI", 12, "bold"), padx=20, pady=10,
            state="disabled",
        )
        self.btn_stop.pack(side="right", fill="x", expand=True, padx=(5, 0))

        # Botão SCPO — linha separada abaixo
        frame_botoes2 = tk.Frame(col_dir, bg=COR_FUNDO)
        frame_botoes2.pack(side="bottom", fill="x", pady=(4, 0))
        self.btn_scpo = tk.Button(
            frame_botoes2, text="🌐 PREENCHER SCPO",
            command=self._iniciar_scpo,
            bg="#1a6b3c", fg=COR_TEXTO, relief="flat",
            font=("Segoe UI", 11, "bold"), padx=20, pady=8,
        )
        self.btn_scpo.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.btn_scpo_captcha = tk.Button(
            frame_botoes2, text="✔ Código digitado",
            command=self._scpo_liberar_captcha,
            bg="#27ae60", fg=COR_TEXTO, relief="flat",
            font=("Segoe UI", 10), padx=10, pady=8,
            state="disabled",
        )
        self.btn_scpo_captcha.pack(side="left", padx=(0, 5))
        self.btn_scpo_envio = tk.Button(
            frame_botoes2, text="✔ Confirmar envio SCPO",
            command=self._scpo_liberar_envio,
            bg="#e67e22", fg=COR_TEXTO, relief="flat",
            font=("Segoe UI", 10), padx=10, pady=8,
            state="disabled",
        )
        self.btn_scpo_envio.pack(side="left")

        # Canvas scrollável para o restante da coluna direita
        canvas_dir = tk.Canvas(col_dir, bg=COR_FUNDO, highlightthickness=0)
        sb_dir = tk.Scrollbar(col_dir, orient="vertical", command=canvas_dir.yview)
        canvas_dir.configure(yscrollcommand=sb_dir.set)
        sb_dir.pack(side="right", fill="y")
        canvas_dir.pack(side="left", fill="both", expand=True)
        scroll_frame = tk.Frame(canvas_dir, bg=COR_FUNDO)
        wid_dir = canvas_dir.create_window((0, 0), window=scroll_frame, anchor="nw")
        def _on_cf(e):
            canvas_dir.configure(scrollregion=canvas_dir.bbox("all"))
            canvas_dir.itemconfig(wid_dir, width=canvas_dir.winfo_width())
        scroll_frame.bind("<Configure>", _on_cf)
        # Forçar render inicial
        self.after(50, lambda: canvas_dir.configure(
            scrollregion=canvas_dir.bbox("all")))
        self.after(100, lambda: canvas_dir.itemconfig(
            wid_dir, width=canvas_dir.winfo_width()))
        canvas_dir.bind("<MouseWheel>",
            lambda e: canvas_dir.yview_scroll(int(-1*(e.delta/120)), "units"))
        scroll_frame.bind("<MouseWheel>",
            lambda e: canvas_dir.yview_scroll(int(-1*(e.delta/120)), "units"))
        p = scroll_frame  # alias — todo widget da col_dir usa p

        # ── 1. QUANTIDADE DE CASAS ──
        self._secao_label(p, "QUANTIDADE DE CASAS")
        self.var_qtd_casas = tk.IntVar(value=1)
        tk.Spinbox(
            p, from_=1, to=50, textvariable=self.var_qtd_casas,
            width=5, bg=COR_CAMPO, fg=COR_TEXTO,
            insertbackground=COR_TEXTO, relief="flat",
        ).pack(anchor="w", pady=3)

        # ── 2. LOTE DE ESQUINA ──
        self._secao_label(p, "LOTE")
        self.var_esquina = tk.BooleanVar(value=False)
        self._chk_esquina_widget = tk.Checkbutton(
            p, text="Lote de esquina (frente para mais de uma rua)",
            variable=self.var_esquina,
            bg=COR_FUNDO, fg=COR_TEXTO, selectcolor=COR_CAMPO,
            activebackground=COR_FUNDO, activeforeground=COR_TEXTO,
            command=self._toggle_ruas_esquina,
        )
        self._chk_esquina_widget.pack(anchor="w", pady=3)
        # Frame de ruas — campos dinâmicos (um por casa)
        self.frame_ruas_esquina = tk.Frame(p, bg=COR_FUNDO)
        self._canvas_dir = canvas_dir  # referência para scroll
        self._p_dir = p                # referência para criar campos
        self._entries_ruas = []        # lista de Entry, um por casa
        # Tracer: recria campos quando qtd_casas muda
        self.var_qtd_casas.trace_add("write",
            lambda *_: self.after(100, self._rebuild_ruas_esquina))

        # ── 3. OPÇÕES ──
        self._secao_label(p, "OPÇÕES")
        self.var_esgoto = tk.BooleanVar(value=False)
        tk.Checkbutton(
            p, text="Sistema público de esgoto (SIM)",
            variable=self.var_esgoto,
            bg=COR_FUNDO, fg=COR_TEXTO, selectcolor=COR_CAMPO,
            activebackground=COR_FUNDO, activeforeground=COR_TEXTO,
        ).pack(anchor="w", pady=3)
        tk.Label(
            p, text="↳ Template Word selecionado automaticamente",
            bg=COR_FUNDO, fg=COR_TEXTO_SEC, font=("Segoe UI", 8),
        ).pack(anchor="w", padx=20)

        self._secao_label(p, "CASAS GEMINADAS")
        _og = ["Não se aplica", "Sim", "Não"]
        self.var_gem_lot = tk.StringVar(value="Não se aplica")
        tk.Label(p, text="Condomínios",
                 bg=COR_FUNDO, fg=COR_TEXTO_SEC, font=("Segoe UI", 8)).pack(anchor="w")
        self.var_gem_cond = tk.StringVar(value="Não se aplica")
        ttk.Combobox(p, textvariable=self.var_gem_cond, values=_og,
                     state="readonly", font=("Segoe UI", 9)).pack(fill="x", pady=(0, 8))

        # ── SCPO ──
        self._secao_label(p, "SCPO")
        self._campo_simples(p, self.var_scpo_data_inicio,
                            "Data de Início da Obra (DD/MM/AAAA)")
        tk.Label(p, text="Senha SCPO",
                 bg=COR_FUNDO, fg=COR_TEXTO_SEC, font=("Segoe UI", 8)
                 ).pack(anchor="w")
        frame_senha_scpo = tk.Frame(p, bg=COR_FUNDO)
        frame_senha_scpo.pack(fill="x", pady=(0, 4))
        self._ent_scpo_senha = tk.Entry(
            frame_senha_scpo, textvariable=self.var_scpo_senha,
            bg=COR_CAMPO, fg=COR_TEXTO, insertbackground=COR_TEXTO,
            relief="flat", font=("Segoe UI", 10), show="*"
        )
        self._ent_scpo_senha.pack(side="left", fill="x", expand=True)
        tk.Button(
            frame_senha_scpo, text="👁",
            bg=COR_CAMPO, fg=COR_TEXTO, relief="flat",
            command=lambda: self._ent_scpo_senha.config(
                show="" if self._ent_scpo_senha.cget("show") == "*" else "*")
        ).pack(side="left", padx=(4, 0))
        # Salvar senha automaticamente sempre que for alterada
        self.var_scpo_senha.trace_add(
            "write",
            lambda *_: _config_salvar({"scpo_senha": self.var_scpo_senha.get()})
        )

        # ── LOG ──
        self._secao_label(p, "LOG")
        frame_log = tk.Frame(p, bg=COR_LOG_FUNDO)
        frame_log.pack(fill="x", pady=3)
        sb_log = tk.Scrollbar(frame_log, orient="vertical")
        sb_log.pack(side="right", fill="y")
        self.txt_log = tk.Text(
            frame_log, height=10, bg=COR_LOG_FUNDO, fg=COR_LOG_TEXTO,
            font=("Consolas", 9), relief="flat",
            yscrollcommand=sb_log.set,
        )
        self.txt_log.pack(side="left", fill="both", expand=True)
        sb_log.config(command=self.txt_log.yview)

        # ── PROGRESSO ──
        self._secao_label(p, "PROGRESSO")
        self.progress = ttk.Progressbar(p, mode="determinate", length=400)
        self.progress.pack(fill="x", pady=3)
        self.var_status = tk.StringVar(value="Aguardando...")
        tk.Label(
            p, textvariable=self.var_status,
            bg=COR_FUNDO, fg=COR_TEXTO_SEC, font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(0, 10))

    # ------------------------------------------------------------------
    # Helpers de UI
    # ------------------------------------------------------------------
    def _secao_label(self, parent, texto):
        tk.Label(
            parent, text=texto,
            bg=COR_FUNDO, fg=COR_TEXTO, font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", pady=(10, 3))

    def _campo_simples(self, parent, var, hint):
        tk.Label(
            parent, text=hint,
            bg=COR_FUNDO, fg=COR_TEXTO_SEC, font=("Segoe UI", 8),
        ).pack(anchor="w")
        tk.Entry(
            parent, textvariable=var,
            bg=COR_CAMPO, fg=COR_TEXTO, insertbackground=COR_TEXTO,
            relief="flat", font=("Segoe UI", 10),
        ).pack(fill="x", pady=(0, 3))

    def _campo_arquivo(self, parent, var, hint):
        tk.Label(
            parent, text=hint,
            bg=COR_FUNDO, fg=COR_TEXTO_SEC, font=("Segoe UI", 8),
        ).pack(anchor="w")
        frame = tk.Frame(parent, bg=COR_FUNDO)
        frame.pack(fill="x", pady=(0, 3))
        tk.Entry(
            frame, textvariable=var,
            bg=COR_CAMPO, fg=COR_TEXTO, insertbackground=COR_TEXTO,
            relief="flat", font=("Segoe UI", 10),
        ).pack(side="left", fill="x", expand=True)
        tk.Button(
            frame, text="📁", command=lambda: self._selecionar_arquivo(
                var, [("Excel", "*.xls *.xlsx")]),
            bg=COR_BOTAO, fg=COR_TEXTO, relief="flat", font=("Segoe UI", 9),
        ).pack(side="right", padx=(5, 0))

    def _selecionar_arquivo(self, var, filetypes):
        caminho = filedialog.askopenfilename(filetypes=filetypes)
        if caminho:
            var.set(caminho)

    def _preencher_eng_campos(self, _event=None):
        nome = self.var_engenheiro.get()
        if nome in ENGENHEIROS:
            self.var_cpf.set(ENGENHEIROS[nome]["cpf"])
            self.var_crea.set(ENGENHEIROS[nome]["crea"])
            self.var_registro_crea.set(ENGENHEIROS[nome]["crea"])

    def log(self, msg):
        """Adiciona mensagem ao log em tempo real (thread-safe via after)."""
        self.after(0, self._log_insert, msg)

    def _log_insert(self, msg):
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")

    def _set_status(self, texto):
        self.after(0, self.var_status.set, texto)

    def _set_progress(self, valor):
        self.after(0, lambda: self.progress.configure(value=valor))

    # ------------------------------------------------------------------
    # Ações
    # ------------------------------------------------------------------
    def _acionar_ocr(self):
        """Dispara leitura OCR da ART em thread separada."""
        pdf = self.var_art_pdf.get().strip()
        if not pdf or not os.path.exists(pdf):
            messagebox.showerror("ART", "Selecione um PDF de ART válido.")
            return
        self.log("🔍 Iniciando OCR da ART...")
        threading.Thread(target=self._ocr_worker, args=(pdf,), daemon=True).start()

    def _ocr_worker(self, pdf):
        try:
            campos = ler_art_ocr(pdf, log=self.log)
            if not campos:
                self.log("✗ OCR não retornou dados utilizáveis")
                return
            # Preencher campos na UI (thread-safe)
            if "art" in campos:
                self.after(0, self.var_art.set, campos["art"])
            if "crea" in campos:
                self.after(0, self.var_registro_crea.set, campos["crea"])
            if "cep" in campos:
                self.after(0, self.var_cep.set, campos["cep"])
            if "quadra_lote" in campos:
                self.after(0, self.var_quadra_lote.set, campos["quadra_lote"])
            if "cidade" in campos:
                self.after(0, self.var_cidade.set, campos["cidade"])
            if "bairro" in campos:
                self.after(0, self.var_bairro.set, campos["bairro"])
            if "logradouro" in campos:
                self.after(0, self.var_logradouro.set, campos["logradouro"])
            self.log(f"✓ OCR concluído — {len(campos)} campos preenchidos")
        except Exception as e:
            self.log(f"✗ Erro OCR: {e}")
            self.log(traceback.format_exc())

    def _solicitar_stop(self):
        """Sinaliza para a thread de processamento parar (NOVO v4)."""
        if not self.processando:
            return
        self.stop_event.set()
        self.log("⚠️ Interrupção solicitada — aguardando etapa atual finalizar...")
        self.btn_stop.configure(state="disabled", text="⏳ PARANDO...")

    def _iniciar_geracao(self):
        """Valida entradas e dispara thread de processamento."""
        # Validações básicas
        if not self.var_memorial.get() or not os.path.exists(self.var_memorial.get()):
            messagebox.showerror("Erro", "Selecione um arquivo Memorial Excel válido.")
            return
        if not self.var_engenheiro.get():
            messagebox.showerror("Erro", "Selecione o engenheiro responsável.")
            return
        if not self.var_art.get().strip():
            messagebox.showerror("Erro", "Informe o número da ART.")
            return

        # Reset
        self.stop_event.clear()
        self.processando = True
        self.btn_gerar.configure(state="disabled")
        self.btn_stop.configure(state="normal", text="⛔ INTERROMPER")
        self.progress.configure(value=0)
        self.txt_log.delete("1.0", "end")

        threading.Thread(target=self._processar, daemon=True).start()

    def _buscar_cep(self):
        """Consulta ViaCEP e preenche logradouro, bairro, cidade e UF."""
        cep = self.var_cep.get().strip()
        self.lbl_cep_status.configure(text="⏳ buscando...", fg=COR_TEXTO_SEC)

        def ok(data):
            self.after(0, self.var_logradouro.set,
                       data.get("logradouro", "").upper())
            self.after(0, self.var_bairro.set,
                       data.get("bairro", "").upper())
            self.after(0, self.var_cidade.set,
                       data.get("localidade", "").upper())
            self.after(0, self.var_uf.set,
                       data.get("uf", "GO").upper())
            self.after(0, self.lbl_cep_status.configure,
                       {"text": "✓ preenchido!", "fg": COR_LOG_TEXTO})

        def erro(msg):
            self.after(0, self.lbl_cep_status.configure,
                       {"text": f"✗ {msg}", "fg": "#e74c3c"})

        buscar_cep(cep, ok, erro)

    def _toggle_ruas_esquina(self):
        """Mostra/oculta e reconstrói campos de rua por casa."""
        if self.var_esquina.get():
            self._rebuild_ruas_esquina()
            self.frame_ruas_esquina.pack(fill="x", pady=(0, 6),
                                         after=self._chk_esquina_widget)
        else:
            self.frame_ruas_esquina.pack_forget()

    def _rebuild_ruas_esquina(self):
        """Reconstrói um Entry por casa dentro do frame_ruas_esquina."""
        if not self.var_esquina.get():
            return
        # Limpar campos anteriores
        for w in self.frame_ruas_esquina.winfo_children():
            w.destroy()
        self._entries_ruas = []
        try:
            qtd = int(self.var_qtd_casas.get())
        except:
            qtd = 1
        for i in range(1, qtd + 1):
            tk.Label(self.frame_ruas_esquina,
                     text=f"Rua — Casa {i}:",
                     bg=COR_FUNDO, fg=COR_TEXTO_SEC,
                     font=("Segoe UI", 8)).pack(anchor="w")
            var_rua = tk.StringVar()
            e = tk.Entry(self.frame_ruas_esquina,
                         textvariable=var_rua,
                         bg=COR_CAMPO, fg=COR_TEXTO,
                         insertbackground=COR_TEXTO,
                         relief="flat", font=("Segoe UI", 10))
            e.pack(fill="x", pady=(0, 4))
            e.bind("<MouseWheel>", lambda ev:
                   self._canvas_dir.yview_scroll(
                       int(-1*(ev.delta/120)), "units"))
            self._entries_ruas.append(var_rua)
        # Forçar atualização do canvas
        self.after(50, lambda: self._canvas_dir.configure(
            scrollregion=self._canvas_dir.bbox("all")))

    def _get_rua_casa(self, num_casa):
        """Retorna a rua da casa N (campo individual) ou logradouro padrão."""
        if not self.var_esquina.get():
            return self.var_logradouro.get()
        idx = num_casa - 1
        if idx < len(self._entries_ruas):
            val = self._entries_ruas[idx].get().strip()
            return val if val else self.var_logradouro.get()
        return self.var_logradouro.get()

    # ── Métodos SCPO ─────────────────────────────────────────────────────────
    def _scpo_liberar_captcha(self):
        self._scpo_evento_captcha.set()
        self.after(0, lambda: self.btn_scpo_captcha.config(state="disabled"))
        self.log("  ✓ Captcha SCPO confirmado.")

    def _scpo_liberar_envio(self):
        self._scpo_evento_envio.set()
        self.after(0, lambda: self.btn_scpo_envio.config(state="disabled"))
        self.log("  ✓ Envio SCPO confirmado.")

    def _iniciar_scpo(self):
        """Valida campos e dispara automação SCPO em thread separada."""
        # Validações
        if not self.var_cep.get().strip():
            messagebox.showwarning("SCPO", "Preencha o CEP antes de iniciar o SCPO.")
            return
        if not self.var_logradouro.get().strip():
            messagebox.showwarning("SCPO", "Preencha o Logradouro antes de iniciar o SCPO.")
            return
        ql = self.var_quadra_lote.get().strip()
        if not ql:
            messagebox.showwarning("SCPO", "Preencha Quadra e Lote antes de iniciar o SCPO.")
            return
        if not self.var_scpo_data_inicio.get().strip():
            messagebox.showwarning("SCPO", "Preencha a Data de Início da Obra (DD/MM/AAAA).")
            return
        if not self.var_scpo_senha.get().strip():
            messagebox.showwarning("SCPO", "Preencha a Senha SCPO.")
            return
        try:
            datetime.datetime.strptime(self.var_scpo_data_inicio.get().strip(), "%d/%m/%Y")
        except ValueError:
            messagebox.showwarning("SCPO", "Data de Início inválida. Use DD/MM/AAAA.")
            return

        # Separar quadra e lote de "QUADRA X LOTE Y"
        import re as _re
        ql_upper = ql.upper()
        m_q = _re.search(r"QUADRA\s+(\S+)", ql_upper)
        m_l = _re.search(r"LOTE\s+(\S+)", ql_upper)
        quadra = m_q.group(1) if m_q else ql
        lote   = m_l.group(1) if m_l else ""

        # Ruas por casa (esquina)
        n_casas   = self.var_qtd_casas.get()
        ruas_casas = []
        if self.var_esquina.get():
            for i in range(1, n_casas + 1):
                ruas_casas.append(self._get_rua_casa(i))
        rua2 = ruas_casas[1] if len(ruas_casas) >= 2 else ""

        dados_scpo = {
            "senha":       self.var_scpo_senha.get().strip(),
            "cep":         self.var_cep.get().strip(),
            "logradouro":  self.var_logradouro.get().strip(),
            "quadra":      quadra,
            "lote":        lote,
            "n_casas":     n_casas,
            "esquina":     self.var_esquina.get(),
            "rua2":        rua2,
            "ruas_casas":  ruas_casas,
            "data_inicio": self.var_scpo_data_inicio.get().strip(),
            "nome_obra":   _scpo_montar_nome_obra(
                               self.var_logradouro.get().strip(), quadra, lote),
            "observacao":  _scpo_montar_observacao(
                               self.var_logradouro.get().strip(), quadra, lote,
                               n_casas, self.var_esquina.get(),
                               rua2, ruas_casas),
        }

        # Resetar eventos
        self._scpo_evento_captcha = threading.Event()
        self._scpo_evento_envio   = threading.Event()

        # Desabilitar botão, limpar log, resetar progresso
        self.btn_scpo.config(state="disabled")
        self.btn_scpo_captcha.config(state="disabled")
        self.btn_scpo_envio.config(state="disabled")
        self.progress["value"] = 0
        self.var_status.set("SCPO: iniciando...")
        self.log("\n═══ INICIANDO SCPO ═══")

        def step_cb(pct, desc):
            self.after(0, lambda: (
                self.progress.config(value=pct),
                self.var_status.set(f"SCPO: {desc}")
            ))

        def done_cb(ok, msg):
            self.after(0, self._scpo_finalizar, ok, msg)

        threading.Thread(
            target=_scpo_executar,
            args=(dados_scpo, step_cb, self.log,
                  done_cb,
                  self._scpo_evento_captcha,
                  lambda: self.after(0, lambda: self.btn_scpo_captcha.config(state="normal")),
                  self._scpo_evento_envio,
                  lambda: self.after(0, lambda: self.btn_scpo_envio.config(state="normal"))),
            daemon=True
        ).start()

    def _scpo_finalizar(self, ok, msg):
        self.btn_scpo.config(state="normal")
        self.btn_scpo_captcha.config(state="disabled")
        self.btn_scpo_envio.config(state="disabled")
        if ok:
            messagebox.showinfo("SCPO", msg)
        else:
            messagebox.showerror("SCPO — Erro", msg)

    def _check_stop(self):
        """Levanta exceção se o usuário solicitou parada."""
        if self.stop_event.is_set():
            raise InterruptedError("Processamento interrompido pelo usuário")

    # ------------------------------------------------------------------
    # Thread de processamento
    # ------------------------------------------------------------------
    def _processar(self):
        template_excel_temp = None  # inicializado antes do try para o finally
        try:
            eng_nome = self.var_engenheiro.get()
            eng_info = ENGENHEIROS[eng_nome]
            assinatura_path = asset(eng_info["assinatura"])

            if not os.path.exists(assinatura_path):
                raise FileNotFoundError(
                    f"Assinatura não encontrada: {assinatura_path}"
                )

            dados = {
                "engenheiro_nome": eng_nome,
                "cpf": self.var_cpf.get(),
                "crea": self.var_crea.get(),
                "art": self.var_art.get(),
                "contratante": self.var_contratante.get(),
                "logradouro": self._get_rua_casa(1),  # atualizado por casa no loop
                "quadra_lote": self.var_quadra_lote.get(),
                "bairro": self.var_bairro.get(),
                "cep": self.var_cep.get(),
                "cidade": self.var_cidade.get(),
                "uf": self.var_uf.get(),
                "assinatura_path": assinatura_path,
            }

            esgoto_sim = self.var_esgoto.get()
            modo_checkbox = "imagem"  # fixo — método calibrado
            modo_mapeado = True  # sempre modo mapeado (não mapeado removido)

            # Atualizar constantes globais de geminadas
            _mg = {"Não se aplica": "nao_se_aplica", "Sim": "sim", "Não": "nao"}
            global GEMINADAS_LOTEAMENTOS, GEMINADAS_CONDOMINIOS
            GEMINADAS_LOTEAMENTOS = _mg.get(self.var_gem_lot.get(), "nao_se_aplica")
            GEMINADAS_CONDOMINIOS = _mg.get(self.var_gem_cond.get(), "nao_se_aplica")
            qtd = self.var_qtd_casas.get()
            template_excel_orig = self.var_memorial.get()

            # Converter .xls → .xlsx UMA VEZ antes do loop.
            # Evita múltiplas instâncias Excel simultâneas (causa de travamento).
            if template_excel_orig.lower().endswith(".xls"):
                self.log("• Convertendo template .xls → .xlsx (uma vez)...")
                template_excel, criou_temp_tpl = _xls_para_xlsx_temp(template_excel_orig)
                if criou_temp_tpl:
                    template_excel_temp = template_excel
                self.log(f"  ✓ Template convertido")
            else:
                template_excel = template_excel_orig

            # Pasta destino — sem subpasta de data
            rua_qd_lt = f"{dados['logradouro']} {dados['quadra_lote']}".strip()
            rua_qd_lt = re.sub(r"[<>:\"/\\|?*]", "", rua_qd_lt)  # sanitizar
            pasta_saida = (
                Path.home() / "Downloads" / PASTA_DESTINO / rua_qd_lt
            )
            pasta_saida.mkdir(parents=True, exist_ok=True)
            self.log(f"📁 Pasta destino: {pasta_saida}")

            total_etapas = qtd * 4  # Word, Word→PDF, Excel, Excel→PDF
            etapa_atual = 0

            for i in range(1, qtd + 1):
                self._check_stop()
                self._set_status(f"Casa {i}/{qtd}...")
                self.log(f"\n═══ CASA {i} ═══")
                dados["logradouro"] = self._get_rua_casa(i)

                base_nome = f"CASA_{i:02d}"

                # 1. Word
                self._check_stop()
                self.log(f"• Gerando Declaração (Word)...")
                docx_path = pasta_saida / f"DECLARACAO_{base_nome}.docx"
                preencher_word(esgoto_sim, str(docx_path), dados, i, log=self.log)
                etapa_atual += 1
                self._set_progress(etapa_atual * 100 / total_etapas)

                # 2. Word → PDF
                self._check_stop()
                self.log(f"• Convertendo Declaração para PDF...")
                pdf_decl = pasta_saida / f"DECLARACAO_{base_nome}.pdf"
                _word_para_pdf(str(docx_path), str(pdf_decl), log=self.log)
                etapa_atual += 1
                self._set_progress(etapa_atual * 100 / total_etapas)

                # 3. Excel
                self._check_stop()
                self.log(f"• Preenchendo Memorial (Excel)...")
                xlsx_path = pasta_saida / f"MEMORIAL_{base_nome}.xlsx"
                _excel_preencher(
                    template_excel, str(xlsx_path), dados, i,
                    modo_mapeado, esgoto_sim, modo_checkbox=modo_checkbox,
                    log=self.log,
                )

                # 3.5 Aplicar checkboxes (método híbrido — NOVO v4)
                self._check_stop()
                self.log(f"• Aplicando checkboxes (modo: {modo_checkbox})...")
                _aplicar_checkboxes(
                    str(xlsx_path), esgoto_sim,
                    modo_checkbox=modo_checkbox, log=self.log,
                )

                etapa_atual += 1
                self._set_progress(etapa_atual * 100 / total_etapas)

                # 4. Excel → PDF
                self._check_stop()
                self.log(f"• Convertendo Memorial para PDF...")
                pdf_mem = pasta_saida / f"MEMORIAL_{base_nome}.pdf"
                _excel_para_pdf(str(xlsx_path), str(pdf_mem), log=self.log)
                etapa_atual += 1
                self._set_progress(etapa_atual * 100 / total_etapas)

            self._set_status("Concluído!")
            self._set_progress(100)
            self.log(f"\n✅ Todos os documentos gerados com sucesso!")
            self.log(f"📂 Pasta: {pasta_saida}")
            self.after(0, lambda: messagebox.showinfo(
                "Sucesso", f"Documentos gerados em:\n{pasta_saida}"))

        except InterruptedError:
            self.log("\n⛔ Processamento interrompido pelo usuário.")
            self._set_status("Interrompido.")
        except Exception as e:
            self.log(f"\n✗ ERRO: {e}")
            self.log(traceback.format_exc())
            self._set_status("Erro.")
            self.after(0, lambda: messagebox.showerror("Erro", str(e)))
        finally:
            # Limpar temp do template se criado
            try:
                if template_excel_temp and os.path.exists(template_excel_temp):
                    os.unlink(template_excel_temp)
            except Exception:
                pass
            self.processando = False
            self.stop_event.clear()
            self.after(0, lambda: self.btn_gerar.configure(state="normal"))
            self.after(0, lambda: self.btn_stop.configure(
                state="disabled", text="⛔ INTERROMPER"))


# ============================================================
# ENTRY POINT
# ============================================================
if __name__ == "__main__":
    App().mainloop()
