# -*- coding: utf-8 -*-
"""
CALIBRADOR DO MEMORIAL — Morais Engenharia
Ferramenta de ajuste visual para assinatura e checkbox do Memorial Excel.

USO:
  1. Execute: python calibrador_memorial.py
  2. Selecione o Memorial (.xls ou .xlsx) e a imagem de assinatura
  3. Ajuste os sliders/campos até o posicionamento ficar correto
  4. Clique "GERAR PREVIEW" para ver o resultado no Excel
  5. Copie os valores finais para o app.py principal

DEPENDÊNCIAS: python-docx, openpyxl, pywin32, pillow, lxml
"""

import multiprocessing
multiprocessing.freeze_support()

import os
import sys
import shutil
import zipfile
import tempfile
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from PIL import Image, ImageDraw
import win32com.client
import pythoncom
from lxml import etree

# ─────────────────────────────────────────────
# PALETA
# ─────────────────────────────────────────────
BG       = "#0f1923"
BG2      = "#1a2535"
CAMPO    = "#243447"
ACENTO   = "#2e86de"
ACENTO2  = "#27ae60"
TEXTO    = "#e8edf3"
TEXTO2   = "#7a99b8"
BORDA    = "#2a3f58"
LOG_BG   = "#0a1118"
LOG_FG   = "#4ecb8d"

# ─────────────────────────────────────────────
# VALORES PADRÃO (copiar do app.py atual)
# ─────────────────────────────────────────────
DEFAULT = {
    # Assinatura
    "ass_ancora":    "AE73",
    "ass_offset_x":  0,
    "ass_offset_y":  0,
    "ass_largura":   150,
    "ass_altura":    60,
    # Checkbox imagem
    "chk_ancora":    "AM70",
    "chk_offset_x":  0,
    "chk_offset_y":  0,
    "chk_largura":   85,
    "chk_altura":    14,
    "esgoto_sim":    True,
}


# ─────────────────────────────────────────────
# HELPERS COM
# ─────────────────────────────────────────────

def _xls_para_xlsx_temp(path):
    if str(path).lower().endswith(".xlsx"):
        return str(path), False
    tmp = tempfile.mktemp(suffix=".xlsx")
    pythoncom.CoInitialize()
    xl = None; wb = None
    try:
        xl = win32com.client.Dispatch("Excel.Application")
        try: xl.Visible = False
        except: pass
        try: xl.DisplayAlerts = False
        except: pass
        wb = xl.Workbooks.Open(os.path.abspath(str(path)))
        wb.SaveAs(os.path.abspath(tmp), FileFormat=51)
        return tmp, True
    finally:
        if wb:
            try: wb.Close(SaveChanges=False)
            except: pass
        if xl:
            try: xl.Quit()
            except: pass
        pythoncom.CoUninitialize()


def _criar_xl():
    """Cria instância Excel limpa, matando zumbis se necessário."""
    import subprocess, time

    def _try():
        xl = win32com.client.Dispatch("Excel.Application")
        try: xl.Visible = False
        except: pass
        try: xl.DisplayAlerts = False
        except: pass
        try: xl.ScreenUpdating = False
        except: pass
        _ = xl.Workbooks.Count  # testa se está vivo
        return xl

    try:
        return _try()
    except Exception:
        pass
    try:
        subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"],
                       capture_output=True, creationflags=0x08000000)
    except Exception:
        pass
    time.sleep(1)
    return _try()


def _inserir_imagem_win32(ws, img_path, ancora, offset_x, offset_y, largura, altura):
    cell = ws.Range(ancora)
    left = cell.Left + offset_x
    top  = cell.Top  + offset_y
    ws.Shapes.AddPicture(
        os.path.abspath(img_path),
        False, True,
        left, top, largura, altura,
    )


# ─────────────────────────────────────────────
# GERAÇÃO DO PREVIEW
# ─────────────────────────────────────────────

def gerar_preview(memorial_path, ass_img_path, cfg, saida_path, log):
    """
    Abre o memorial, insere assinatura + checkbox de preview e exporta PDF.
    Roda em thread separada.
    """
    pythoncom.CoInitialize()
    xl = None; wb = None
    xlsx_tmp = None; xlsx_criou = False

    try:
        # 1. Garantir .xlsx
        log("• Preparando arquivo...")
        xlsx_tmp, xlsx_criou = _xls_para_xlsx_temp(memorial_path)
        # Copiar para o saida_path para não modificar o original
        shutil.copy2(xlsx_tmp, saida_path)
        if xlsx_criou:
            try: os.unlink(xlsx_tmp)
            except: pass

        # 2. Abrir no Excel
        log("• Abrindo no Excel...")
        xl = _criar_xl()
        wb = xl.Workbooks.Open(os.path.abspath(saida_path))
        try:
            ws = wb.Worksheets("ElemConstrutivos")
        except Exception:
            ws = wb.Worksheets(1)

        # 3. Inserir assinatura
        if ass_img_path and os.path.exists(ass_img_path):
            log(f"• Inserindo assinatura em {cfg['ass_ancora']} "
                f"({cfg['ass_largura']}×{cfg['ass_altura']}pt)...")
            _inserir_imagem_win32(
                ws, ass_img_path,
                cfg["ass_ancora"],
                cfg["ass_offset_x"], cfg["ass_offset_y"],
                cfg["ass_largura"],  cfg["ass_altura"],
            )
        else:
            log("  ⚠ Assinatura não selecionada — pulando")

        # 4. Inserir checkbox (imagem placeholder ou real)
        chk_img = _obter_img_checkbox(cfg["esgoto_sim"])
        log(f"• Inserindo checkbox em {cfg['chk_ancora']} "
            f"({cfg['chk_largura']}×{cfg['chk_altura']}pt) "
            f"[{'COM' if cfg['esgoto_sim'] else 'SEM'} esgoto]...")
        _inserir_imagem_win32(
            ws, chk_img,
            cfg["chk_ancora"],
            cfg["chk_offset_x"], cfg["chk_offset_y"],
            cfg["chk_largura"],  cfg["chk_altura"],
        )

        # 5. Salvar e exportar PDF
        wb.Save()
        pdf_path = saida_path.replace(".xlsx", ".pdf")
        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.FitToPagesTall = 1
        ws.ExportAsFixedFormat(0, os.path.abspath(pdf_path), 0, True, False)

        log(f"✅ PDF gerado: {pdf_path}")
        log("─" * 50)
        log("VALORES PARA COPIAR PARA O app.py:")
        log(f'  ASSINATURA_EXCEL_ANCORA     = "{cfg["ass_ancora"]}"')
        log(f'  ASSINATURA_EXCEL_OFFSET_X_PT = {cfg["ass_offset_x"]}')
        log(f'  ASSINATURA_EXCEL_OFFSET_Y_PT = {cfg["ass_offset_y"]}')
        log(f'  ASSINATURA_EXCEL_LARGURA_PT  = {cfg["ass_largura"]}')
        log(f'  ASSINATURA_EXCEL_ALTURA_PT   = {cfg["ass_altura"]}')
        log(f'  CHECKBOX_ANCORA_FALLBACK     = "{cfg["chk_ancora"]}"')
        log(f'  CHECKBOX_OFFSET_X_PT         = {cfg["chk_offset_x"]}')
        log(f'  CHECKBOX_OFFSET_Y_PT         = {cfg["chk_offset_y"]}')
        log(f'  CHECKBOX_LARGURA_PT          = {cfg["chk_largura"]}')
        log(f'  CHECKBOX_ALTURA_PT           = {cfg["chk_altura"]}')

        # Abrir PDF automaticamente
        try:
            os.startfile(pdf_path)
        except Exception:
            pass

    except Exception as e:
        log(f"✗ ERRO: {e}")
        log(traceback.format_exc())
    finally:
        if wb:
            try: wb.Close(SaveChanges=False)
            except: pass
        if xl:
            try: xl.Quit()
            except: pass
            try:
                import subprocess
                subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"],
                               capture_output=True, creationflags=0x08000000)
            except: pass
        pythoncom.CoUninitialize()


def _obter_img_checkbox(esgoto_sim):
    """
    Retorna caminho de imagem de checkbox.
    Tenta encontrar no diretório do script; se não, cria placeholder.
    """
    base = Path(__file__).parent
    candidatos_sim = [
        base / "assets" / "CHECKBOX_COM_ESGOTO.png",
        base / "assets" / "CHECKBOX_COM_ESGOTO.jpeg",
        base / "assets" / "CHECKBOX_COM_ESGOTO.png.jpeg",
    ]
    candidatos_nao = [
        base / "assets" / "CHECKBOX_SEM_ESGOTO.png",
        base / "assets" / "CHECKBOX_SEM_ESGOTO.jpeg",
        base / "assets" / "CHECKBOX_SEM_ESGOTO.png.jpeg",
    ]
    lista = candidatos_sim if esgoto_sim else candidatos_nao
    for c in lista:
        if c.exists():
            return str(c)

    # Placeholder: retângulo azul com "SIM" ou "NÃO" marcado
    tmp = tempfile.mktemp(suffix=".png")
    img = Image.new("RGBA", (200, 40), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)
    draw.rectangle([2, 2, 38, 38], outline=(0, 0, 0, 255), width=2)
    draw.rectangle([102, 2, 138, 38], outline=(0, 0, 0, 255), width=2)
    if esgoto_sim:
        draw.line([5, 20, 15, 35], fill=(0, 100, 200, 255), width=3)
        draw.line([15, 35, 35, 5], fill=(0, 100, 200, 255), width=3)
        draw.text((45, 10), "Sim ✓  Não □", fill=(0, 0, 0, 255))
    else:
        draw.line([105, 20, 115, 35], fill=(0, 100, 200, 255), width=3)
        draw.line([115, 35, 135, 5], fill=(0, 100, 200, 255), width=3)
        draw.text((45, 10), "Sim □  Não ✓", fill=(0, 0, 0, 255))
    img.save(tmp)
    return tmp


# ─────────────────────────────────────────────
# INTERFACE
# ─────────────────────────────────────────────

class Calibrador(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Calibrador do Memorial — Morais Engenharia")
        self.geometry("860x780")
        self.configure(bg=BG)
        self.resizable(True, True)
        self._criar_ui()

    # ── UI ──────────────────────────────────

    def _criar_ui(self):
        # Cabeçalho
        hdr = tk.Frame(self, bg=BG, pady=14)
        hdr.pack(fill="x", padx=24)
        tk.Label(hdr, text="CALIBRADOR DO MEMORIAL",
                 font=("Segoe UI", 16, "bold"), fg=TEXTO, bg=BG).pack(anchor="w")
        tk.Label(hdr, text="Ajuste assinatura e checkbox • copie os valores para o app.py",
                 font=("Segoe UI", 9), fg=TEXTO2, bg=BG).pack(anchor="w")

        tk.Frame(self, bg=BORDA, height=1).pack(fill="x")

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=24, pady=12)

        col_esq = tk.Frame(body, bg=BG)
        col_esq.pack(side="left", fill="both", expand=True, padx=(0, 10))

        col_dir = tk.Frame(body, bg=BG)
        col_dir.pack(side="right", fill="both", expand=True, padx=(10, 0))

        # ── Coluna esquerda ──
        self._titulo(col_esq, "ARQUIVOS")
        self.var_memorial = tk.StringVar()
        self._arquivo(col_esq, self.var_memorial, "Memorial Excel (.xls/.xlsx)",
                      [("Excel", "*.xls *.xlsx")])
        self.var_assinatura = tk.StringVar()
        self._arquivo(col_esq, self.var_assinatura, "Imagem da assinatura (.png/.jpg)",
                      [("Imagem", "*.png *.jpg *.jpeg")])

        self._titulo(col_esq, "ASSINATURA NO MEMORIAL")
        self._linha_campo(col_esq, "Célula âncora", "ass_ancora",  DEFAULT["ass_ancora"])
        self._linha_spin(col_esq,  "Offset X (pt)",  "ass_off_x",  DEFAULT["ass_offset_x"], -200, 200)
        self._linha_spin(col_esq,  "Offset Y (pt)",  "ass_off_y",  DEFAULT["ass_offset_y"], -200, 200)
        self._linha_spin(col_esq,  "Largura (pt)",   "ass_larg",   DEFAULT["ass_largura"],   10, 400)
        self._linha_spin(col_esq,  "Altura (pt)",    "ass_alt",    DEFAULT["ass_altura"],    10, 200)

        tk.Label(col_esq, text="1 cm ≈ 28 pt  |  1 polegada = 72 pt",
                 bg=BG, fg=TEXTO2, font=("Segoe UI", 8)).pack(anchor="w", pady=(2, 8))

        self._titulo(col_esq, "CHECKBOX DE ESGOTO")
        self.var_esgoto = tk.BooleanVar(value=DEFAULT["esgoto_sim"])
        tk.Checkbutton(col_esq, text="Sistema público de esgoto (SIM)",
                       variable=self.var_esgoto,
                       bg=BG, fg=TEXTO, selectcolor=CAMPO,
                       activebackground=BG, activeforeground=TEXTO,
                       font=("Segoe UI", 9)).pack(anchor="w", pady=4)

        self._linha_campo(col_esq, "Célula âncora", "chk_ancora",  DEFAULT["chk_ancora"])
        self._linha_spin(col_esq,  "Offset X (pt)",  "chk_off_x",  DEFAULT["chk_offset_x"], -200, 200)
        self._linha_spin(col_esq,  "Offset Y (pt)",  "chk_off_y",  DEFAULT["chk_offset_y"], -200, 200)
        self._linha_spin(col_esq,  "Largura (pt)",   "chk_larg",   DEFAULT["chk_largura"],   5,  400)
        self._linha_spin(col_esq,  "Altura (pt)",    "chk_alt",    DEFAULT["chk_altura"],    5,  200)

        # ── Coluna direita: log ──
        self._titulo(col_dir, "LOG")
        frame_log = tk.Frame(col_dir, bg=LOG_BG)
        frame_log.pack(fill="both", expand=True, pady=4)
        sb = tk.Scrollbar(frame_log)
        sb.pack(side="right", fill="y")
        self.txt_log = tk.Text(frame_log, bg=LOG_BG, fg=LOG_FG,
                               font=("Consolas", 9), relief="flat",
                               yscrollcommand=sb.set, wrap="word")
        self.txt_log.pack(fill="both", expand=True)
        sb.config(command=self.txt_log.yview)

        self._titulo(col_dir, "VALORES PARA COPIAR")
        self.txt_copy = tk.Text(col_dir, height=12, bg=CAMPO, fg=ACENTO2,
                                font=("Consolas", 9), relief="flat")
        self.txt_copy.pack(fill="x", pady=4)
        self._atualizar_copy()

        tk.Button(col_dir, text="📋  COPIAR VALORES",
                  command=self._copiar_valores,
                  bg=BORDA, fg=TEXTO, relief="flat",
                  font=("Segoe UI", 9, "bold"), pady=6).pack(fill="x", pady=(0, 8))

        # Botão principal
        tk.Frame(self, bg=BORDA, height=1).pack(fill="x")
        rodape = tk.Frame(self, bg=BG, pady=12)
        rodape.pack(fill="x", padx=24)

        self.btn = tk.Button(rodape, text="⚡  GERAR PREVIEW (abre PDF automaticamente)",
                             command=self._iniciar,
                             bg=ACENTO, fg=TEXTO, relief="flat",
                             font=("Segoe UI", 12, "bold"), pady=12)
        self.btn.pack(fill="x")

        self.var_status = tk.StringVar(value="Aguardando...")
        tk.Label(rodape, textvariable=self.var_status,
                 bg=BG, fg=TEXTO2, font=("Segoe UI", 9)).pack(anchor="w", pady=(6, 0))

        # Vincular atualização automática dos valores
        for v in self._vars.values():
            v.trace_add("write", lambda *_: self.after(100, self._atualizar_copy))
        self.var_esgoto.trace_add("write", lambda *_: self.after(100, self._atualizar_copy))

    # ── Helpers de UI ───────────────────────

    def _titulo(self, parent, txt):
        tk.Label(parent, text=txt, font=("Segoe UI", 10, "bold"),
                 fg=TEXTO, bg=BG).pack(anchor="w", pady=(12, 4))

    def _arquivo(self, parent, var, hint, ftypes):
        tk.Label(parent, text=hint, fg=TEXTO2, bg=BG,
                 font=("Segoe UI", 8)).pack(anchor="w")
        fr = tk.Frame(parent, bg=BG)
        fr.pack(fill="x", pady=(0, 4))
        tk.Entry(fr, textvariable=var, bg=CAMPO, fg=TEXTO,
                 insertbackground=TEXTO, relief="flat",
                 font=("Segoe UI", 9)).pack(side="left", fill="x", expand=True)
        tk.Button(fr, text="📁",
                  command=lambda: self._sel(var, ftypes),
                  bg=ACENTO, fg=TEXTO, relief="flat",
                  font=("Segoe UI", 9)).pack(side="right", padx=(4, 0))

    def _sel(self, var, ftypes):
        p = filedialog.askopenfilename(filetypes=ftypes)
        if p:
            var.set(p)

    _vars: dict = {}

    def _linha_campo(self, parent, label, key, default):
        if not hasattr(self, "_vars"):
            self._vars = {}
        fr = tk.Frame(parent, bg=BG)
        fr.pack(fill="x", pady=2)
        tk.Label(fr, text=label, width=18, anchor="w",
                 fg=TEXTO2, bg=BG, font=("Segoe UI", 9)).pack(side="left")
        var = tk.StringVar(value=str(default))
        self._vars[key] = var
        tk.Entry(fr, textvariable=var, width=10, bg=CAMPO, fg=TEXTO,
                 insertbackground=TEXTO, relief="flat",
                 font=("Segoe UI", 10, "bold")).pack(side="left")

    def _linha_spin(self, parent, label, key, default, mn, mx):
        if not hasattr(self, "_vars"):
            self._vars = {}
        fr = tk.Frame(parent, bg=BG)
        fr.pack(fill="x", pady=2)
        tk.Label(fr, text=label, width=18, anchor="w",
                 fg=TEXTO2, bg=BG, font=("Segoe UI", 9)).pack(side="left")
        var = tk.StringVar(value=str(default))
        self._vars[key] = var
        tk.Spinbox(fr, textvariable=var, from_=mn, to=mx, width=8,
                   bg=CAMPO, fg=TEXTO, insertbackground=TEXTO,
                   buttonbackground=BORDA, relief="flat",
                   font=("Segoe UI", 10, "bold")).pack(side="left")
        # botões ±5
        tk.Button(fr, text="-5",
                  command=lambda k=key, v=var: self._nudge(v, -5),
                  bg=BORDA, fg=TEXTO2, relief="flat",
                  font=("Segoe UI", 8), padx=4).pack(side="left", padx=(6, 1))
        tk.Button(fr, text="+5",
                  command=lambda k=key, v=var: self._nudge(v, +5),
                  bg=BORDA, fg=TEXTO2, relief="flat",
                  font=("Segoe UI", 8), padx=4).pack(side="left", padx=1)

    def _nudge(self, var, delta):
        try:
            var.set(str(int(var.get()) + delta))
        except Exception:
            pass

    def _cfg(self):
        v = self._vars
        def i(k): return int(v[k].get())
        def s(k): return v[k].get().strip()
        return {
            "ass_ancora":  s("ass_ancora"),
            "ass_offset_x": i("ass_off_x"),
            "ass_offset_y": i("ass_off_y"),
            "ass_largura":  i("ass_larg"),
            "ass_altura":   i("ass_alt"),
            "chk_ancora":  s("chk_ancora"),
            "chk_offset_x": i("chk_off_x"),
            "chk_offset_y": i("chk_off_y"),
            "chk_largura":  i("chk_larg"),
            "chk_altura":   i("chk_alt"),
            "esgoto_sim":   self.var_esgoto.get(),
        }

    def _atualizar_copy(self):
        try:
            cfg = self._cfg()
        except Exception:
            return
        txt = (
            f'ASSINATURA_EXCEL_ANCORA      = "{cfg["ass_ancora"]}"\n'
            f'ASSINATURA_EXCEL_OFFSET_X_PT = {cfg["ass_offset_x"]}\n'
            f'ASSINATURA_EXCEL_OFFSET_Y_PT = {cfg["ass_offset_y"]}\n'
            f'ASSINATURA_EXCEL_LARGURA_PT  = {cfg["ass_largura"]}\n'
            f'ASSINATURA_EXCEL_ALTURA_PT   = {cfg["ass_altura"]}\n'
            f'CHECKBOX_ANCORA_FALLBACK     = "{cfg["chk_ancora"]}"\n'
            f'CHECKBOX_OFFSET_X_PT         = {cfg["chk_offset_x"]}\n'
            f'CHECKBOX_OFFSET_Y_PT         = {cfg["chk_offset_y"]}\n'
            f'CHECKBOX_LARGURA_PT          = {cfg["chk_largura"]}\n'
            f'CHECKBOX_ALTURA_PT           = {cfg["chk_altura"]}\n'
        )
        self.txt_copy.delete("1.0", "end")
        self.txt_copy.insert("1.0", txt)

    def _copiar_valores(self):
        self.clipboard_clear()
        self.clipboard_append(self.txt_copy.get("1.0", "end").strip())
        self.var_status.set("✓ Valores copiados para a área de transferência!")

    # ── Ação principal ──────────────────────

    def _iniciar(self):
        memorial = self.var_memorial.get().strip()
        if not memorial or not os.path.exists(memorial):
            messagebox.showerror("Erro", "Selecione um arquivo Memorial válido.")
            return
        try:
            cfg = self._cfg()
        except ValueError as e:
            messagebox.showerror("Erro", f"Valor inválido: {e}")
            return

        ass_img = self.var_assinatura.get().strip() or None

        saida = str(Path(memorial).parent / "PREVIEW_MEMORIAL_CALIBRADOR.xlsx")

        self.btn.configure(state="disabled", text="⏳  Gerando preview...")
        self.txt_log.delete("1.0", "end")
        self.var_status.set("Processando...")

        threading.Thread(
            target=self._worker,
            args=(memorial, ass_img, cfg, saida),
            daemon=True,
        ).start()

    def _worker(self, memorial, ass_img, cfg, saida):
        def log(msg):
            self.after(0, self._log_insert, msg)

        try:
            gerar_preview(memorial, ass_img, cfg, saida, log)
            self.after(0, self.var_status.set, "✅ Preview gerado — verifique o PDF!")
        except Exception as e:
            self.after(0, self.var_status.set, f"✗ Erro: {e}")
        finally:
            self.after(0, self.btn.configure, {
                "state": "normal",
                "text": "⚡  GERAR PREVIEW (abre PDF automaticamente)"
            })

    def _log_insert(self, msg):
        self.txt_log.insert("end", msg + "\n")
        self.txt_log.see("end")


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────

if __name__ == "__main__":
    Calibrador().mainloop()
