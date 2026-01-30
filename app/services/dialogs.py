# app/services/dialogs.py
from __future__ import annotations

import ctypes
import tkinter as tk
from tkinter import filedialog
import win32gui

# Mantém um root único do Tk (evita criar/destruir em loop)
_tk_root: tk.Tk | None = None


def _get_root() -> tk.Tk:
    global _tk_root
    if _tk_root is None:
        _tk_root = tk.Tk()
        _tk_root.withdraw()
    return _tk_root


def _try_attach_to_foreground(root: tk.Tk) -> None:
    """
    Tenta manter o diálogo do Tk "em cima" e associado ao app principal.
    Se falhar, segue normal.
    """
    try:
        hwnd_main = win32gui.GetForegroundWindow()
    except Exception:
        hwnd_main = None

    if not hwnd_main:
        return

    try:
        root.wm_attributes("-toolwindow", True)
        root.wm_attributes("-topmost", True)
        root.lift()
        root.focus_force()

        # GWL_HWNDPARENT = -8
        ctypes.windll.user32.SetWindowLongW(root.winfo_id(), -8, hwnd_main)
    except Exception:
        # silencioso por padrão
        pass


def selecionar_diretorio() -> str:
    root = _get_root()
    _try_attach_to_foreground(root)

    folder_selected = filedialog.askdirectory(
        parent=root,
        title="Selecione um diretório de armazenamento",
    )
    return folder_selected or ""


def selecionar_arquivo() -> list[str]:
    root = _get_root()
    _try_attach_to_foreground(root)

    arquivos = filedialog.askopenfilenames(
        parent=root,
        title="Selecione um ou mais arquivos TXT",
        filetypes=[("Arquivos TXT", "*.txt")],
    )
    return list(arquivos) if arquivos else []


def selecionar_planilha() -> str:
    """
    Seleciona uma planilha Excel no padrão AUTO_CL.
    Retorna caminho do arquivo ou "" se cancelado.
    """
    root = _get_root()
    _try_attach_to_foreground(root)

    file_selected = filedialog.askopenfilename(
        parent=root,
        title="Selecione a planilha (.xlsx) no padrão AUTO_CL",
        filetypes=[("Planilhas Excel", "*.xlsx *.xlsm")],
    )
    return file_selected or ""
