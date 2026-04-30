# -*- coding: utf-8 -*-
# Launcher Classic Sports
import sys, os, subprocess
from pathlib import Path

PASTA = Path(__file__).resolve().parent

# Define icone correto na barra de tarefas
try:
    import ctypes
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
        "ClassicSports.Relatorios.1.0")
except Exception:
    pass

# Executa o painel
sys.path.insert(0, str(PASTA))
os.chdir(str(PASTA))

painel = PASTA / "sportbay_painel.py"
if painel.exists():
    with open(painel, encoding="utf-8") as f:
        code = f.read()
    exec(compile(code, str(painel), "exec"), {"__file__": str(painel), "__name__": "__main__"})
