# -*- coding: utf-8 -*-
import sys
import os
import shutil
import subprocess
import threading
import hashlib
import json
import importlib.util
from pathlib import Path
from tkinter import filedialog

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

# ── PASTAS ────────────────────────────────────────────────────
PASTA       = Path(__file__).parent
PASTA_DADOS = PASTA / "dados"
PASTA_DADOS.mkdir(exist_ok=True)

# ── VERSÃO ────────────────────────────────────────────────────
VERSAO_ATUAL = "1.0.0"

# ── ATUALIZADOR ───────────────────────────────────────────────
try:
    import atualizador
    ATUALIZADOR_OK = True
except ImportError:
    ATUALIZADOR_OK = False

ARQ_USUARIOS = PASTA_DADOS / "usuarios.json"
ARQ_CONFIG   = PASTA / "config.py"

# ── CORES ─────────────────────────────────────────────────────
COR_BG      = "#001228"
COR_HEADER  = "#00142E"
COR_BLUE    = "#0070CC"
COR_CYAN    = "#00AAFF"
COR_CARD    = "#002850"
COR_BRIGHT  = "#E0ECFA"
COR_MUTED   = "#7A9DBE"
COR_SUCCESS = "#2ECC71"
COR_WARNING = "#F39C12"
COR_DANGER  = "#E74C3C"
COR_BTN     = "#0D2D4E"
COR_BTN_TXT = "#7ECFFF"

LOJAS = [
    ("Classic Barracao", "sportbay_classic_barracao.py"),
    ("CS",               "sportbay_cs.py"),
    ("Juliana",          "sportbay_juliana.py"),
    ("Marcus",           "sportbay_marcus.py"),
    ("11",               "sportbay_11.py"),
    ("12",               "sportbay_12.py"),
    ("Imports",          "sportbay_imports.py"),
    ("708",              "sportbay_708.py"),
    ("AdrenalineX",      "sportbay_adrenalinex.py"),
    ("AdventureX",       "sportbay_adventurex.py"),
    ("AM15",             "sportbay_am15.py"),
    ("AM20",             "sportbay_am20.py"),
    ("Planet",           "sportbay_planet.py"),
    ("RAS",              "sportbay_ras.py"),
    ("DL",               "sportbay_dl.py"),
    ("DM",               "sportbay_dm.py"),
    ("BL",               "sportbay_bl.py"),
    ("GJ",               "sportbay_gj.py"),
    ("FF",               "sportbay_ff.py"),
    ("JA",               "sportbay_ja.py"),
    ("SS",               "sportbay_ss.py"),
]

NOMES_LOJAS = [n for n, _ in LOJAS]

# ═════════════════════════════════════════════════════════════
# GERENCIAMENTO DE USUÁRIOS
# ═════════════════════════════════════════════════════════════
def hash_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

def carregar_usuarios():
    if not ARQ_USUARIOS.exists():
        return {}
    try:
        return json.loads(ARQ_USUARIOS.read_text(encoding="utf-8"))
    except Exception:
        return {}

def salvar_usuarios(dados):
    ARQ_USUARIOS.write_text(json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")

def usuario_existe(usuario):
    return usuario in carregar_usuarios()

def criar_usuario(usuario, senha="1234"):
    dados = carregar_usuarios()
    dados[usuario] = {
        "senha_hash": hash_senha(senha),
        "primeiro_acesso": True,
        "email_sportbay": "",
        "senha_sportbay": "",
        "lojas": {}
    }
    salvar_usuarios(dados)

def verificar_senha(usuario, senha):
    dados = carregar_usuarios()
    if usuario not in dados:
        return False
    return dados[usuario]["senha_hash"] == hash_senha(senha)

def trocar_senha(usuario, nova_senha):
    dados = carregar_usuarios()
    dados[usuario]["senha_hash"] = hash_senha(nova_senha)
    dados[usuario]["primeiro_acesso"] = False
    salvar_usuarios(dados)

def get_usuario(usuario):
    return carregar_usuarios().get(usuario, {})

def salvar_dados_usuario(usuario, campo, valor):
    dados = carregar_usuarios()
    if usuario in dados:
        dados[usuario][campo] = valor
        salvar_usuarios(dados)

def salvar_config_loja(usuario, loja, email, senha):
    dados = carregar_usuarios()
    if usuario in dados:
        if "lojas" not in dados[usuario]:
            dados[usuario]["lojas"] = {}
        dados[usuario]["lojas"][loja] = {"email": email, "senha": senha}
        salvar_usuarios(dados)

def get_config_loja(usuario, loja):
    dados = carregar_usuarios()
    return dados.get(usuario, {}).get("lojas", {}).get(loja, {"email": "", "senha": ""})

# ── CONFIG ────────────────────────────────────────────────────
def carregar_config():
    if not ARQ_CONFIG.exists():
        return None
    spec = importlib.util.spec_from_file_location("config", ARQ_CONFIG)
    cfg  = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(cfg)
    return cfg

def salvar_valor_config(chave, valor):
    if not ARQ_CONFIG.exists():
        return
    linhas = []
    for linha in ARQ_CONFIG.read_text(encoding="utf-8").splitlines():
        if linha.strip().startswith(chave + " ="):
            linhas.append(f'{chave} = "{valor}"')
        else:
            linhas.append(linha)
    ARQ_CONFIG.write_text("\n".join(linhas), encoding="utf-8")

# ═════════════════════════════════════════════════════════════
# GERAR SCRIPT DE LOJA PARA FUNCIONÁRIO
# ═════════════════════════════════════════════════════════════
def gerar_script_usuario(usuario, nome_loja, script_nome):
    """Gera script personalizado com credenciais do usuário para a loja."""
    cfg_loja = get_config_loja(usuario, nome_loja)
    email = cfg_loja.get("email", "")
    senha = cfg_loja.get("senha", "")

    # Lê o script base
    script_base = PASTA / script_nome
    if not script_base.exists():
        return None

    txt = script_base.read_text(encoding="utf-8")

    # Substitui credenciais
    import re
    txt = re.sub(r'EMAIL\s*=\s*"[^"]*"', f'EMAIL = "{email}"', txt)
    txt = re.sub(r'SENHA\s*=\s*"[^"]*"', f'SENHA = "{senha}"', txt)

    # Salva script temporário
    pasta_tmp = PASTA / f"tmp_{usuario}"
    pasta_tmp.mkdir(exist_ok=True)
    script_tmp = pasta_tmp / script_nome
    script_tmp.write_text(txt, encoding="utf-8")

    # Copia config.py para pasta tmp
    if ARQ_CONFIG.exists():
        shutil.copy2(str(ARQ_CONFIG), str(pasta_tmp / "config.py"))

    return script_tmp

# ═════════════════════════════════════════════════════════════
# JANELA DE LOGIN
# ═════════════════════════════════════════════════════════════
usuario_logado = None

def abrir_login():
    global usuario_logado

    win = tk.Tk()
    win.title("Classic Sports - Login")
    win.configure(bg=COR_BG)
    win.geometry("420x520")
    win.resizable(False, False)

    # Centraliza
    win.update_idletasks()
    x = (win.winfo_screenwidth()  // 2) - 210
    y = (win.winfo_screenheight() // 2) - 260
    win.geometry(f"420x520+{x}+{y}")

    # Header
    tk.Frame(win, bg=COR_HEADER, height=80).pack(fill="x")
    hdr = win.children[list(win.children)[-1]]
    tk.Label(hdr, text="Classic ", font=("Arial", 20, "bold"),
             bg=COR_HEADER, fg="white").place(x=30, y=22)
    tk.Label(hdr, text="Sports", font=("Arial", 20, "bold"),
             bg=COR_HEADER, fg=COR_CYAN).place(x=108, y=22)
    tk.Label(hdr, text="Automação de Relatórios", font=("Arial", 9),
             bg=COR_HEADER, fg=COR_MUTED).place(x=30, y=52)
    tk.Frame(win, bg=COR_BLUE, height=2).pack(fill="x")

    # Corpo
    corpo = tk.Frame(win, bg=COR_BG)
    corpo.pack(fill="both", expand=True, padx=40, pady=30)

    tk.Label(corpo, text="Bem-vindo!", font=("Arial", 16, "bold"),
             bg=COR_BG, fg=COR_BRIGHT).pack(pady=(0,4))
    tk.Label(corpo, text="Entre com suas credenciais para continuar",
             font=("Arial", 9), bg=COR_BG, fg=COR_MUTED).pack(pady=(0,24))

    # Usuário
    tk.Label(corpo, text="USUÁRIO", font=("Arial", 9, "bold"),
             bg=COR_BG, fg=COR_MUTED).pack(anchor="w")
    var_user = tk.StringVar()
    entry_user = tk.Entry(corpo, textvariable=var_user, font=("Arial", 12),
                          bg="#001838", fg=COR_BRIGHT, insertbackground=COR_CYAN,
                          relief="flat", highlightbackground="#1A4A7A",
                          highlightthickness=1)
    entry_user.pack(fill="x", ipady=8, pady=(2,14))

    # Senha
    tk.Label(corpo, text="SENHA", font=("Arial", 9, "bold"),
             bg=COR_BG, fg=COR_MUTED).pack(anchor="w")
    var_pass = tk.StringVar()
    entry_pass = tk.Entry(corpo, textvariable=var_pass, font=("Arial", 12),
                          bg="#001838", fg=COR_BRIGHT, insertbackground=COR_CYAN,
                          relief="flat", highlightbackground="#1A4A7A",
                          highlightthickness=1, show="●")
    entry_pass.pack(fill="x", ipady=8, pady=(2,6))

    lbl_erro = tk.Label(corpo, text="", font=("Arial", 9),
                        bg=COR_BG, fg=COR_DANGER)
    lbl_erro.pack(pady=(0,16))

    def tentar_login(event=None):
        global usuario_logado
        u = var_user.get().strip()
        p = var_pass.get().strip()

        if not u or not p:
            lbl_erro.config(text="Preencha usuário e senha.")
            return

        # Primeiro acesso: cria usuário automaticamente
        if not usuario_existe(u):
            criar_usuario(u, "1234")

        if not verificar_senha(u, p):
            lbl_erro.config(text="Usuário ou senha incorretos.")
            return

        usuario_logado = u
        info = get_usuario(u)

        win.destroy()

        # Primeiro acesso: forçar troca de senha
        if info.get("primeiro_acesso", True):
            abrir_troca_senha_obrigatoria(u)
        else:
            abrir_painel(u)

    btn_login = tk.Button(corpo, text="ENTRAR",
                          font=("Arial", 12, "bold"), bg=COR_BLUE, fg="white",
                          relief="flat", cursor="hand2", pady=10,
                          command=tentar_login)
    btn_login.pack(fill="x")

    tk.Label(corpo, text="Primeiro acesso? Use a senha padrão: 1234",
             font=("Arial", 8), bg=COR_BG, fg=COR_MUTED).pack(pady=(12,0))

    entry_pass.bind("<Return>", tentar_login)
    entry_user.bind("<Return>", lambda e: entry_pass.focus())
    entry_user.focus()
    win.mainloop()


# ═════════════════════════════════════════════════════════════
# TROCA DE SENHA OBRIGATÓRIA (PRIMEIRO ACESSO)
# ═════════════════════════════════════════════════════════════
def abrir_troca_senha_obrigatoria(usuario):
    win = tk.Tk()
    win.title("Classic Sports - Definir Senha")
    win.configure(bg=COR_BG)
    win.geometry("420x480")
    win.resizable(False, False)
    win.update_idletasks()
    x = (win.winfo_screenwidth()  // 2) - 210
    y = (win.winfo_screenheight() // 2) - 240
    win.geometry(f"420x480+{x}+{y}")

    tk.Frame(win, bg=COR_HEADER, height=70).pack(fill="x")
    hdr = win.children[list(win.children)[-1]]
    tk.Label(hdr, text="Classic ", font=("Arial", 18, "bold"),
             bg=COR_HEADER, fg="white").place(x=30, y=18)
    tk.Label(hdr, text="Sports", font=("Arial", 18, "bold"),
             bg=COR_HEADER, fg=COR_CYAN).place(x=103, y=18)
    tk.Frame(win, bg=COR_BLUE, height=2).pack(fill="x")

    corpo = tk.Frame(win, bg=COR_BG)
    corpo.pack(fill="both", expand=True, padx=40, pady=24)

    tk.Label(corpo, text="Defina sua senha", font=("Arial", 15, "bold"),
             bg=COR_BG, fg=COR_BRIGHT).pack(pady=(0,4))
    tk.Label(corpo, text=f"Olá, {usuario}! Este é seu primeiro acesso.\nDefina uma senha pessoal para continuar.",
             font=("Arial", 9), bg=COR_BG, fg=COR_MUTED, justify="center").pack(pady=(0,20))

    tk.Label(corpo, text="NOVA SENHA", font=("Arial", 9, "bold"),
             bg=COR_BG, fg=COR_MUTED).pack(anchor="w")
    var_nova = tk.StringVar()
    tk.Entry(corpo, textvariable=var_nova, font=("Arial", 12),
             bg="#001838", fg=COR_BRIGHT, insertbackground=COR_CYAN,
             relief="flat", highlightbackground="#1A4A7A",
             highlightthickness=1, show="●").pack(fill="x", ipady=8, pady=(2,14))

    tk.Label(corpo, text="CONFIRMAR SENHA", font=("Arial", 9, "bold"),
             bg=COR_BG, fg=COR_MUTED).pack(anchor="w")
    var_conf = tk.StringVar()
    tk.Entry(corpo, textvariable=var_conf, font=("Arial", 12),
             bg="#001838", fg=COR_BRIGHT, insertbackground=COR_CYAN,
             relief="flat", highlightbackground="#1A4A7A",
             highlightthickness=1, show="●").pack(fill="x", ipady=8, pady=(2,6))

    lbl_erro = tk.Label(corpo, text="", font=("Arial", 9),
                        bg=COR_BG, fg=COR_DANGER)
    lbl_erro.pack(pady=(0,14))

    def confirmar():
        nova  = var_nova.get().strip()
        conf  = var_conf.get().strip()
        if len(nova) < 4:
            lbl_erro.config(text="A senha deve ter pelo menos 4 caracteres.")
            return
        if nova != conf:
            lbl_erro.config(text="As senhas não coincidem.")
            return
        trocar_senha(usuario, nova)
        win.destroy()
        abrir_painel(usuario)

    tk.Button(corpo, text="DEFINIR SENHA E ENTRAR",
              font=("Arial", 12, "bold"), bg=COR_BLUE, fg="white",
              relief="flat", cursor="hand2", pady=10,
              command=confirmar).pack(fill="x")

    win.mainloop()


# ═════════════════════════════════════════════════════════════
# PAINEL PRINCIPAL
# ═════════════════════════════════════════════════════════════
def abrir_painel(usuario):
    status_lojas = {nome: "idle" for nome, _ in LOJAS}
    widgets_loja = {}

    root = tk.Tk()
    root.title(f"Classic Sports - Automação de Relatórios [{usuario}]")
    root.configure(bg=COR_BG)
    root.geometry("1100x750")
    root.minsize(800, 550)

    # ── HEADER ──────────────────────────────────────────────
    header = tk.Frame(root, bg=COR_HEADER, height=66)
    header.pack(fill="x")
    header.pack_propagate(False)

    tk.Label(header, text="Classic ", font=("Arial", 18, "bold"),
             bg=COR_HEADER, fg="white").pack(side="left", padx=(20,0), pady=16)
    tk.Label(header, text="Sports", font=("Arial", 18, "bold"),
             bg=COR_HEADER, fg=COR_CYAN).pack(side="left")
    tk.Frame(header, bg=COR_MUTED, width=1).pack(
        side="left", fill="y", padx=16, pady=16)
    tk.Label(header, text="AUTOMAÇÃO DE RELATÓRIOS", font=("Arial", 9),
             bg=COR_HEADER, fg=COR_MUTED).pack(side="left")

    # Usuário no header
    tk.Label(header, text=f"👤 {usuario}", font=("Arial", 10, "bold"),
             bg=COR_HEADER, fg=COR_CYAN).pack(side="right", padx=(0,20))

    # Botão de atualização
    lbl_versao = tk.Label(header, text=f"v{VERSAO_ATUAL}",
                           font=("Arial", 8), bg=COR_HEADER, fg=COR_MUTED)
    lbl_versao.pack(side="right", padx=(0,4))

    def abrir_atualizador():
        if not ATUALIZADOR_OK:
            messagebox.showinfo("Atualização",
                "Módulo de atualização não encontrado.\n"
                "Certifique-se que o arquivo 'atualizador.py' está na pasta do programa.")
            return

        win_upd = tk.Toplevel(root)
        win_upd.title("Atualização do Sistema")
        win_upd.configure(bg=COR_BG)
        win_upd.geometry("480x320")
        win_upd.resizable(False, False)
        win_upd.grab_set()

        tk.Frame(win_upd, bg=COR_HEADER, height=55).pack(fill="x")
        hdr_u = win_upd.children[list(win_upd.children)[-1]]
        tk.Label(hdr_u, text="Classic ", font=("Arial", 14, "bold"),
                 bg=COR_HEADER, fg="white").place(x=16, y=14)
        tk.Label(hdr_u, text="Sports", font=("Arial", 14, "bold"),
                 bg=COR_HEADER, fg=COR_CYAN).place(x=84, y=14)
        tk.Label(hdr_u, text="— Atualização do Sistema",
                 font=("Arial", 9), bg=COR_HEADER, fg=COR_MUTED).place(x=162, y=18)
        tk.Frame(win_upd, bg=COR_BLUE, height=2).pack(fill="x")

        corpo_u = tk.Frame(win_upd, bg=COR_BG)
        corpo_u.pack(fill="both", expand=True, padx=28, pady=20)

        versao_local = atualizador.get_versao_local()
        tk.Label(corpo_u, text=f"Versão instalada: {versao_local}",
                 font=("Arial", 10, "bold"), bg=COR_BG, fg=COR_BRIGHT).pack(anchor="w")

        lbl_status_upd = tk.Label(corpo_u, text="Verificando atualizações...",
                                   font=("Arial", 9), bg=COR_BG, fg=COR_MUTED)
        lbl_status_upd.pack(anchor="w", pady=(4,12))

        progress_upd = ttk.Progressbar(corpo_u, length=420, mode="determinate")
        progress_upd.pack(fill="x", pady=(0,10))

        txt_notas = tk.Text(corpo_u, height=5, font=("Arial", 9),
                            bg="#001020", fg=COR_BRIGHT, relief="flat",
                            highlightbackground="#1A4A7A", highlightthickness=1,
                            state="disabled", wrap="word")
        txt_notas.pack(fill="x", pady=(0,12))

        btn_reiniciar = tk.Button(corpo_u, text="🔄  Reiniciar e Aplicar",
                                   font=("Arial", 11, "bold"), bg=COR_SUCCESS,
                                   fg="#001228", relief="flat", cursor="hand2",
                                   pady=8, state="disabled",
                                   command=lambda: atualizador.reiniciar_aplicacao())
        btn_reiniciar.pack(fill="x")

        def on_progresso(msg, pct):
            lbl_status_upd.config(text=msg, fg=COR_CYAN)
            progress_upd["value"] = pct
            win_upd.update()

        def on_fim(atualizado, nova_versao, notas):
            progress_upd["value"] = 100
            if atualizado:
                lbl_versao.config(text=f"v{nova_versao} ✓", fg=COR_SUCCESS)
                lbl_status_upd.config(
                    text=f"✓ Atualizado para v{nova_versao}! Reinicie para aplicar.",
                    fg=COR_SUCCESS)
                txt_notas.config(state="normal")
                txt_notas.delete(1.0, tk.END)
                txt_notas.insert(tk.END, f"Novidades da v{nova_versao}:\n{notas}")
                txt_notas.config(state="disabled")
                btn_reiniciar.config(state="normal")
            else:
                lbl_status_upd.config(text=notas or "✓ Sistema já está atualizado.",
                                       fg=COR_SUCCESS if not notas else COR_MUTED)
                progress_upd["value"] = 100

        atualizador.verificar_e_atualizar(
            callback_progresso=on_progresso,
            callback_fim=on_fim)

    btn_upd = tk.Button(header, text="🔄 Verificar Atualizações",
                         font=("Arial", 9, "bold"), bg=COR_BTN, fg=COR_BTN_TXT,
                         padx=10, pady=4, cursor="hand2", relief="flat",
                         command=abrir_atualizador)
    btn_upd.pack(side="right", padx=(0,8))

    # Verificação silenciosa ao abrir
    def verificar_silencioso():
        if not ATUALIZADOR_OK:
            return
        def on_fim_silencioso(atualizado, nova_versao, notas):
            if atualizado:
                lbl_versao.config(
                    text=f"v{nova_versao} disponível! →",
                    fg=COR_WARNING)
                btn_upd.config(bg=COR_WARNING, fg="#001228")
        atualizador.verificar_e_atualizar(
            callback_fim=on_fim_silencioso,
            silencioso=True)

    root.after(3000, verificar_silencioso)  # verifica 3s após abrir

    frame_nav = tk.Frame(header, bg=COR_HEADER)
    frame_nav.pack(side="right", padx=8)

    frames_aba = {}

    def estilo_btn(btn, ativo):
        btn.config(bg=COR_BLUE if ativo else COR_BTN,
                   fg="white" if ativo else COR_BTN_TXT)

    def trocar_aba(nome, btn_clicado):
        for b in btns_nav:
            estilo_btn(b, False)
        estilo_btn(btn_clicado, True)
        for f in frames_aba.values():
            f.pack_forget()
        frames_aba[nome].pack(fill="both", expand=True)

    btn_lojas  = tk.Button(frame_nav, text="LOJAS",
                           font=("Arial", 10, "bold"), padx=14, pady=6,
                           cursor="hand2", relief="flat")
    btn_config = tk.Button(frame_nav, text="CONFIGURAÇÕES",
                           font=("Arial", 10, "bold"), padx=14, pady=6,
                           cursor="hand2", relief="flat")
    btns_nav = [btn_lojas, btn_config]
    btn_lojas.config( command=lambda: trocar_aba("lojas",  btn_lojas))
    btn_config.config(command=lambda: trocar_aba("config", btn_config))
    btn_lojas.pack(side="left", padx=3)
    btn_config.pack(side="left", padx=3)
    estilo_btn(btn_lojas, True)
    estilo_btn(btn_config, False)

    tk.Frame(root, bg=COR_BLUE, height=2).pack(fill="x")
    area = tk.Frame(root, bg=COR_BG)
    area.pack(fill="both", expand=True)

    # ════════════════════════════════════════════════════════
    # ABA LOJAS
    # ════════════════════════════════════════════════════════
    frame_lojas = tk.Frame(area, bg=COR_BG)
    frames_aba["lojas"] = frame_lojas

    # Stats
    frame_stats = tk.Frame(frame_lojas, bg=COR_BG)
    frame_stats.pack(fill="x", padx=20, pady=(16,8))

    stat_vars = {k: tk.StringVar(value=v) for k, v in
                 [("total","21"),("done","0"),("running","0"),("idle","21")]}

    for i, (key, label, cor) in enumerate([
        ("total","Total de Lojas",COR_BRIGHT),("done","Concluídas",COR_SUCCESS),
        ("running","Em Execução",COR_WARNING),("idle","Aguardando",COR_CYAN)
    ]):
        frame_stats.columnconfigure(i, weight=1)
        card = tk.Frame(frame_stats, bg=COR_CARD,
                        highlightbackground="#1A4A7A", highlightthickness=1)
        card.grid(row=0, column=i, padx=4, sticky="ew")
        tk.Label(card, textvariable=stat_vars[key],
                 font=("Arial", 22, "bold"), bg=COR_CARD, fg=cor).pack(pady=(10,2))
        tk.Label(card, text=label, font=("Arial", 9),
                 bg=COR_CARD, fg=COR_MUTED).pack(pady=(0,10))

    def atualizar_stats():
        vals = list(status_lojas.values())
        stat_vars["done"].set(str(vals.count("done")))
        stat_vars["running"].set(str(vals.count("running")))
        stat_vars["idle"].set(str(vals.count("idle")))

    # Top bar
    frame_top = tk.Frame(frame_lojas, bg=COR_BG)
    frame_top.pack(fill="x", padx=20, pady=(4,8))
    tk.Label(frame_top, text="Minhas ", font=("Arial", 14, "bold"),
             bg=COR_BG, fg=COR_BRIGHT).pack(side="left")
    tk.Label(frame_top, text="Lojas", font=("Arial", 14, "bold"),
             bg=COR_BG, fg=COR_CYAN).pack(side="left")
    btn_todas = tk.Button(frame_top, text="▶▶ Rodar Todas",
                          font=("Arial", 10, "bold"), bg=COR_BTN, fg=COR_BTN_TXT,
                          padx=14, pady=5, cursor="hand2", relief="flat")
    btn_todas.pack(side="right")

    # Grid responsivo
    frame_scroll_outer = tk.Frame(frame_lojas, bg=COR_BG)
    frame_scroll_outer.pack(fill="both", expand=True, padx=20, pady=(0,10))
    canvas_lojas = tk.Canvas(frame_scroll_outer, bg=COR_BG, highlightthickness=0)
    sb_lojas = ttk.Scrollbar(frame_scroll_outer, orient="vertical",
                              command=canvas_lojas.yview)
    frame_grid = tk.Frame(canvas_lojas, bg=COR_BG)
    frame_grid.bind("<Configure>", lambda e: canvas_lojas.configure(
        scrollregion=canvas_lojas.bbox("all")))
    win_id = canvas_lojas.create_window((0,0), window=frame_grid, anchor="nw")
    canvas_lojas.configure(yscrollcommand=sb_lojas.set)
    canvas_lojas.pack(side="left", fill="both", expand=True)
    sb_lojas.pack(side="right", fill="y")
    canvas_lojas.bind_all("<MouseWheel>",
        lambda e: canvas_lojas.yview_scroll(int(-1*(e.delta/120)), "units"))

    CARD_MIN_W = 260

    def redistribuir_cards(larg=None):
        if larg is None:
            larg = canvas_lojas.winfo_width()
        if larg < 10:
            larg = 800
        ncols = max(1, larg // CARD_MIN_W)
        for i, (nome, _) in enumerate(LOJAS):
            w = widgets_loja.get(nome)
            if w:
                w["card"].grid(row=i//ncols, column=i%ncols,
                               padx=6, pady=6, sticky="nsew")
        for c in range(ncols):
            frame_grid.columnconfigure(c, weight=1)

    def on_canvas_resize(event):
        canvas_lojas.itemconfig(win_id, width=event.width)
        redistribuir_cards(event.width)
    canvas_lojas.bind("<Configure>", on_canvas_resize)

    # ── EXECUTAR LOJA ────────────────────────────────────────
    def rodar_loja(nome, script):
        cfg_loja = get_config_loja(usuario, nome)
        if not cfg_loja.get("email") or not cfg_loja.get("senha"):
            messagebox.showwarning("Credenciais faltando",
                f"Configure o email e senha do SportBay Hub para a loja '{nome}'\n"
                f"na aba Configurações → Credenciais SportBay Hub.")
            return
        if status_lojas[nome] == "running":
            return

        script_tmp = gerar_script_usuario(usuario, nome, script)
        if not script_tmp:
            messagebox.showerror("Erro", f"Script base não encontrado: {script}")
            return

        w = widgets_loja[nome]
        status_lojas[nome] = "running"
        w["btn"].config(text="Rodando...", bg=COR_WARNING,
                        fg="#001228", state="disabled")
        w["status"].config(text="Em execução...", fg=COR_WARNING)
        w["card"].config(highlightbackground=COR_WARNING)
        w["barra"].config(bg=COR_WARNING)
        w["log"].config(state="normal")
        w["log"].delete(1.0, tk.END)
        w["log"].insert(tk.END, f"Iniciando {nome}...\n")
        w["log"].config(state="disabled")
        w["log"].pack(fill="x", padx=6, pady=(0,6))
        atualizar_stats()

        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"

        def executar():
            try:
                proc = subprocess.Popen(
                    [sys.executable, "-u", str(script_tmp)],
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    cwd=str(script_tmp.parent), env=env,
                )
                for lb in proc.stdout:
                    linha = lb.decode("utf-8", errors="replace")
                    w["log"].config(state="normal")
                    w["log"].insert(tk.END, linha)
                    w["log"].see(tk.END)
                    w["log"].config(state="disabled")
                proc.wait()
                sucesso = proc.returncode == 0
            except Exception as e:
                sucesso = False
                w["log"].config(state="normal")
                w["log"].insert(tk.END, f"ERRO: {e}\n")
                w["log"].config(state="disabled")

            # Move resultado para pasta do usuário
            nome_arq = nome.lower().replace(" ","_")
            pasta_result_tmp  = script_tmp.parent / f"relatorios_{nome_arq}"
            pasta_result_user = PASTA / f"relatorios_{usuario}_{nome_arq}"
            if pasta_result_tmp.exists():
                if pasta_result_user.exists():
                    shutil.rmtree(str(pasta_result_user))
                shutil.move(str(pasta_result_tmp), str(pasta_result_user))

            if sucesso:
                status_lojas[nome] = "done"
                w["btn"].config(text="✓ Concluído — Rodar Novamente",
                                bg=COR_SUCCESS, fg="#001228", state="normal")
                w["status"].config(text="Concluído", fg=COR_SUCCESS)
                w["card"].config(highlightbackground=COR_SUCCESS)
                w["barra"].config(bg=COR_SUCCESS)
            else:
                status_lojas[nome] = "error"
                w["btn"].config(text="Erro — Rodar Novamente",
                                bg=COR_DANGER, fg="white", state="normal")
                w["status"].config(text="Erro", fg=COR_DANGER)
                w["card"].config(highlightbackground=COR_DANGER)
                w["barra"].config(bg=COR_DANGER)
            atualizar_stats()

        threading.Thread(target=executar, daemon=True).start()

    def abrir_pasta_loja(nome):
        nome_arq = nome.lower().replace(" ","_")
        pasta_rel = PASTA / f"relatorios_{usuario}_{nome_arq}"
        pasta_rel.mkdir(exist_ok=True)
        subprocess.Popen(f'explorer "{pasta_rel}"')

    # Monta cards
    for nome, script in LOJAS:
        card = tk.Frame(frame_grid, bg=COR_CARD,
                        highlightbackground="#1A4A7A", highlightthickness=1)
        barra = tk.Frame(card, bg=COR_BLUE, width=4)
        barra.pack(side="left", fill="y")
        inner = tk.Frame(card, bg=COR_CARD)
        inner.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        tk.Label(inner, text=nome, font=("Arial", 12, "bold"),
                 bg=COR_CARD, fg=COR_BRIGHT, anchor="w").pack(fill="x")
        status_lbl = tk.Label(inner, text="Aguardando",
                               font=("Arial", 9, "bold"),
                               bg=COR_CARD, fg=COR_MUTED, anchor="w")
        status_lbl.pack(fill="x", pady=(2,6))
        frame_btns = tk.Frame(inner, bg=COR_CARD)
        frame_btns.pack(fill="x")
        btn_run = tk.Button(frame_btns, text="▶ Rodar",
                            font=("Arial", 9, "bold"), bg=COR_BTN, fg=COR_BTN_TXT,
                            padx=10, pady=4, cursor="hand2", relief="flat",
                            command=lambda n=nome, s=script: rodar_loja(n, s))
        btn_run.pack(side="left", fill="x", expand=True, padx=(0,4))
        btn_pasta = tk.Button(frame_btns, text="📁",
                              font=("Arial", 11), bg=COR_BTN, fg=COR_BTN_TXT,
                              padx=8, pady=4, cursor="hand2", relief="flat",
                              command=lambda n=nome: abrir_pasta_loja(n))
        btn_pasta.pack(side="right")
        log_box = scrolledtext.ScrolledText(inner, height=3,
                                             font=("Courier New", 8),
                                             bg="#000D1A", fg="#5EE89A",
                                             state="disabled", relief="flat")
        widgets_loja[nome] = {
            "card": card, "barra": barra,
            "status": status_lbl, "btn": btn_run, "log": log_box,
        }

    redistribuir_cards(1100)

    def rodar_todas():
        if not messagebox.askyesno("Rodar Todas",
                "Rodar os 21 scripts em sequência?\nPode demorar bastante."):
            return
        def seq():
            import time
            for nome, script in LOJAS:
                if status_lojas[nome] != "running":
                    rodar_loja(nome, script)
                    while status_lojas[nome] == "running":
                        time.sleep(1)
        threading.Thread(target=seq, daemon=True).start()

    btn_todas.config(command=rodar_todas)

    # ════════════════════════════════════════════════════════
    # ABA CONFIGURAÇÕES
    # ════════════════════════════════════════════════════════
    frame_config = tk.Frame(area, bg=COR_BG)
    frames_aba["config"] = frame_config

    canvas_cfg = tk.Canvas(frame_config, bg=COR_BG, highlightthickness=0)
    sb_cfg = ttk.Scrollbar(frame_config, orient="vertical",
                            command=canvas_cfg.yview)
    frame_cfg_i = tk.Frame(canvas_cfg, bg=COR_BG)
    frame_cfg_i.bind("<Configure>", lambda e: canvas_cfg.configure(
        scrollregion=canvas_cfg.bbox("all")))
    win_cfg = canvas_cfg.create_window((0,0), window=frame_cfg_i, anchor="nw")
    canvas_cfg.configure(yscrollcommand=sb_cfg.set)
    canvas_cfg.pack(side="left", fill="both", expand=True)
    sb_cfg.pack(side="right", fill="y")
    canvas_cfg.bind("<Configure>",
        lambda e: canvas_cfg.itemconfig(win_cfg, width=e.width))

    def secao(titulo):
        outer = tk.Frame(frame_cfg_i, bg=COR_CARD,
                         highlightbackground="#1A4A7A", highlightthickness=1)
        outer.pack(fill="x", padx=20, pady=6)
        tk.Label(outer, text=titulo, font=("Arial", 10, "bold"),
                 bg=COR_CARD, fg=COR_CYAN).pack(anchor="w", padx=14, pady=(12,4))
        tk.Frame(outer, bg="#1A4A7A", height=1).pack(fill="x", padx=14)
        inner = tk.Frame(outer, bg=COR_CARD)
        inner.pack(fill="x", padx=14, pady=10)
        return inner

    # ── CREDENCIAIS SPORTBAY HUB POR LOJA ───────────────────
    sec_cred = secao("Credenciais SportBay Hub — Suas Lojas")
    tk.Label(sec_cred,
             text="Configure o email e senha do SportBay Hub para cada loja.",
             font=("Arial", 9), bg=COR_CARD, fg=COR_MUTED).pack(anchor="w", pady=(0,10))

    # Scroll interno para as lojas
    frame_cred_scroll = tk.Frame(sec_cred, bg=COR_CARD)
    frame_cred_scroll.pack(fill="x")

    vars_cred = {}
    for i, nome in enumerate(NOMES_LOJAS):
        cfg_loja = get_config_loja(usuario, nome)
        row_f = tk.Frame(frame_cred_scroll, bg=COR_CARD)
        row_f.pack(fill="x", pady=4)

        tk.Label(row_f, text=nome, font=("Arial", 10, "bold"),
                 bg=COR_CARD, fg=COR_BRIGHT, width=18, anchor="w").pack(side="left")

        var_email = tk.StringVar(value=cfg_loja.get("email",""))
        var_senha = tk.StringVar(value=cfg_loja.get("senha",""))
        vars_cred[nome] = (var_email, var_senha)

        tk.Label(row_f, text="Email:", font=("Arial", 9),
                 bg=COR_CARD, fg=COR_MUTED).pack(side="left", padx=(0,4))
        tk.Entry(row_f, textvariable=var_email, font=("Arial", 10),
                 bg="#001020", fg=COR_BRIGHT, insertbackground=COR_CYAN,
                 relief="flat", highlightbackground="#1A4A7A",
                 highlightthickness=1, width=28).pack(side="left", ipady=4, padx=(0,12))

        tk.Label(row_f, text="Senha:", font=("Arial", 9),
                 bg=COR_CARD, fg=COR_MUTED).pack(side="left", padx=(0,4))
        tk.Entry(row_f, textvariable=var_senha, font=("Arial", 10),
                 bg="#001020", fg=COR_BRIGHT, insertbackground=COR_CYAN,
                 relief="flat", highlightbackground="#1A4A7A",
                 highlightthickness=1, width=20, show="●").pack(side="left", ipady=4)

    def salvar_credenciais():
        for nome, (ve, vs) in vars_cred.items():
            salvar_config_loja(usuario, nome, ve.get().strip(), vs.get().strip())
        messagebox.showinfo("Salvo", "Credenciais salvas com sucesso!")

    tk.Button(sec_cred, text="💾  Salvar Credenciais",
              font=("Arial", 11, "bold"), bg=COR_BLUE, fg="white",
              relief="flat", cursor="hand2", pady=8,
              command=salvar_credenciais).pack(anchor="w", pady=(12,0))

    # ── TABELAS DE REFERÊNCIA ────────────────────────────────
    cfg_atual = carregar_config()
    sec_tab = secao("Tabelas de Referência")

    TABELAS_INFO = [
        ("TABELA_PRECOS",        "Tabela de Preços Geral",
         getattr(cfg_atual, "TABELA_PRECOS", "tabela_precos.xlsx")),
        ("TABELA_PRECO_SKU_KITS","Preço SKU dos Kits",
         getattr(cfg_atual, "TABELA_PRECO_SKU_KITS", "Preco SKU dos KITS - 08-04-26.xlsx")),
        ("TABELA_MEUS_KITS",     "Meus Kits",
         getattr(cfg_atual, "TABELA_MEUS_KITS", "meus_kits.xlsx")),
    ]

    for col_idx, (chave, label, valor_atual) in enumerate(TABELAS_INFO):
        sec_tab.columnconfigure(col_idx, weight=1)
        f = tk.Frame(sec_tab, bg=COR_CARD)
        f.grid(row=0, column=col_idx, padx=6, pady=4, sticky="ew")
        tk.Label(f, text=label.upper(), font=("Arial", 8, "bold"),
                 bg=COR_CARD, fg=COR_MUTED).pack(anchor="w")
        var = tk.StringVar(value=valor_atual)
        frow = tk.Frame(f, bg=COR_CARD)
        frow.pack(fill="x")
        tk.Entry(frow, textvariable=var, font=("Arial", 10),
                 bg="#001020", fg=COR_BRIGHT, insertbackground=COR_CYAN,
                 relief="flat", highlightbackground="#1A4A7A",
                 highlightthickness=1, state="readonly").pack(
                     side="left", fill="x", expand=True, ipady=5)

        def fazer_upload(c=chave, v=var):
            caminho = filedialog.askopenfilename(
                title=f"Selecione o arquivo para {c}",
                filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
            if not caminho:
                return
            src = Path(caminho)
            dst = PASTA / src.name
            try:
                shutil.copy2(str(src), str(dst))
                v.set(src.name)
                salvar_valor_config(c, src.name)
                messagebox.showinfo("Sucesso",
                    f"Arquivo '{src.name}' copiado e config atualizado!")
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível copiar:\n{e}")

        tk.Button(frow, text="📂 Enviar", font=("Arial", 9, "bold"),
                  bg=COR_BLUE, fg="white", padx=8, pady=5,
                  cursor="hand2", relief="flat",
                  command=fazer_upload).pack(side="right", padx=(4,0))
        existe = (PASTA / valor_atual).exists()
        tk.Label(f,
                 text="✓ Encontrado" if existe else "✗ Não encontrado",
                 font=("Arial", 8), bg=COR_CARD,
                 fg=COR_SUCCESS if existe else COR_DANGER).pack(anchor="w", pady=(2,0))

    # ── PARÂMETROS FINANCEIROS ───────────────────────────────
    sec_fin = secao("Parâmetros Financeiros")
    for col_idx, (label, valor, hint) in enumerate([
        ("Margem Mínima Padrão (%)",      "16", "Quando não encontrada na tabela"),
        ("Margem Mínima Alternativa (%)", "12", "Premium / secundária"),
        ("Alíquota de Impostos (%)",      "17", "Aplicada no preço mínimo"),
    ]):
        sec_fin.columnconfigure(col_idx, weight=1)
        f = tk.Frame(sec_fin, bg=COR_CARD)
        f.grid(row=0, column=col_idx, padx=6, pady=4, sticky="ew")
        tk.Label(f, text=label.upper(), font=("Arial", 8, "bold"),
                 bg=COR_CARD, fg=COR_MUTED).pack(anchor="w")
        var = tk.StringVar(value=valor)
        tk.Entry(f, textvariable=var, font=("Arial", 13),
                 bg="#001020", fg=COR_BRIGHT, insertbackground=COR_CYAN,
                 relief="flat", highlightbackground="#1A4A7A",
                 highlightthickness=1, justify="center").pack(fill="x", ipady=6)
        tk.Label(f, text=hint, font=("Arial", 8),
                 bg=COR_CARD, fg=COR_MUTED).pack(anchor="w", pady=(2,0))

    # ── TROCAR SENHA ─────────────────────────────────────────
    sec_senha = secao("Alterar Minha Senha")
    f_senha = tk.Frame(sec_senha, bg=COR_CARD)
    f_senha.pack(fill="x")

    for col_idx, (label, show) in enumerate([
        ("Senha Atual", "●"), ("Nova Senha", "●"), ("Confirmar Nova Senha", "●")
    ]):
        f_senha.columnconfigure(col_idx, weight=1)
        f = tk.Frame(f_senha, bg=COR_CARD)
        f.grid(row=0, column=col_idx, padx=6, sticky="ew")
        tk.Label(f, text=label.upper(), font=("Arial", 8, "bold"),
                 bg=COR_CARD, fg=COR_MUTED).pack(anchor="w")
        var = tk.StringVar()
        tk.Entry(f, textvariable=var, font=("Arial", 11),
                 bg="#001020", fg=COR_BRIGHT, insertbackground=COR_CYAN,
                 relief="flat", highlightbackground="#1A4A7A",
                 highlightthickness=1, show=show).pack(fill="x", ipady=6)
        if col_idx == 0: var_atual = var
        elif col_idx == 1: var_nova = var
        else: var_conf = var

    def alterar_senha():
        atual = var_atual.get().strip()
        nova  = var_nova.get().strip()
        conf  = var_conf.get().strip()
        if not verificar_senha(usuario, atual):
            messagebox.showerror("Erro", "Senha atual incorreta.")
            return
        if len(nova) < 4:
            messagebox.showerror("Erro", "A nova senha deve ter pelo menos 4 caracteres.")
            return
        if nova != conf:
            messagebox.showerror("Erro", "As novas senhas não coincidem.")
            return
        trocar_senha(usuario, nova)
        var_atual.set(""); var_nova.set(""); var_conf.set("")
        messagebox.showinfo("Sucesso", "Senha alterada com sucesso!")

    tk.Button(sec_senha, text="🔒  Alterar Senha",
              font=("Arial", 11, "bold"), bg=COR_BLUE, fg="white",
              relief="flat", cursor="hand2", pady=8,
              command=alterar_senha).pack(anchor="w", pady=(12,0))

    # Fórmula
    f_form = tk.Frame(frame_cfg_i, bg="#001830",
                      highlightbackground="#003060", highlightthickness=1)
    f_form.pack(fill="x", padx=20, pady=6)
    tk.Label(f_form, text="FÓRMULA DO PREÇO MÍNIMO",
             font=("Arial", 8, "bold"), bg="#001830", fg=COR_MUTED).pack(
                 anchor="w", padx=14, pady=(10,2))
    tk.Label(f_form,
             text="Preço Mínimo = (Custo + Frete) / (1 - Margem% - Imposto% - Taxa_ML%)",
             font=("Courier New", 11), bg="#001830", fg=COR_BRIGHT).pack(
                 anchor="w", padx=14)
    tk.Label(f_form,
             text="Frete = Tarifa Frete Grátis (R$)   |   Taxa_ML = Porcentagem cobrada do ML (%)",
             font=("Arial", 9), bg="#001830", fg=COR_MUTED).pack(
                 anchor="w", padx=14, pady=(0,10))

    frames_aba["lojas"].pack(fill="both", expand=True)
    root.mainloop()


# ═════════════════════════════════════════════════════════════
# ENTRY POINT
# ═════════════════════════════════════════════════════════════
if __name__ == "__main__":
    abrir_login()
