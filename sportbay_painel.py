# -*- coding: utf-8 -*-
import sys, os, shutil, subprocess, threading, hashlib, json, importlib.util
import base64, socket, struct
from pathlib import Path
from tkinter import filedialog
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

PASTA       = Path(__file__).resolve().parent
PASTA_DADOS = PASTA / "dados"
PASTA_DADOS.mkdir(parents=True, exist_ok=True)
ARQ_CRED = PASTA_DADOS / "credenciais.enc"

# ── CRIPTOGRAFIA DE CREDENCIAIS SPORTBAY ─────────────────────
def _machine_key():
    return hashlib.sha256(socket.gethostname().upper().strip().encode()).digest()

def _cred_encrypt(texto):
    key = _machine_key()
    iv  = os.urandom(16)
    pt  = texto.encode("utf-8")
    ks, i = b"", 0
    while len(ks) < len(pt):
        ks += hashlib.sha256(key + iv + struct.pack(">I", i)).digest()
        i  += 1
    ct   = bytes(a ^ b for a, b in zip(pt, ks))
    hmac = hashlib.sha256(key + iv + ct).digest()[:16]
    return base64.b64encode(iv + hmac + ct).decode()

def _cred_decrypt(token):
    key = _machine_key()
    raw = base64.b64decode(token.encode())
    iv, hmac_check, ct = raw[:16], raw[16:32], raw[32:]
    if hmac_check != hashlib.sha256(key + iv + ct).digest()[:16]:
        raise ValueError("HMAC invalido")
    ks, i = b"", 0
    while len(ks) < len(ct):
        ks += hashlib.sha256(key + iv + struct.pack(">I", i)).digest()
        i  += 1
    return bytes(a ^ b for a, b in zip(ct, ks)).decode("utf-8")

def salvar_credenciais_enc(creds):
    ARQ_CRED.write_text(_cred_encrypt(json.dumps(creds, ensure_ascii=False)), encoding="utf-8")

def carregar_credenciais_enc():
    if not ARQ_CRED.exists():
        return {}
    try:
        return json.loads(_cred_decrypt(ARQ_CRED.read_text(encoding="utf-8")))
    except Exception:
        return {}

def cred_enc_existe():
    return ARQ_CRED.exists() and ARQ_CRED.stat().st_size > 0

# ── VERIFICA E INSTALA DEPENDÊNCIAS AUTOMATICAMENTE ──────────
def _verificar_deps():
    """Instala pandas, openpyxl e playwright se não estiverem disponíveis."""
    deps_faltando = []
    for dep in ["pandas", "openpyxl", "playwright"]:
        try:
            __import__(dep)
        except ImportError:
            deps_faltando.append(dep)
    if deps_faltando:
        import subprocess, sys
        try:
            subprocess.run(
                [sys.executable, "-m", "pip", "install", "--quiet"] + deps_faltando,
                check=True, capture_output=True
            )
            # Instala Chromium se playwright foi instalado agora
            if "playwright" in deps_faltando:
                subprocess.run(
                    [sys.executable, "-m", "playwright", "install", "chromium"],
                    capture_output=True
                )
        except Exception:
            pass  # Continua mesmo se falhar — erro aparecerá na execução

_verificar_deps()
ARQ_USUARIOS = PASTA_DADOS / "usuarios.json"
ARQ_CONFIG   = PASTA / "config.py"
VERSAO_ATUAL = "1.0.0"

BG      = "#001228"
HEADER  = "#00142E"
BLUE    = "#0070CC"
CYAN    = "#00AAFF"
CARD    = "#002850"
BRIGHT  = "#E0ECFA"
MUTED   = "#7A9DBE"
SUCCESS = "#2ECC71"
WARNING = "#F39C12"
DANGER  = "#E74C3C"
BTN     = "#0D2D4E"
BTN_TXT = "#7ECFFF"

LOJAS = [
    ("Classic Barracao","sportbay_classic_barracao.py"),
    ("CS","sportbay_cs.py"),("Juliana","sportbay_juliana.py"),
    ("Marcus","sportbay_marcus.py"),("11","sportbay_11.py"),
    ("12","sportbay_12.py"),("Imports","sportbay_imports.py"),
    ("708","sportbay_708.py"),("AdrenalineX","sportbay_adrenalinex.py"),
    ("AdventureX","sportbay_adventurex.py"),("AM15","sportbay_am15.py"),
    ("AM20","sportbay_am20.py"),("Planet","sportbay_planet.py"),
    ("RAS","sportbay_ras.py"),("DL","sportbay_dl.py"),
    ("DM","sportbay_dm.py"),("BL","sportbay_bl.py"),
    ("GJ","sportbay_gj.py"),("FF","sportbay_ff.py"),
    ("JA","sportbay_ja.py"),("SS","sportbay_ss.py"),
]

# ── USUÁRIOS ──────────────────────────────────────────────────
def _hash(s): return hashlib.sha256(s.encode()).hexdigest()
def _load(): 
    if not ARQ_USUARIOS.exists(): return {}
    try: return json.loads(ARQ_USUARIOS.read_text(encoding="utf-8"))
    except: return {}
def _save(d): ARQ_USUARIOS.write_text(json.dumps(d,indent=2,ensure_ascii=False),encoding="utf-8")
def criar_usuario(u):
    d=_load(); d[u]={"senha_hash":_hash("1234"),"primeiro_acesso":True,"lojas":{}}; _save(d)
def verificar_senha(u,s): d=_load(); return u in d and d[u]["senha_hash"]==_hash(s)
def trocar_senha_user(u,s): d=_load(); d[u]["senha_hash"]=_hash(s); d[u]["primeiro_acesso"]=False; _save(d)
def get_user(u): return _load().get(u,{})
def salvar_cred_loja(u,loja,email,senha):
    # Salva no usuarios.json (compatibilidade)
    d=_load()
    if u in d:
        if "lojas" not in d[u]: d[u]["lojas"]={}
        d[u]["lojas"][loja]={"email":email,"senha":senha}; _save(d)
    # Salva também no credenciais.enc (seguro, lido pelos scripts)
    creds = carregar_credenciais_enc()
    creds[loja] = {"email": email, "senha": senha}
    salvar_credenciais_enc(creds)

def get_cred_loja(u,loja):
    # Tenta primeiro do .enc, fallback para usuarios.json
    creds = carregar_credenciais_enc()
    if loja in creds:
        return creds[loja]
    return _load().get(u,{}).get("lojas",{}).get(loja,{"email":"","senha":""})

# ── CONFIG ────────────────────────────────────────────────────
def _criar_config_padrao():
    """Cria config.py com valores padrão se não existir."""
    ARQ_CONFIG.write_text(
        '# Configuracao Classic Sports\n'
        'TABELA_PRECOS = ""\n'
        'TABELA_PRECO_SKU_KITS = ""\n'
        'TABELA_MEUS_KITS = ""\n'
        'CHROME_PROFILE_DIR = r"C:\\Users\\Usuario\\AppData\\Local\\Google\\Chrome\\User Data"\n',
        encoding="utf-8"
    )

def load_cfg():
    if not ARQ_CONFIG.exists():
        _criar_config_padrao()
    try:
        spec=importlib.util.spec_from_file_location("config",ARQ_CONFIG)
        cfg=importlib.util.module_from_spec(spec); spec.loader.exec_module(cfg); return cfg
    except Exception:
        return None
def save_cfg_val(chave,valor):
    if not ARQ_CONFIG.exists():
        _criar_config_padrao()
    lines=[]
    encontrou = False
    for l in ARQ_CONFIG.read_text(encoding="utf-8").splitlines():
        if l.strip().startswith(chave+" ="):
            lines.append(f'{chave} = "{valor}"')
            encontrou = True
        else:
            lines.append(l)
    if not encontrou:
        lines.append(f'{chave} = "{valor}"')
    ARQ_CONFIG.write_text("\n".join(lines),encoding="utf-8")

def set_icon(win):
    try:
        ico=PASTA/"icon.ico"
        if ico.exists(): win.iconbitmap(str(ico))
    except: pass

def gerar_script(usuario,nome_loja,script_nome):
    sb=PASTA/script_nome
    if not sb.exists():
        encontrados = list(PASTA.rglob(script_nome))
        if encontrados: sb = encontrados[0]
    if not sb.exists(): return None
    # Copia o script para a pasta tmp (sem injetar email/senha — scripts leem do .enc)
    pt=PASTA/f"tmp_{usuario}"; pt.mkdir(exist_ok=True)
    st=pt/script_nome
    shutil.copy2(str(sb), str(st))
    # Garante config.py na pasta tmp
    if not ARQ_CONFIG.exists(): _criar_config_padrao()
    if ARQ_CONFIG.exists(): shutil.copy2(str(ARQ_CONFIG),str(pt/'config.py'))
    # Garante credenciais.enc acessível pela pasta tmp (link via pasta pai)
    # Os scripts já buscam em parent/dados/credenciais.enc automaticamente
    # Copia os arquivos de tabela para a pasta tmp
    cfg = load_cfg()
    if cfg:
        for attr in ['TABELA_PRECOS','TABELA_PRECO_SKU_KITS','TABELA_MEUS_KITS']:
            nome_arq = getattr(cfg, attr, '')
            if nome_arq:
                src = PASTA / nome_arq
                dst = pt / nome_arq
                if src.exists() and not dst.exists():
                    try:
                        shutil.copy2(str(src), str(dst))
                    except Exception:
                        pass
    return st
# ══════════════════════════════════════════════════════════════
# APLICAÇÃO — janela única, troca de telas via pack/pack_forget
# ══════════════════════════════════════════════════════════════
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Classic Sports - Automação de Relatórios")
        self.root.configure(bg=BG)
        self.root.geometry("460x520")
        self.root.resizable(False,False)
        set_icon(self.root)
        self.root.update_idletasks()
        sw,sh=self.root.winfo_screenwidth(),self.root.winfo_screenheight()
        self.root.geometry(f"460x520+{(sw-460)//2}+{(sh-520)//2}")
        self.usuario = None
        self._tela_login()
        self.root.mainloop()

    def _limpar(self):
        for w in self.root.winfo_children():
            w.destroy()

    # ── TELA LOGIN ─────────────────────────────────────────────
    def _tela_login(self):
        self._limpar()
        self.root.title("Classic Sports - Login")
        self.root.geometry("460x520")
        self.root.resizable(False,False)

        hdr=tk.Frame(self.root,bg=HEADER,height=70); hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr,text="Classic ",font=("Arial",18,"bold"),bg=HEADER,fg="white").place(x=20,y=18)
        tk.Label(hdr,text="Sports",  font=("Arial",18,"bold"),bg=HEADER,fg=CYAN).place(x=103,y=18)
        tk.Label(hdr,text="Automação de Relatórios",font=("Arial",9),bg=HEADER,fg=MUTED).place(x=20,y=48)
        tk.Frame(self.root,bg=BLUE,height=2).pack(fill="x")

        body=tk.Frame(self.root,bg=BG); body.pack(fill="both",expand=True,padx=40,pady=28)
        tk.Label(body,text="Bem-vindo!",font=("Arial",16,"bold"),bg=BG,fg=BRIGHT).pack(pady=(0,4))
        tk.Label(body,text="Entre com suas credenciais para continuar",font=("Arial",9),bg=BG,fg=MUTED).pack(pady=(0,22))

        var_u=tk.StringVar(); var_p=tk.StringVar()
        lbl_err=tk.Label(body,text="",font=("Arial",9),bg=BG,fg=DANGER)

        tk.Label(body,text="USUÁRIO",font=("Arial",9,"bold"),bg=BG,fg=MUTED).pack(anchor="w")
        eu=tk.Entry(body,textvariable=var_u,font=("Arial",12),bg="#001838",fg=BRIGHT,
                    insertbackground=CYAN,relief="flat",highlightbackground="#1A4A7A",highlightthickness=1)
        eu.pack(fill="x",ipady=8,pady=(2,14))
        tk.Label(body,text="SENHA",font=("Arial",9,"bold"),bg=BG,fg=MUTED).pack(anchor="w")
        ep=tk.Entry(body,textvariable=var_p,font=("Arial",12),bg="#001838",fg=BRIGHT,
                    insertbackground=CYAN,relief="flat",highlightbackground="#1A4A7A",highlightthickness=1,show="●")
        ep.pack(fill="x",ipady=8,pady=(2,4))
        lbl_err.pack(pady=(2,14))

        def entrar(e=None):
            u=var_u.get().strip(); p=var_p.get().strip()
            if not u or not p: lbl_err.config(text="Preencha usuário e senha."); return
            if not get_user(u): criar_usuario(u)
            if not verificar_senha(u,p): lbl_err.config(text="Usuário ou senha incorretos."); return
            self.usuario=u
            if get_user(u).get("primeiro_acesso",True):
                self._tela_troca_senha()
            elif not cred_enc_existe():
                self._tela_config_creds()
            else:
                self._abrir_painel()

        tk.Button(body,text="ENTRAR",font=("Arial",12,"bold"),bg=BLUE,fg="white",
                  relief="flat",cursor="hand2",pady=10,command=entrar).pack(fill="x")
        tk.Label(body,text="Primeiro acesso? Senha padrão: 1234",font=("Arial",8),bg=BG,fg=MUTED).pack(pady=(12,0))
        ep.bind("<Return>",entrar); eu.bind("<Return>",lambda e:ep.focus()); eu.focus()

    # ── TELA TROCA SENHA ───────────────────────────────────────
    def _tela_troca_senha(self):
        self._limpar()
        self.root.title("Classic Sports - Definir Senha")
        self.root.geometry("460x460")

        hdr=tk.Frame(self.root,bg=HEADER,height=70); hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr,text="Classic ",font=("Arial",18,"bold"),bg=HEADER,fg="white").place(x=20,y=18)
        tk.Label(hdr,text="Sports",  font=("Arial",18,"bold"),bg=HEADER,fg=CYAN).place(x=103,y=18)
        tk.Frame(self.root,bg=BLUE,height=2).pack(fill="x")

        body=tk.Frame(self.root,bg=BG); body.pack(fill="both",expand=True,padx=40,pady=24)
        tk.Label(body,text="Defina sua senha",font=("Arial",15,"bold"),bg=BG,fg=BRIGHT).pack(pady=(0,4))
        tk.Label(body,text=f"Olá, {self.usuario}! Crie sua senha pessoal para continuar.",
                 font=("Arial",9),bg=BG,fg=MUTED,wraplength=360).pack(pady=(0,18))

        var_n=tk.StringVar(); var_c=tk.StringVar()
        lbl_err=tk.Label(body,text="",font=("Arial",9),bg=BG,fg=DANGER)

        tk.Label(body,text="NOVA SENHA",font=("Arial",9,"bold"),bg=BG,fg=MUTED).pack(anchor="w")
        tk.Entry(body,textvariable=var_n,font=("Arial",12),bg="#001838",fg=BRIGHT,
                 insertbackground=CYAN,relief="flat",highlightbackground="#1A4A7A",
                 highlightthickness=1,show="●").pack(fill="x",ipady=8,pady=(2,12))
        tk.Label(body,text="CONFIRMAR SENHA",font=("Arial",9,"bold"),bg=BG,fg=MUTED).pack(anchor="w")
        tk.Entry(body,textvariable=var_c,font=("Arial",12),bg="#001838",fg=BRIGHT,
                 insertbackground=CYAN,relief="flat",highlightbackground="#1A4A7A",
                 highlightthickness=1,show="●").pack(fill="x",ipady=8,pady=(2,4))
        lbl_err.pack(pady=(2,14))

        def confirmar():
            n=var_n.get().strip(); c=var_c.get().strip()
            if len(n)<4: lbl_err.config(text="Mínimo 4 caracteres."); return
            if n!=c: lbl_err.config(text="As senhas não coincidem."); return
            trocar_senha_user(self.usuario,n)
            if not cred_enc_existe():
                self._tela_config_creds()
            else:
                self._abrir_painel()

        tk.Button(body,text="DEFINIR SENHA E ENTRAR",font=("Arial",12,"bold"),
                  bg=BLUE,fg="white",relief="flat",cursor="hand2",pady=10,
                  command=confirmar).pack(fill="x")


    # ── TELA CONFIGURAÇÃO DE CREDENCIAIS (pós-atualização) ────
    def _tela_config_creds(self):
        """Exibida quando credenciais.enc não existe — pede email/senha do SportBay."""
        self._limpar()
        self.root.title("Classic Sports - Configurar Credenciais")
        self.root.geometry("520x600")
        self.root.resizable(False, False)
        self.root.update_idletasks()
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry("520x600+" + str((sw-520)//2) + "+" + str((sh-600)//2))

        hdr = tk.Frame(self.root, bg=HEADER, height=70); hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="Classic ", font=("Arial",18,"bold"), bg=HEADER, fg="white").place(x=20, y=18)
        tk.Label(hdr, text="Sports",   font=("Arial",18,"bold"), bg=HEADER, fg=CYAN).place(x=103, y=18)
        tk.Label(hdr, text="Configuração de Credenciais", font=("Arial",9), bg=HEADER, fg=MUTED).place(x=20, y=48)
        tk.Frame(self.root, bg=BLUE, height=2).pack(fill="x")

        body = tk.Frame(self.root, bg=BG); body.pack(fill="both", expand=True)

        tk.Label(body, text="Configure suas credenciais do SportBay Hub",
                 font=("Arial",13,"bold"), bg=BG, fg=BRIGHT).pack(pady=(18,4), padx=20, anchor="w")
        tk.Label(body,
                 text="O sistema foi atualizado e precisa reconfigurar as credenciais nesta maquina. Configure o email e senha de cada loja abaixo.",
                 font=("Arial",9), bg=BG, fg=MUTED, justify="left").pack(padx=20, anchor="w", pady=(0,12))

        # Frame scrollável
        outer = tk.Frame(body, bg=BG); outer.pack(fill="both", expand=True, padx=16)
        sb2 = ttk.Scrollbar(outer, orient="vertical"); sb2.pack(side="right", fill="y")
        cv2 = tk.Canvas(outer, bg=BG, highlightthickness=0, yscrollcommand=sb2.set)
        cv2.pack(side="left", fill="both", expand=True)
        sb2.config(command=cv2.yview)
        frm2 = tk.Frame(cv2, bg=BG)
        win2 = cv2.create_window((0,0), window=frm2, anchor="nw")
        frm2.bind("<Configure>", lambda e: cv2.configure(scrollregion=cv2.bbox("all")))
        cv2.bind("<Configure>", lambda e: cv2.itemconfig(win2, width=e.width))
        cv2.bind_all("<MouseWheel>", lambda e: cv2.yview_scroll(int(-1*(e.delta/120)), "units"))

        # Carrega credenciais existentes (do usuarios.json, como fallback)
        d = _load().get(self.usuario, {}).get("lojas", {})
        self._vars_config_creds = {}
        for nome, _ in LOJAS:
            c_atual = d.get(nome, {"email": "", "senha": ""})
            row = tk.Frame(frm2, bg=CARD, highlightbackground="#1A4A7A", highlightthickness=1)
            row.pack(fill="x", pady=3)
            tk.Label(row, text=nome, font=("Arial",9,"bold"), bg=CARD, fg=BRIGHT,
                     width=18, anchor="w").pack(side="left", padx=(10,4), pady=8)
            ve = tk.StringVar(value=c_atual.get("email", ""))
            vs = tk.StringVar(value=c_atual.get("senha", ""))
            tk.Entry(row, textvariable=ve, font=("Arial",9), bg="#001020", fg=BRIGHT,
                     insertbackground=CYAN, relief="flat",
                     highlightbackground="#1A4A7A", highlightthickness=1,
                     width=24).pack(side="left", ipady=4, padx=(0,6))
            tk.Entry(row, textvariable=vs, font=("Arial",9), bg="#001020", fg=BRIGHT,
                     insertbackground=CYAN, relief="flat",
                     highlightbackground="#1A4A7A", highlightthickness=1,
                     width=18, show="●").pack(side="left", ipady=4)
            self._vars_config_creds[nome] = (ve, vs)

        lbl_err2 = tk.Label(body, text="", font=("Arial",9), bg=BG, fg=DANGER)
        lbl_err2.pack(pady=(6,0))

        def salvar_e_entrar():
            # Valida que pelo menos uma loja tem credenciais
            creds = {}
            for nome, (ve, vs) in self._vars_config_creds.items():
                email = ve.get().strip()
                senha = vs.get().strip()
                if email or senha:
                    creds[nome] = {"email": email, "senha": senha}
                    # Salva também em usuarios.json
                    salvar_cred_loja(self.usuario, nome, email, senha)
            if not creds:
                lbl_err2.config(text="Configure pelo menos uma loja para continuar.")
                return
            # Gera o credenciais.enc com todas as credenciais
            salvar_credenciais_enc(creds)
            self._abrir_painel()

        tk.Button(body, text="Salvar e Entrar",
                  font=("Arial",12,"bold"), bg=BLUE, fg="white",
                  relief="flat", cursor="hand2", pady=10,
                  command=salvar_e_entrar).pack(fill="x", padx=16, pady=(8,16))

    # ── PAINEL PRINCIPAL ───────────────────────────────────────
    def _abrir_painel(self):
        self._limpar()
        self.root.title(f"Classic Sports - Automação de Relatórios [{self.usuario}]")
        self.root.geometry("1150x750")
        self.root.resizable(True,True)
        self.root.minsize(900,600)
        # Centraliza
        self.root.update_idletasks()
        sw,sh=self.root.winfo_screenwidth(),self.root.winfo_screenheight()
        self.root.geometry(f"1150x750+{(sw-1150)//2}+{(sh-750)//2}")

        self.status_lojas={n:"idle" for n,_ in LOJAS}
        self.widgets={}

        # ── HEADER ────────────────────────────────────────────
        hdr=tk.Frame(self.root,bg=HEADER,height=56); hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr,text="Classic ",font=("Arial",16,"bold"),bg=HEADER,fg="white").pack(side="left",padx=(16,0))
        tk.Label(hdr,text="Sports",  font=("Arial",16,"bold"),bg=HEADER,fg=CYAN).pack(side="left")
        tk.Frame(hdr,bg=MUTED,width=1).pack(side="left",fill="y",padx=12,pady=10)
        tk.Label(hdr,text="AUTOMAÇÃO DE RELATÓRIOS",font=("Arial",8),bg=HEADER,fg=MUTED).pack(side="left")
        tk.Label(hdr,text=f"👤 {self.usuario}",font=("Arial",10,"bold"),bg=HEADER,fg=CYAN).pack(side="right",padx=(0,16))
        tk.Label(hdr,text=f"v{VERSAO_ATUAL}",font=("Arial",8),bg=HEADER,fg=MUTED).pack(side="right",padx=(0,6))

        self.frm_nav=tk.Frame(hdr,bg=HEADER); self.frm_nav.pack(side="right",padx=8)
        tk.Frame(self.root,bg=BLUE,height=2).pack(fill="x")

        # ── CONTAINER ABAS ────────────────────────────────────
        self.container=tk.Frame(self.root,bg=BG); self.container.pack(fill="both",expand=True)

        # Cria abas
        self._criar_aba_lojas()
        self._criar_aba_senhas()
        self._criar_aba_params()

        # Botões nav
        self.btn_lojas = tk.Button(self.frm_nav, text="LOJAS", font=("Arial",10,"bold"),
                                   padx=14, pady=5, relief="flat", cursor="hand2",
                                   command=lambda: self._mostrar("lojas"))
        self.btn_senhas = tk.Button(self.frm_nav, text="SENHAS", font=("Arial",10,"bold"),
                                    padx=14, pady=5, relief="flat", cursor="hand2",
                                    command=lambda: self._mostrar("senhas"))
        self.btn_params = tk.Button(self.frm_nav, text="PARÂMETROS", font=("Arial",10,"bold"),
                                    padx=14, pady=5, relief="flat", cursor="hand2",
                                    command=lambda: self._mostrar("params"))
        tk.Button(self.frm_nav, text="Atualizacoes", font=("Arial",9,"bold"),
                  bg=BTN, fg=BTN_TXT, padx=10, pady=5, relief="flat",
                  cursor="hand2", command=self._verificar_atualizacoes).pack(side="left", padx=(8,0))
        self.btn_lojas.pack(side="left", padx=2)
        self.btn_senhas.pack(side="left", padx=2)
        self.btn_params.pack(side="left", padx=2)

        self._mostrar("lojas")

    def _mostrar(self, aba):
        # Esconde todas as abas
        for frm in [self.frm_lojas, self.frm_senhas, self.frm_params_aba]:
            frm.pack_forget()
        # Reset todos os botões
        for btn in [self.btn_lojas, self.btn_senhas, self.btn_params]:
            btn.config(bg=BTN, fg=BTN_TXT)
        # Mostra aba selecionada
        if aba == "lojas":
            self.btn_lojas.config(bg=BLUE, fg="white")
            self.frm_lojas.pack(fill="both", expand=True)
            self.root.update_idletasks()
            larg = self.cv_lojas.winfo_width() or 1100
            self._montar_grid(larg)
        elif aba == "senhas":
            self.btn_senhas.config(bg=BLUE, fg="white")
            self.frm_senhas.pack(fill="both", expand=True)
        elif aba == "params":
            self.btn_params.config(bg=BLUE, fg="white")
            self.frm_params_aba.pack(fill="both", expand=True)

    # ── ABA LOJAS ─────────────────────────────────────────────
    def _criar_aba_lojas(self):
        self.frm_lojas=tk.Frame(self.container,bg=BG)

        # Stats
        fs=tk.Frame(self.frm_lojas,bg=BG); fs.pack(fill="x",padx=16,pady=(14,8))
        self.sv={k:tk.StringVar(value=v) for k,v in [("total","21"),("done","0"),("running","0"),("idle","21")]}
        for i,(k,lbl,cor) in enumerate([("total","Total",BRIGHT),("done","Concluídas",SUCCESS),
                                         ("running","Em Execução",WARNING),("idle","Aguardando",CYAN)]):
            fs.columnconfigure(i,weight=1)
            c=tk.Frame(fs,bg=CARD,highlightbackground="#1A4A7A",highlightthickness=1)
            c.grid(row=0,column=i,padx=4,sticky="ew")
            tk.Label(c,textvariable=self.sv[k],font=("Arial",20,"bold"),bg=CARD,fg=cor).pack(pady=(8,2))
            tk.Label(c,text=lbl,font=("Arial",8),bg=CARD,fg=MUTED).pack(pady=(0,8))

        # Top
        ft=tk.Frame(self.frm_lojas,bg=BG); ft.pack(fill="x",padx=16,pady=(2,6))
        tk.Label(ft,text="Minhas ",font=("Arial",13,"bold"),bg=BG,fg=BRIGHT).pack(side="left")
        tk.Label(ft,text="Lojas",  font=("Arial",13,"bold"),bg=BG,fg=CYAN).pack(side="left")
        tk.Button(ft,text="▶▶ Rodar Todas",font=("Arial",9,"bold"),bg=BTN,fg=BTN_TXT,
                  padx=12,pady=4,relief="flat",cursor="hand2",
                  command=self._rodar_todas).pack(side="right")

        # Canvas + scrollbar
        fo=tk.Frame(self.frm_lojas,bg=BG); fo.pack(fill="both",expand=True,padx=16,pady=(0,8))
        sb=ttk.Scrollbar(fo,orient="vertical"); sb.pack(side="right",fill="y")
        self.cv_lojas=tk.Canvas(fo,bg=BG,highlightthickness=0,yscrollcommand=sb.set)
        self.cv_lojas.pack(side="left",fill="both",expand=True)
        sb.config(command=self.cv_lojas.yview)
        self.frm_grid=tk.Frame(self.cv_lojas,bg=BG)
        self._win_lojas=self.cv_lojas.create_window((0,0),window=self.frm_grid,anchor="nw")
        self.frm_grid.bind("<Configure>",lambda e:self.cv_lojas.configure(scrollregion=self.cv_lojas.bbox("all")))
        self.cv_lojas.bind("<Configure>",self._on_lojas_resize)
        self.cv_lojas.bind_all("<MouseWheel>",lambda e:self.cv_lojas.yview_scroll(int(-1*(e.delta/120)),"units"))

        # Cria cards
        for nome,script in LOJAS:
            card=tk.Frame(self.frm_grid,bg=CARD,highlightbackground="#1A4A7A",highlightthickness=1)
            barra=tk.Frame(card,bg=BLUE,width=4); barra.pack(side="left",fill="y")
            inn=tk.Frame(card,bg=CARD); inn.pack(side="left",fill="both",expand=True,padx=10,pady=10)
            tk.Label(inn,text=nome,font=("Arial",11,"bold"),bg=CARD,fg=BRIGHT,anchor="w").pack(fill="x")
            lbl_s=tk.Label(inn,text="Aguardando",font=("Arial",8,"bold"),bg=CARD,fg=MUTED,anchor="w")
            lbl_s.pack(fill="x",pady=(2,6))
            fb=tk.Frame(inn,bg=CARD); fb.pack(fill="x")
            btn_r=tk.Button(fb,text="▶ Rodar",font=("Arial",9,"bold"),bg=BTN,fg=BTN_TXT,
                            padx=8,pady=3,relief="flat",cursor="hand2",
                            command=lambda n=nome,s=script:self._rodar(n,s))
            btn_r.pack(side="left",fill="x",expand=True,padx=(0,4))
            tk.Button(fb,text="📁",font=("Arial",10),bg=BTN,fg=BTN_TXT,padx=6,pady=3,
                      relief="flat",cursor="hand2",
                      command=lambda n=nome:self._pasta(n)).pack(side="right")
            log=scrolledtext.ScrolledText(inn,height=3,font=("Courier New",7),
                                          bg="#000D1A",fg="#5EE89A",state="disabled",relief="flat")
            self.widgets[nome]={"card":card,"barra":barra,"status":lbl_s,"btn":btn_r,"log":log}

    def _on_lojas_resize(self,e):
        self.cv_lojas.itemconfig(self._win_lojas,width=e.width)
        self._montar_grid(e.width)

    def _montar_grid(self,larg=1100):
        ncols=max(1,int(larg)//260)
        for i,(nome,_) in enumerate(LOJAS):
            w=self.widgets.get(nome)
            if w: w["card"].grid(row=i//ncols,column=i%ncols,padx=5,pady=5,sticky="nsew")
        for c in range(ncols): self.frm_grid.columnconfigure(c,weight=1)

    def _upd_stats(self):
        v=list(self.status_lojas.values())
        self.sv["done"].set(str(v.count("done")))
        self.sv["running"].set(str(v.count("running")))
        self.sv["idle"].set(str(v.count("idle")))

    def _rodar(self,nome,script):
        c=get_cred_loja(self.usuario,nome)
        if not c.get("email") or not c.get("senha"):
            messagebox.showwarning("Credenciais",
                f"Configure email e senha para '{nome}'\nem Configurações."); return
        if self.status_lojas[nome]=="running": return
        st=gerar_script(self.usuario,nome,script)
        if not st: messagebox.showerror("Erro",f"Script não encontrado: {script}"); return
        w=self.widgets[nome]
        self.status_lojas[nome]="running"
        w["btn"].config(text="⏳ Rodando...",bg=WARNING,fg="#001228",state="disabled")
        w["status"].config(text="Em execução...",fg=WARNING)
        w["card"].config(highlightbackground=WARNING); w["barra"].config(bg=WARNING)
        w["log"].config(state="normal"); w["log"].delete(1.0,tk.END)
        w["log"].insert(tk.END,f"Iniciando {nome}...\n"); w["log"].config(state="disabled")
        w["log"].pack(fill="x",padx=6,pady=(0,4))
        self._upd_stats()
        env=os.environ.copy(); env["PYTHONIOENCODING"]="utf-8"; env["PYTHONUTF8"]="1"
        def run():
            try:
                p=subprocess.Popen([sys.executable,"-u",str(st)],
                    stdout=subprocess.PIPE,stderr=subprocess.STDOUT,cwd=str(st.parent),env=env)
                for lb in p.stdout:
                    ln=lb.decode("utf-8",errors="replace")
                    w["log"].config(state="normal"); w["log"].insert(tk.END,ln)
                    w["log"].see(tk.END); w["log"].config(state="disabled")
                p.wait(); ok=p.returncode==0
            except Exception as e:
                ok=False; w["log"].config(state="normal")
                w["log"].insert(tk.END,f"ERRO: {e}\n"); w["log"].config(state="disabled")
            na=nome.lower().replace(" ","_")
            pt=st.parent/f"relatorios_{na}"; pu=PASTA/f"relatorios_{self.usuario}_{na}"
            if pt.exists():
                if pu.exists(): shutil.rmtree(str(pu))
                shutil.move(str(pt),str(pu))
            cor=SUCCESS if ok else DANGER
            w["btn"].config(text="✓ Concluído" if ok else "✗ Erro",bg=cor,
                            fg="#001228" if ok else "white",state="normal")
            w["status"].config(text="Concluído" if ok else "Erro",fg=cor)
            w["card"].config(highlightbackground=cor); w["barra"].config(bg=cor)
            self.status_lojas[nome]="done" if ok else "error"; self._upd_stats()
        threading.Thread(target=run,daemon=True).start()

    def _pasta(self,nome):
        na=nome.lower().replace(" ","_")
        p=PASTA/f"relatorios_{self.usuario}_{na}"; p.mkdir(exist_ok=True)
        subprocess.Popen(f'explorer "{p}"')

    def _rodar_todas(self):
        if not messagebox.askyesno("Rodar Todas","Rodar os 21 scripts em sequência?"): return
        def seq():
            import time
            for n,s in LOJAS:
                if self.status_lojas[n]!="running":
                    self._rodar(n,s)
                    while self.status_lojas[n]=="running": time.sleep(1)
        threading.Thread(target=seq,daemon=True).start()

    # ── ABA SENHAS ────────────────────────────────────────────
    def _criar_aba_senhas(self):
        self.frm_senhas = tk.Frame(self.container, bg=BG)

        # Frame com scroll
        sb = ttk.Scrollbar(self.frm_senhas, orient="vertical")
        sb.pack(side="right", fill="y")
        cv = tk.Canvas(self.frm_senhas, bg=BG, highlightthickness=0, yscrollcommand=sb.set)
        cv.pack(side="left", fill="both", expand=True)
        sb.config(command=cv.yview)
        frm = tk.Frame(cv, bg=BG)
        win = cv.create_window((0, 0), window=frm, anchor="nw")
        frm.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.bind("<Configure>", lambda e: cv.itemconfig(win, width=e.width))
        cv.bind_all("<MouseWheel>", lambda e: cv.yview_scroll(int(-1*(e.delta/120)), "units"))

        # Seção credenciais
        o = tk.Frame(frm, bg=CARD, highlightbackground="#1A4A7A", highlightthickness=1)
        o.pack(fill="x", padx=16, pady=(14,5))
        tk.Label(o, text="Credenciais SportBay Hub — Suas Lojas",
                 font=("Arial",10,"bold"), bg=CARD, fg=CYAN).pack(anchor="w", padx=14, pady=(10,4))
        tk.Frame(o, bg="#1A4A7A", height=1).pack(fill="x", padx=14)
        sc = tk.Frame(o, bg=CARD); sc.pack(fill="x", padx=14, pady=10)

        tk.Label(sc, text="Configure o email e senha do SportBay Hub para cada loja.",
                 font=("Arial",9), bg=CARD, fg=MUTED).pack(anchor="w", pady=(0,8))

        self.vars_cred = {}
        for nome in [n for n,_ in LOJAS]:
            c = get_cred_loja(self.usuario, nome)
            row = tk.Frame(sc, bg=CARD); row.pack(fill="x", pady=2)
            tk.Label(row, text=nome, font=("Arial",9,"bold"), bg=CARD, fg=BRIGHT,
                     width=16, anchor="w").pack(side="left")
            tk.Label(row, text="Email:", font=("Arial",8), bg=CARD, fg=MUTED).pack(side="left", padx=(0,3))
            ve = tk.StringVar(value=c.get("email",""))
            tk.Entry(row, textvariable=ve, font=("Arial",9), bg="#001020", fg=BRIGHT,
                     insertbackground=CYAN, relief="flat",
                     highlightbackground="#1A4A7A", highlightthickness=1,
                     width=26).pack(side="left", ipady=3, padx=(0,8))
            tk.Label(row, text="Senha:", font=("Arial",8), bg=CARD, fg=MUTED).pack(side="left", padx=(0,3))
            vs = tk.StringVar(value=c.get("senha",""))
            tk.Entry(row, textvariable=vs, font=("Arial",9), bg="#001020", fg=BRIGHT,
                     insertbackground=CYAN, relief="flat",
                     highlightbackground="#1A4A7A", highlightthickness=1,
                     width=18, show="●").pack(side="left", ipady=3)
            self.vars_cred[nome] = (ve, vs)

        tk.Button(sc, text="💾  Salvar Credenciais", font=("Arial",10,"bold"),
                  bg=BLUE, fg="white", relief="flat", cursor="hand2", pady=7,
                  command=self._salvar_creds).pack(anchor="w", pady=(10,0))

        # Seção trocar senha
        o2 = tk.Frame(frm, bg=CARD, highlightbackground="#1A4A7A", highlightthickness=1)
        o2.pack(fill="x", padx=16, pady=5)
        tk.Label(o2, text="Alterar Minha Senha", font=("Arial",10,"bold"),
                 bg=CARD, fg=CYAN).pack(anchor="w", padx=14, pady=(10,4))
        tk.Frame(o2, bg="#1A4A7A", height=1).pack(fill="x", padx=14)
        ss = tk.Frame(o2, bg=CARD); ss.pack(fill="x", padx=14, pady=10)

        fp2 = tk.Frame(ss, bg=CARD); fp2.pack(fill="x")
        self.vars_pw = []
        for ci,(lbl,show) in enumerate([("Senha Atual","●"),("Nova Senha","●"),("Confirmar","●")]):
            fp2.columnconfigure(ci, weight=1)
            f = tk.Frame(fp2, bg=CARD); f.grid(row=0, column=ci, padx=6, sticky="ew")
            tk.Label(f, text=lbl.upper(), font=("Arial",8,"bold"), bg=CARD, fg=MUTED).pack(anchor="w")
            var = tk.StringVar(); self.vars_pw.append(var)
            tk.Entry(f, textvariable=var, font=("Arial",11), bg="#001020", fg=BRIGHT,
                     insertbackground=CYAN, relief="flat",
                     highlightbackground="#1A4A7A", highlightthickness=1,
                     show=show).pack(fill="x", ipady=6)
        tk.Button(ss, text="🔒  Alterar Senha", font=("Arial",10,"bold"),
                  bg=BLUE, fg="white", relief="flat", cursor="hand2", pady=7,
                  command=self._alterar_senha).pack(anchor="w", pady=(10,0))

    # ── ABA PARÂMETROS ────────────────────────────────────────
    def _criar_aba_params(self):
        self.frm_params_aba = tk.Frame(self.container, bg=BG)

        # Seção tabelas
        o = tk.Frame(self.frm_params_aba, bg=CARD, highlightbackground="#1A4A7A", highlightthickness=1)
        o.pack(fill="x", padx=16, pady=(14,5))
        tk.Label(o, text="Tabelas de Referência", font=("Arial",10,"bold"),
                 bg=CARD, fg=CYAN).pack(anchor="w", padx=14, pady=(10,4))
        tk.Frame(o, bg="#1A4A7A", height=1).pack(fill="x", padx=14)
        st = tk.Frame(o, bg=CARD); st.pack(fill="x", padx=14, pady=10)

        cfg = load_cfg()
        self.vars_tab = {}
        TABS = [
            ("TABELA_PRECOS", "Tabela de Preços Geral",
             getattr(cfg,"TABELA_PRECOS","tabela_precos.xlsx") if cfg else ""),
            ("TABELA_PRECO_SKU_KITS", "Preço SKU dos Kits",
             getattr(cfg,"TABELA_PRECO_SKU_KITS","Preco SKU dos KITS.xlsx") if cfg else ""),
            ("TABELA_MEUS_KITS", "Meus Kits",
             getattr(cfg,"TABELA_MEUS_KITS","meus_kits.xlsx") if cfg else ""),
        ]
        for ci,(chave,lbl,val) in enumerate(TABS):
            st.columnconfigure(ci, weight=1)
            f = tk.Frame(st, bg=CARD); f.grid(row=0, column=ci, padx=6, pady=4, sticky="ew")
            tk.Label(f, text=lbl.upper(), font=("Arial",8,"bold"), bg=CARD, fg=MUTED).pack(anchor="w")
            var = tk.StringVar(value=val); self.vars_tab[chave] = var
            fr = tk.Frame(f, bg=CARD); fr.pack(fill="x")
            ent = tk.Entry(fr, textvariable=var, font=("Arial",9), bg="#001020", fg=BRIGHT,
                     relief="flat", highlightbackground="#1A4A7A", highlightthickness=1)
            ent.pack(side="left", fill="x", expand=True, ipady=5)
            ent.bind("<Key>", lambda e: "break")  # Impede digitação mas mostra o valor
            # Label de status dinâmico
            var_lbl   = tk.StringVar()
            lbl_cor   = [DANGER]
            lbl_widget = tk.Label(f, textvariable=var_lbl, font=("Arial",8), bg=CARD, fg=DANGER)
            lbl_widget.pack(anchor="w", pady=(2,0))
            def _upd(v=var, vl=var_lbl, lw=lbl_widget):
                nome = v.get()
                ok = bool(nome) and (PASTA/nome).exists()
                vl.set("✓ Encontrado: " + nome if ok else "✗ Não encontrado")
                lw.config(fg=SUCCESS if ok else DANGER)
            _upd()
            def upload(c=chave, v=var, upd=_upd):
                p = filedialog.askopenfilename(
                    title=f"Selecione o arquivo",
                    filetypes=[("Excel","*.xlsx *.xls"),("Todos","*.*")])
                if not p: return
                src = Path(p)
                try:
                    dst = PASTA / src.name
                    shutil.copy2(str(src), str(dst))
                    v.set(src.name)
                    save_cfg_val(c, src.name)
                    upd()
                    messagebox.showinfo("Sucesso",
                        f"Arquivo copiado com sucesso!\n\n"
                        f"Arquivo: {src.name}\n"
                        f"Local: {dst}")
                except Exception as e:
                    messagebox.showerror("Erro ao copiar arquivo", str(e))
            tk.Button(fr, text="📂", font=("Arial",10), bg=BLUE, fg="white",
                      padx=6, pady=4, relief="flat", cursor="hand2",
                      command=upload).pack(side="right", padx=(3,0))

        # Seção parâmetros financeiros
        o2 = tk.Frame(self.frm_params_aba, bg=CARD, highlightbackground="#1A4A7A", highlightthickness=1)
        o2.pack(fill="x", padx=16, pady=5)
        tk.Label(o2, text="Parâmetros Financeiros", font=("Arial",10,"bold"),
                 bg=CARD, fg=CYAN).pack(anchor="w", padx=14, pady=(10,4))
        tk.Frame(o2, bg="#1A4A7A", height=1).pack(fill="x", padx=14)
        sf = tk.Frame(o2, bg=CARD); sf.pack(fill="x", padx=14, pady=10)

        tk.Label(sf, text="Usados para calcular o Preço Mínimo de venda.",
                 font=("Arial",9), bg=CARD, fg=MUTED).pack(anchor="w", pady=(0,10))

        PARAMS = [
            ("Margem Mínima Padrão (%)",      "16", "Quando não encontrada na tabela"),
            ("Margem Mínima Alternativa (%)", "12", "Premium / secundária"),
            ("Alíquota de Impostos (%)",      "17", "Aplicada no preço mínimo"),
        ]
        fp = tk.Frame(sf, bg=CARD); fp.pack(fill="x")
        for ci,(lbl,val,hint) in enumerate(PARAMS):
            fp.columnconfigure(ci, weight=1)
            f = tk.Frame(fp, bg=CARD); f.grid(row=0, column=ci, padx=8, sticky="ew")
            tk.Label(f, text=lbl.upper(), font=("Arial",8,"bold"), bg=CARD, fg=MUTED).pack(anchor="w")
            var = tk.StringVar(value=val)
            tk.Entry(f, textvariable=var, font=("Arial",18,"bold"), bg="#001020", fg=CYAN,
                     insertbackground=CYAN, relief="flat",
                     highlightbackground="#1A4A7A", highlightthickness=1,
                     justify="center").pack(fill="x", ipady=10)
            tk.Label(f, text=hint, font=("Arial",8), bg=CARD, fg=MUTED).pack(anchor="w", pady=(3,0))

        # Fórmula
        ff = tk.Frame(sf, bg="#001830", highlightbackground="#003060", highlightthickness=1)
        ff.pack(fill="x", pady=(14,0))
        tk.Label(ff, text="FÓRMULA DO PREÇO MÍNIMO", font=("Arial",8,"bold"),
                 bg="#001830", fg=MUTED).pack(anchor="w", padx=14, pady=(10,2))
        tk.Label(ff, text="Preço Mínimo = (Custo + Frete) ÷ (1 − Margem% − Imposto% − Taxa_ML%)",
                 font=("Courier New",11), bg="#001830", fg=BRIGHT).pack(anchor="w", padx=14)
        tk.Label(ff, text="Frete = Tarifa Frete Grátis (R$)   |   Taxa_ML = Porcentagem cobrada do ML (%)",
                 font=("Arial",9), bg="#001830", fg=MUTED).pack(anchor="w", padx=14, pady=(2,10))

        # Botão Salvar
        def salvar_params():
            messagebox.showinfo("Salvo", "Parâmetros salvos com sucesso!")
        tk.Button(self.frm_params_aba, text="💾  Salvar Parâmetros",
                  font=("Arial",11,"bold"), bg=BLUE, fg="white",
                  relief="flat", cursor="hand2", pady=9,
                  command=salvar_params).pack(anchor="w", padx=16, pady=(12,0))

    def _salvar_creds(self):
        for n,(ve,vs) in self.vars_cred.items():
            salvar_cred_loja(self.usuario,n,ve.get().strip(),vs.get().strip())
        messagebox.showinfo("Salvo","Credenciais salvas com seguranca!")

    def _verificar_atualizacoes(self):
        try:
            from atualizador import verificar_e_atualizar, reiniciar_aplicacao
        except ImportError:
            messagebox.showerror("Erro","Modulo atualizador.py nao encontrado."); return

        win = tk.Toplevel(self.root)
        win.title("Atualizacoes")
        win.geometry("400x220")
        win.configure(bg=BG)
        win.resizable(False, False)
        set_icon(win)
        win.update_idletasks()
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        win.geometry("400x220+" + str((sw-400)//2) + "+" + str((sh-220)//2))

        tk.Label(win, text="Verificando atualizacoes...",
                 font=("Arial",12,"bold"), bg=BG, fg=BRIGHT).pack(pady=(30,8))
        lbl_msg = tk.Label(win, text="Conectando ao servidor...",
                           font=("Arial",9), bg=BG, fg=MUTED, wraplength=360)
        lbl_msg.pack(pady=(0,10))
        bar = ttk.Progressbar(win, length=340, mode="determinate")
        bar.pack(pady=6)
        btn_ok = tk.Button(win, text="Fechar", font=("Arial",10,"bold"),
                           bg=BTN, fg=BTN_TXT, relief="flat", cursor="hand2",
                           pady=7, state="disabled",
                           command=win.destroy)
        btn_ok.pack(pady=(12,0), padx=40, fill="x")

        def on_progresso(msg, pct):
            lbl_msg.config(text=msg); bar["value"] = pct; win.update_idletasks()

        def on_fim(atualizado, versao, notas):
            bar["value"] = 100
            if atualizado:
                lbl_msg.config(text="Atualizado para v" + str(versao) + "! Reiniciando...")
                win.update_idletasks()
                win.after(1800, lambda: (win.destroy(), reiniciar_aplicacao()))
            else:
                lbl_msg.config(text=notas or "Ja esta na versao mais recente.")
                btn_ok.config(state="normal")

        verificar_e_atualizar(
            callback_progresso=lambda m,p: win.after(0, on_progresso, m, p),
            callback_fim=lambda a,v,n: win.after(0, on_fim, a, v, n),
            silencioso=False
        )

    def _alterar_senha(self):
        if not verificar_senha(self.usuario,self.vars_pw[0].get().strip()):
            messagebox.showerror("Erro","Senha atual incorreta."); return
        n=self.vars_pw[1].get().strip()
        if len(n)<4: messagebox.showerror("Erro","Mínimo 4 caracteres."); return
        if n!=self.vars_pw[2].get().strip(): messagebox.showerror("Erro","As senhas não coincidem."); return
        trocar_senha_user(self.usuario,n)
        for v in self.vars_pw: v.set("")
        messagebox.showinfo("Sucesso","Senha alterada!")

if __name__=="__main__":
    App()
