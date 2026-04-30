# -*- coding: utf-8 -*-
"""
Módulo de atualização automática via GitHub
Classic Sports - Automação de Relatórios
"""
import sys
import os
import json
import shutil
import subprocess
import threading
import urllib.request
import urllib.error
from pathlib import Path

# ── CONFIGURAÇÃO DO REPOSITÓRIO ───────────────────────────────
# Após criar seu repositório no GitHub, atualize estas linhas:
GITHUB_USER    = "Marcus-Carvalho"
GITHUB_REPO    = "classicsports-relatorios"
GITHUB_BRANCH  = "main"

# URLs base
URL_VERSION = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}/versao.json"
URL_RAW_BASE = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}/arquivos"

PASTA_LOCAL = Path(__file__).parent
ARQ_VERSION_LOCAL = PASTA_LOCAL / "versao.json"

# ── ARQUIVOS QUE SERÃO ATUALIZADOS AUTOMATICAMENTE ───────────
ARQUIVOS_PARA_ATUALIZAR = [
    "sportbay_painel.py",
    "atualizador.py",
    "sportbay_11.py",
    "sportbay_12.py",
    "sportbay_708.py",
    "sportbay_adrenalinex.py",
    "sportbay_adventurex.py",
    "sportbay_am15.py",
    "sportbay_am20.py",
    "sportbay_bl.py",
    "sportbay_classic_barracao.py",
    "sportbay_cs.py",
    "sportbay_dl.py",
    "sportbay_dm.py",
    "sportbay_ff.py",
    "sportbay_gj.py",
    "sportbay_imports.py",
    "sportbay_ja.py",
    "sportbay_juliana.py",
    "sportbay_marcus.py",
    "sportbay_planet.py",
    "sportbay_ras.py",
    "sportbay_ss.py",
]

def get_versao_local():
    """Lê a versão instalada localmente."""
    if not ARQ_VERSION_LOCAL.exists():
        return "0.0.0"
    try:
        dados = json.loads(ARQ_VERSION_LOCAL.read_text(encoding="utf-8"))
        return dados.get("versao", "0.0.0")
    except Exception:
        return "0.0.0"

def get_versao_remota():
    """Busca a versão mais recente no GitHub. Retorna (versao, notas) ou None."""
    try:
        with urllib.request.urlopen(URL_VERSION, timeout=5) as resp:
            dados = json.loads(resp.read().decode("utf-8"))
            return dados.get("versao", "0.0.0"), dados.get("notas", "")
    except Exception:
        return None, None

def versao_maior(nova, atual):
    """Compara versões no formato X.Y.Z."""
    try:
        n = tuple(int(x) for x in nova.split("."))
        a = tuple(int(x) for x in atual.split("."))
        return n > a
    except Exception:
        return False

def baixar_arquivo(nome_arquivo):
    """Baixa um arquivo do GitHub e substitui o local."""
    url = f"{URL_RAW_BASE}/{nome_arquivo}"
    destino = PASTA_LOCAL / nome_arquivo
    destino_bak = PASTA_LOCAL / f"{nome_arquivo}.bak"

    try:
        # Faz backup do arquivo atual
        if destino.exists():
            shutil.copy2(str(destino), str(destino_bak))

        with urllib.request.urlopen(url, timeout=15) as resp:
            conteudo = resp.read()
        destino.write_bytes(conteudo)

        # Remove backup se OK
        if destino_bak.exists():
            destino_bak.unlink()
        return True
    except Exception as e:
        # Restaura backup em caso de erro
        if destino_bak.exists():
            shutil.copy2(str(destino_bak), str(destino))
            destino_bak.unlink()
        return False

def salvar_versao_local(versao, notas=""):
    """Salva a versão instalada localmente."""
    dados = {"versao": versao, "notas": notas}
    ARQ_VERSION_LOCAL.write_text(
        json.dumps(dados, indent=2, ensure_ascii=False), encoding="utf-8")

def verificar_e_atualizar(callback_progresso=None, callback_fim=None, silencioso=False):
    """
    Verifica e aplica atualização em thread separada.
    
    callback_progresso(msg, pct): chamado durante o processo
    callback_fim(atualizado, nova_versao, notas): chamado ao terminar
    silencioso: se True, não avisa quando já está na versão mais recente
    """
    def _run():
        versao_local  = get_versao_local()
        versao_remote, notas = get_versao_remota()

        if versao_remote is None:
            # Sem conexão ou repositório não configurado
            if callback_fim:
                callback_fim(False, versao_local, "Sem conexão com o servidor de atualizações.")
            return

        if not versao_maior(versao_remote, versao_local):
            if callback_fim:
                callback_fim(False, versao_local,
                             "Você já está na versão mais recente." if not silencioso else "")
            return

        # Há atualização disponível — baixa os arquivos
        total = len(ARQUIVOS_PARA_ATUALIZAR)
        erros = []

        for i, arq in enumerate(ARQUIVOS_PARA_ATUALIZAR):
            if callback_progresso:
                pct = int((i / total) * 100)
                callback_progresso(f"Atualizando: {arq}", pct)
            ok = baixar_arquivo(arq)
            if not ok:
                erros.append(arq)

        if not erros:
            salvar_versao_local(versao_remote, notas)
            # Apaga credenciais.enc para forçar reconfiguração na próxima abertura.
            # Cada máquina precisa gerar seu próprio .enc com a chave local.
            arq_cred = PASTA_LOCAL / "dados" / "credenciais.enc"
            if arq_cred.exists():
                try:
                    arq_cred.unlink()
                except Exception:
                    pass

        if callback_fim:
            callback_fim(True, versao_remote, notas)

    threading.Thread(target=_run, daemon=True).start()

def reiniciar_aplicacao():
    """Reinicia o painel após atualização."""
    script = PASTA_LOCAL / "sportbay_painel.py"
    subprocess.Popen([sys.executable, str(script)], cwd=str(PASTA_LOCAL))
    sys.exit(0)
