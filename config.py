# ═══════════════════════════════════════════════════════════════
# CONFIGURACAO - ARQUIVOS DE REFERENCIA
# Atualize aqui quando os arquivos mudarem. Nao mexa nos scripts!
# ═══════════════════════════════════════════════════════════════

# Tabela de precos geral (por SKU)
TABELA_PRECOS = "tabela_precos_20260408_1740.xlsx"

# Tabela de precos dos Kits (por titulo)
TABELA_PRECO_SKU_KITS = "Preco SKU dos KITS - 08-04-26.xlsx"

# Tabela Meus Kits (composicao dos kits do sistema)
TABELA_MEUS_KITS = "meus_kits_2026-04-08.xlsx"

# ═══════════════════════════════════════════════════════════════
# CONFIGURACAO - LOJAS E PERFIS DO CHROME
# ═══════════════════════════════════════════════════════════════

import os as _os
CHROME_PROFILE_DIR = _os.path.join(_os.path.expanduser("~"), "AppData", "Local", "Google", "Chrome", "User Data")

LOJAS = [
    ("Classic Barracao", "Profile 1",  "acessorioscarrosantigos@gmail.com", "YnZiUg3YF58k!Gd"),
    ("CS",               "Profile 10", "classic.sports.plus@gmail.com",     "EBGReXjP3@36Sp#"),
    ("Juliana",          "Profile 11", "classic.sports.digital@gmail.com",  "shC!4VQr2Se4wg7"),
    ("Marcus",           "Profile 12", "classic.sports.fullcommerce@gmail.com", "rKBH@G@Natdu8@E"),
    ("11",               "Profile 13", "classic.sports.full.11@gmail.com",  "CM86EFh.PAbxqPr"),
    ("12",               "Profile 14", "classic.sports.full.12@gmail.com",  "E8CQ4g!AxAzeAjX"),
    ("Imports",          "Profile 19", "financeiro@classicsports.com.br",   "!Ap7hpCu8LARxVK"),
    ("708",              "Profile 2",  "contato@classicsports.com.br",      "pM#NAQ4azYiE*SJ"),
    ("AdrenalineX",      "Profile 20", "gevanracing808@gmail.com",          "QfiJ747Sf@HsgUL"),
    ("AdventureX",       "Profile 21", "drijuegevanracing@gmail.com",       "Nk5wcawBQ@C2GPR"),
    ("AM15",             "Profile 24", "amracing15@gmail.com",              "sRiZSaM6VqNUQ3d!"),
    ("AM20",             "Profile 25", "amracing20@gmail.com",              "3RpAyWrc@3RCrqe"),
    ("Planet",           "Profile 27", "planetmotors1969@gmail.com",        "!mH9qbkuTBjFEp6"),
    ("RAS",              "Profile 3",  "classic.sports.ml@gmail.com",       "u8!eT@FfsgywvNf"),
    ("DL",               "Profile 30", "danilodelai02@gmail.com",           "QMC7DJpL2jcb5v3!"),
    ("DM",               "Profile 31", "dmmotos02@gmail.com",               "BF@JR6!H3vDm9Aa"),
    ("BL",               "Profile 32", "blmotos02@gmail.com",               "2T4kmSX!qCfesQ3"),
    ("GJ",               "Profile 5",  "classic.sports.brasil@gmail.com",   "V@AqcJ47uwch@qy"),
    ("FF",               "Profile 7",  "classic.sports.ecommerce@gmail.com","23r@88D5FBx8A62"),
    ("JA",               "Profile 8",  "classic.sports.full@gmail.com",     "EnV7CFbSDe@@2zQ"),
    ("SS",               "Profile 9",  "marcus.classic.sports@gmail.com",   "JKsqg4nw78JA.r#"),
]
