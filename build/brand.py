"""Identidade visual HayaiDataSystems (cores extraídas do logo oficial)."""
from pathlib import Path

BUILD = Path(__file__).resolve().parent
ROOT = BUILD.parent                      # Treinamento_HayaiDataSystems/
PROJECT = ROOT.parent                    # pasta do curso
ASSETS = BUILD / "assets"
LOGO_SRC = PROJECT / "IMAGEM DA EMPRESA MODELO DE CORES E IDENTIDADE VISUAL.png"

COURSE = "Design Thinking na Prática"
SUBTITLE = "Inovação centrada nas pessoas para empresas"
BRAND = "HayaiDataSystems"
TAGLINE = "Dados que potencializam decisões"
INSTRUCTOR = "Milton Carlos Ribeiro"
HOURS = "12 horas · 3 encontros de 4h"

# Paleta (hex sem '#')
BG = "1A1A1A"        # fundo do logo
SURFACE = "242826"   # cartões
SURFACE2 = "2E3431"  # cartões em destaque / bordas
LINE = "3A413D"
TEXT = "F2F2EC"      # branco quente do logo
MUTED = "B4B8B0"     # prata do logotipo
DIM = "7D827B"

LIME = "9CD848"      # verde-limão da seta
GREEN = "4DB35A"     # verde da nuvem
TEAL = "2FA38A"      # verde-azulado de "Data"
BLUE = "2A9FC0"      # azul dos circuitos
DEEP = "1F7FA8"      # azul profundo

# Uma cor por etapa do Design Thinking (motivo visual recorrente)
STAGE_COLORS = {
    "Empatizar": LIME,
    "Definir": GREEN,
    "Idear": TEAL,
    "Prototipar": BLUE,
    "Testar": DEEP,
}
MODULE_COLORS = {1: LIME, 2: LIME, 3: GREEN, 4: TEAL, 5: BLUE, 6: DEEP}


def rgb(hexstr):
    return tuple(int(hexstr[i:i + 2], 16) for i in (0, 2, 4))
