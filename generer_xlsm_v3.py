# -*- coding: utf-8 -*-
"""
generer_xlsm_v3.py
================================================================
Génère le classeur Excel professionnel :
    Tableau_suivi_gasoil_et_production_eau_v3.xlsm

Contenu :
 - 11 onglets (Paramètres, Référentiels, Saisie_Gasoil, Saisie_Eau,
   Saisie_Rapide_Gasoil, Saisie_Rapide_Eau, Synthèse_Gasoil,
   Synthèse_Eau, Tableau_Bord, Fiche_Terrain, Mode_Emploi)
 - styles harmonisés, validations de données, mises en forme
   conditionnelles, tableaux structurés, noms définis,
 - 5 graphiques professionnels sur le Tableau_Bord,
 - modules VBA exportés dans le dossier ./vba/,
 - tentative d'injection automatique des macros via pywin32
   (Windows + Excel). Sinon : instructions d'import manuel.

Prérequis :
    pip install openpyxl
    pip install pywin32   # optionnel, Windows + Excel uniquement

Utilisation :
    python generer_xlsm_v3.py
================================================================
"""

import os
import sys
import traceback

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import LineChart, BarChart, PieChart, Reference, Series
from openpyxl.chart.marker import Marker
from openpyxl.chart.data_source import NumDataSource, NumRef
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.marker import DataPoint
from openpyxl.formatting.rule import CellIsRule, FormulaRule, DataBarRule
from openpyxl.workbook.defined_name import DefinedName

# ================================================================
# CONFIGURATION GÉNÉRALE
# ================================================================

FICHIER_SANS_MACROS = "Tableau_suivi_gasoil_et_production_eau_v3.xlsx"
FICHIER_XLSM        = "Tableau_suivi_gasoil_et_production_eau_v3.xlsm"
DOSSIER_VBA         = "vba"

# ----------------------------------------------------------------
# PALETTE DE COULEURS DU CLASSEUR
# ----------------------------------------------------------------
BLUE_DARK  = "1F4E78"   # bleu foncé (titres, accents)
BLUE_MED   = "2E75B6"   # bleu moyen (en-têtes de colonnes)
BLUE_LIGHT = "D9E1F2"   # bleu clair (sections)
BLUE_PALE  = "EAF1FB"   # bleu très pâle (valeurs KPI)
INPUT_FILL = "FFF2CC"   # jaune : cellules de saisie
CALC_FILL  = "F2F2F2"   # gris  : cellules calculées
ALERT_RED  = "FFC7CE"   # rouge : alerte / erreur
ALERT_ORG  = "FFEB9C"   # orange : avertissement
OK_GREEN   = "C6EFCE"   # vert : conformité
GREY_LIGHT = "F8F9FA"   # gris très clair (fond secondaire)
GREY_BORDER = "BFBFBF"  # bordures fines
DARK_GREEN  = "507E32"  # contour des barres vertes

# Palette harmonisée pour les graphiques
CHART_COLORS = ["1F4E78", "2E75B6", "ED7D31", "70AD47",
                "FFC000", "C00000", "7030A0", "808080"]

NB_LIGNES_GASOIL = 100   # lignes de saisie gasoil (4 à 103)
NB_LIGNES_EAU    = 200   # lignes de saisie eau    (4 à 203)


# ================================================================
# HELPERS DE STYLE
# ================================================================

def fill(hex_color):
    """Retourne un remplissage uni à partir d'une couleur hexa."""
    return PatternFill("solid", fgColor=hex_color)


def thin_border():
    """Retourne une bordure fine grise standard."""
    side = Side(style="thin", color=GREY_BORDER)
    return Border(left=side, right=side, top=side, bottom=side)


def title(ws, cell_range, text, size=14, height=28):
    """Titre principal : fond bleu foncé, texte blanc, gras, indent 1."""
    ws.merge_cells(cell_range)
    first = cell_range.split(":")[0]
    c = ws[first]
    c.value = text
    c.font = Font(bold=True, color="FFFFFF", size=size)
    c.fill = fill(BLUE_DARK)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    row = int("".join(ch for ch in first if ch.isdigit()))
    ws.row_dimensions[row].height = height
    # Applique le fond sur toute la plage fusionnée (bord droit propre)
    for col_cells in ws[cell_range]:
        for cc in col_cells:
            cc.fill = fill(BLUE_DARK)
    return c


def section(ws, cell_range, text):
    """Section : fond bleu clair, texte bleu foncé, gras."""
    ws.merge_cells(cell_range)
    first = cell_range.split(":")[0]
    c = ws[first]
    c.value = text
    c.font = Font(bold=True, color=BLUE_DARK, size=11)
    c.fill = fill(BLUE_LIGHT)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for col_cells in ws[cell_range]:
        for cc in col_cells:
            cc.fill = fill(BLUE_LIGHT)
            cc.border = thin_border()
    return c


def header(ws, row, start_col, headers_list):
    """En-têtes de colonnes : fond bleu moyen, texte blanc, centré, wrap."""
    for i, h in enumerate(headers_list):
        c = ws.cell(row=row, column=start_col + i, value=h)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = fill(BLUE_MED)
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        c.border = thin_border()
    ws.row_dimensions[row].height = 42


def label(ws, coord, text, bold=False, italic=False, color="000000"):
    """Libellé classique (feuille de type formulaire)."""
    c = ws[coord]
    c.value = text
    c.font = Font(size=10, bold=bold, italic=italic, color=color)
    c.alignment = Alignment(horizontal="left", vertical="center")
    return c


def input_cell(ws, coord, fmt=None):
    """Cellule de saisie : fond jaune + bordure fine."""
    c = ws[coord]
    c.fill = fill(INPUT_FILL)
    c.border = thin_border()
    c.alignment = Alignment(horizontal="left", vertical="center")
    if fmt:
        c.number_format = fmt
    return c


def calc_cell(ws, coord, formula=None, fmt=None, bold=False,
              font_color="000000", align="center", bg=CALC_FILL):
    """Cellule calculée : fond gris + bordure fine."""
    c = ws[coord]
    if formula is not None:
        c.value = formula
    c.fill = fill(bg)
    c.border = thin_border()
    c.font = Font(bold=bold, color=font_color, size=10)
    c.alignment = Alignment(horizontal=align, vertical="center")
    if fmt:
        c.number_format = fmt
    return c


def set_widths(ws, mapping):
    """Définit les largeurs de colonnes : {'A': 5, 'B': 40, ...}."""
    for col, width in mapping.items():
        ws.column_dimensions[col].width = width


def styled_chart(ch, title_txt, x_title=None, y_title=None, style=2):
    """Harmonise titre / axes / style de tous les graphiques."""
    ch.title = title_txt
    ch.style = style                      # style intégré Office n°2
    ch.legend.position = "b"
    if x_title is not None:
        ch.x_axis.title = x_title
        ch.y_axis.title = y_title
        # Axes visibles et lisibles
        ch.x_axis.delete = False
        ch.y_axis.delete = False
    else:
        # Cas des camemberts : pas d'axes
        ch.legend.position = "r"
    return ch


def page_setup_landscape(ws):
    """Mise en page paysage + ajustement à la largeur."""
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


# ================================================================
# FEUILLE 1 — PARAMÈTRES
# ================================================================

def creer_parametres(wb):
    """Feuille Paramètres : valeurs globales utilisées par les formules."""
    ws = wb.active
    ws.title = "Paramètres"
    set_widths(ws, {"A": 3, "B": 42, "C": 3, "D": 25})

    title(ws, "B2:D2", "PARAMÈTRES GÉNÉRAUX", size=14, height=30)

    # (libellé, valeur, format nombre)
    champs = [
        ("Nom de l'entreprise", "Mon Entreprise", None),
        ("Site principal", "Site A", None),
        ("Devise", "FCFA", None),
        ("Prix du gasoil (/L)", 750, "#,##0"),
        ("Facteur CO₂ gasoil (kg/L)", 2.68, "0.00"),
        ("Seuil alerte consommation (L/h)", 5, "0.00"),
        ("Seuil alerte charge (%)", 30, "0"),
        ("Seuil alerte fuite (L)", 1, "0.00"),
    ]
    r = 4
    for lib, val, fmt in champs:
        c = ws.cell(row=r, column=2, value=lib)          # colonne B
        c.font = Font(size=10, bold=False)
        c.alignment = Alignment(vertical="center")
        c.border = thin_border()
        v = ws.cell(row=r, column=4, value=val)          # colonne D
        v.fill = fill(INPUT_FILL)
        v.border = thin_border()
        v.alignment = Alignment(horizontal="center", vertical="center")
        if fmt:
            v.number_format = fmt
        ws.row_dimensions[r].height = 20
        r += 1

    note = ws["B13"]
    note.value = ("Ces paramètres alimentent les formulas des synthèses "
                  "(prix, CO₂) et les seuils d'alerte.")
    note.font = Font(size=9, italic=True, color="808080")
    page_setup_landscape(ws)
    return ws


# ================================================================
# FEUILLE 2 — RÉFÉRENTIELS
# ================================================================

REFERENTIELS = {
    "Sites":            ["Site A", "Site B", "Site C"],
    "Équipements":      ["GE-01", "GE-02", "GE-03"],
    "Opérateurs":       ["Opérateur 1", "Opérateur 2", "Opérateur 3"],
    "Types de charge":  ["Pompage", "Éclairage", "Mixte", "Autre"],
    "Débitmètres":      ["DN-01", "DN-02", "DN-03"],
    "Unités eau":       ["m³", "L"],
    "Oui/Non":          ["Oui", "Non"],
    "Sources d'énergie": ["Gasoil", "Solaire", "Réseau", "Mixte"],
}


def creer_referentiels(wb):
    """Feuille Référentiels : listes de valeurs pour les validations."""
    ws = wb.create_sheet("Référentiels")
    title(ws, "A1:H1", "RÉFÉRENTIELS — LISTES DE VALEURS", size=14, height=26)

    # En-tête ligne 3 + valeurs à partir de la ligne 4
    for i, (nom, valeurs) in enumerate(REFERENTIELS.items(), start=1):
        letter = get_column_letter(i)
        c = ws.cell(row=3, column=i, value=nom)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = fill(BLUE_MED)
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        c.border = thin_border()
        for j, v in enumerate(valeurs):
            vc = ws.cell(row=4 + j, column=i, value=v)
            vc.fill = fill(INPUT_FILL)
            vc.border = thin_border()
        ws.column_dimensions[letter].width = 18

    ws.row_dimensions[3].height = 30
    aide = ws["A14"]
    aide.value = ("Ajoutez vos propres valeurs dans les colonnes ci-dessus : "
                  "les menus déroulants des feuilles de saisie s'actualisent.")
    aide.font = Font(size=9, italic=True, color="808080")
    return ws


def validation_liste(ws, ref_plage, cible, titre_msg=None):
    """Ajoute une validation de liste pointant vers Référentiels."""
    dv = DataValidation(type="list",
                        formula1=f"=Référentiels!{ref_plage}",
                        allow_blank=True, showDropDown=False)
    dv.error = "Valeur non autorisée : choisissez un élément de la liste."
    dv.errorTitle = titre_msg or "Liste de valeurs"
    ws.add_data_validation(dv)
    dv.add(cible)
    return dv


# ================================================================
# FEUILLE 3 — SAISIE_GASOIL
# ================================================================

HEADERS_GASOIL = [
    "N°", "Date", "Site", "Code équipement", "Type de charge",
    "Heure début", "Heure fin",
    "Compteur début (h)", "Compteur fin (h)", "Durée mesurée (h)",
    "Niveau initial référence", "Niveau final avant appoint",
    "Gasoil ajouté (L)", "Autres ajouts (L)", "Fuites / pertes (L)",
    "Conso corrigée (L)", "Conso moyenne (L/h)", "Charge estimée (%)",
    "Énergie produite (kWh)", "Conso spécifique (L/kWh)",
    "État / conditions", "Opérateur", "Observations / anomalies",
    "Vérif. niveau identique ?", "Validation responsable",
]


def creer_saisie_gasoil(wb):
    """Feuille de saisie détaillée gasoil (tableau structuré tblGasoil)."""
    ws = wb.create_sheet("Saisie_Gasoil")
    last_row = 3 + NB_LIGNES_GASOIL  # 103

    title(ws, "A1:Y1",
          "SUIVI DE LA CONSOMMATION DE GASOIL — GROUPE ÉLECTROGÈNE",
          size=14, height=26)
    m = ws["A2"]
    m.value = ("Méthode « plein à plein » : faire le plein, relever le "
               "compteur, consigner chaque ajout, replein en fin d'essai.")
    m.font = Font(size=9, italic=True, color="808080")

    header(ws, 3, 1, HEADERS_GASOIL)

    # Largeurs de colonnes adaptées
    largeurs = {"A": 5, "B": 11, "C": 10, "D": 13, "E": 13, "F": 10, "G": 10,
                "H": 12, "I": 12, "J": 11, "K": 13, "L": 13, "M": 12, "N": 12,
                "O": 12, "P": 12, "Q": 12, "R": 11, "S": 13, "T": 13,
                "U": 16, "V": 13, "W": 24, "X": 13, "Y": 13}
    set_widths(ws, largeurs)

    # Formules automatiques + mise en forme des colonnes saisie/calcul
    cols_input = ["B", "C", "D", "E", "F", "G", "H", "I", "K", "L",
                  "M", "N", "O", "R", "S", "U", "V", "W", "X", "Y"]
    fmt_map = {"B": "dd/mm/yyyy", "F": "hh:mm", "G": "hh:mm",
               "H": "0.0", "I": "0.0", "J": "0.00", "K": "0.0", "L": "0.0",
               "M": "0.0", "N": "0.0", "O": "0.0", "P": "0.00", "Q": "0.00",
               "R": "0", "S": "0.0", "T": "0.000"}
    for r in range(4, last_row + 1):
        # N° auto
        calc_cell(ws, f"A{r}", f'=IF(B{r}="","",ROW()-3)', fmt="0")
        # Durée mesurée (colonne J)
        calc_cell(ws, f"J{r}",
                  f'=IF(OR(H{r}="",I{r}=""),"",I{r}-H{r})', fmt="0.00")
        # Conso corrigée (colonne P)
        calc_cell(ws, f"P{r}",
                  f'=IF(M{r}="","",M{r}+IF(N{r}="",0,N{r})-IF(O{r}="",0,O{r}))',
                  fmt="0.00")
        # Conso moyenne (colonne Q)
        calc_cell(ws, f"Q{r}", f'=IFERROR(P{r}/J{r},"")', fmt="0.00")
        # Conso spécifique (colonne T)
        calc_cell(ws, f"T{r}", f'=IFERROR(P{r}/S{r},"")', fmt="0.000")
        # Colonnes de saisie
        for col in cols_input:
            c = ws[f"{col}{r}"]
            c.fill = fill(INPUT_FILL)
            c.border = thin_border()
            if col in fmt_map:
                c.number_format = fmt_map[col]

    # Tableau structuré tblGasoil
    tab = Table(displayName="tblGasoil", ref=f"A3:Y{last_row}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2",
                                        showRowStripes=True,
                                        showColumnStripes=False)
    ws.add_table(tab)

    # Validations de données (listes issues de Référentiels)
    validation_liste(ws, "$A$4:$A$6", f"C4:C{last_row}")     # Sites
    validation_liste(ws, "$B$4:$B$6", f"D4:D{last_row}")     # Équipements
    validation_liste(ws, "$D$4:$D$7", f"E4:E{last_row}")     # Types de charge
    validation_liste(ws, "$C$4:$C$6", f"V4:V{last_row}")     # Opérateurs
    validation_liste(ws, "$G$4:$G$5", f"X4:X{last_row}")     # Oui/Non

    # Nombres décimaux >= 0 pour les volumes M:N:O
    dv_pos = DataValidation(type="decimal", operator="greaterThanOrEqual",
                            formula1="0", allow_blank=True)
    dv_pos.error = "La valeur doit être un nombre positif ou nul."
    dv_pos.errorTitle = "Valeur invalide"
    ws.add_data_validation(dv_pos)
    dv_pos.add(f"M4:O{last_row}")

    # Mises en forme conditionnelles
    ws.conditional_formatting.add(
        f"O4:O{last_row}",
        CellIsRule(operator="greaterThan", formula=["0"],
                   stopIfTrue=False, fill=fill(ALERT_RED)))
    ws.conditional_formatting.add(
        f"X4:X{last_row}",
        CellIsRule(operator="equal", formula=['"Non"'],
                   stopIfTrue=False, fill=fill(ALERT_ORG)))
    ws.conditional_formatting.add(
        f"X4:X{last_row}",
        CellIsRule(operator="equal", formula=['"Oui"'],
                   stopIfTrue=False, fill=fill(OK_GREEN)))
    # Barres de données sur la conso corrigée (colonne P)
    ws.conditional_formatting.add(
        f"P4:P{last_row}",
        DataBarRule(start_type="num", start_value=0,
                    end_type="max", color=BLUE_MED,
                    showValue=True, minLength=None, maxLength=None))

    ws.freeze_panes = "C4"
    page_setup_landscape(ws)
    return ws


# ================================================================
# FEUILLE 4 — SAISIE_EAU
# ================================================================

HEADERS_EAU = [
    "N°", "Date", "Heure", "Site", "Débitmètre / conduite",
    "Index début", "Index fin", "Unité", "Volume", "Durée (h)",
    "Débit moyen (m³/h)", "Énergie utilisée (kWh)", "Source d'énergie",
    "Observations",
]


def creer_saisie_eau(wb):
    """Feuille de saisie production d'eau (tableau structuré tblEau)."""
    ws = wb.create_sheet("Saisie_Eau")
    last_row = 3 + NB_LIGNES_EAU  # 203

    title(ws, "A1:N1", "SUIVI DE LA PRODUCTION D'EAU — DÉBITMÈTRES",
          size=14, height=26)
    m = ws["A2"]
    m.value = ("Relever les index du débitmètre en début et fin de pompage ; "
               "le volume et le débit sont calculés automatiquement.")
    m.font = Font(size=9, italic=True, color="808080")

    header(ws, 3, 1, HEADERS_EAU)

    set_widths(ws, {"A": 5, "B": 11, "C": 9, "D": 10, "E": 18, "F": 12,
                    "G": 12, "H": 8, "I": 12, "J": 10, "K": 14, "L": 15,
                    "M": 13, "N": 26})

    fmt_map = {"B": "dd/mm/yyyy", "C": "hh:mm", "F": "0.0", "G": "0.0",
               "I": "0.0", "J": "0.00", "K": "0.00", "L": "0.0"}
    for r in range(4, last_row + 1):
        calc_cell(ws, f"A{r}", f'=IF(B{r}="","",ROW()-3)', fmt="0")
        calc_cell(ws, f"I{r}",
                  f'=IF(OR(F{r}="",G{r}=""),"",G{r}-F{r})', fmt="0.0")
        calc_cell(ws, f"K{r}", f'=IFERROR(I{r}/J{r},"")', fmt="0.00")
        for col in ["B", "C", "D", "E", "F", "G", "H", "J", "L", "M", "N"]:
            c = ws[f"{col}{r}"]
            c.fill = fill(INPUT_FILL)
            c.border = thin_border()
            if col in fmt_map:
                c.number_format = fmt_map[col]

    tab = Table(displayName="tblEau", ref=f"A3:N{last_row}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2",
                                        showRowStripes=True,
                                        showColumnStripes=False)
    ws.add_table(tab)

    validation_liste(ws, "$A$4:$A$6", f"D4:D{last_row}")      # Sites
    validation_liste(ws, "$E$4:$E$6", f"E4:E{last_row}")      # Débitmètres
    validation_liste(ws, "$F$4:$F$5", f"H4:H{last_row}")      # Unités eau
    validation_liste(ws, "$H$4:$H$7", f"M4:M{last_row}")      # Sources énergie

    dv_pos = DataValidation(type="decimal", operator="greaterThanOrEqual",
                            formula1="0", allow_blank=True)
    dv_pos.error = "La valeur doit être un nombre positif ou nul."
    ws.add_data_validation(dv_pos)
    dv_pos.add(f"F4:G{last_row}")
    dv_pos.add(f"J4:J{last_row}")
    dv_pos.add(f"L4:L{last_row}")

    ws.freeze_panes = "D4"
    page_setup_landscape(ws)
    return ws


# ================================================================
# FEUILLE 5 — SAISIE_RAPIDE_GASOIL (formulaire)
# ================================================================

def creer_saisie_rapide_gasoil(wb):
    """Formulaire de saisie rapide d'un essai gasoil."""
    ws = wb.create_sheet("Saisie_Rapide_Gasoil")
    set_widths(ws, {"A": 2, "B": 40, "C": 3, "D": 28,
                    "E": 3, "F": 40, "G": 3, "H": 28})

    title(ws, "B2:H2", "SAISIE RAPIDE — ESSAI GASOIL", size=16, height=42)
    st = ws["B3"]
    ws.merge_cells("B3:H3")
    st.value = "Remplissez les champs puis cliquez sur « Valider et ajouter »."
    st.font = Font(size=10, italic=True, color="808080")
    st.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # ---- Section 1 : identification (gauche) -------------------
    section(ws, "B5:D5", "1. IDENTIFICATION")
    gauche = [("Date de l'essai *", "dd/mm/yyyy"),
              ("Site *", None),
              ("Code équipement *", None),
              ("Type de charge", None),
              ("Opérateur", None)]
    for i, (lib, fmt) in enumerate(gauche):
        r = 6 + i
        label(ws, f"B{r}", lib)
        input_cell(ws, f"D{r}", fmt=fmt)
        ws.row_dimensions[r].height = 22

    # ---- Section 2 : relevés horaires (droite) ------------------
    section(ws, "F5:H5", "2. RELEVÉS HORAIRES")
    droite = [("Heure début", "hh:mm"),
              ("Heure fin", "hh:mm"),
              ("Compteur début (h)", "0.0"),
              ("Compteur fin (h)", "0.0"),
              ("Charge estimée (%)", "0")]
    for i, (lib, fmt) in enumerate(droite):
        r = 6 + i
        label(ws, f"F{r}", lib)
        input_cell(ws, f"H{r}", fmt=fmt)

    # ---- Section 3 : mesures carburant --------------------------
    section(ws, "B12:H12", "3. MESURES CARBURANT")
    mesures = [("Niveau initial (repère plein)", "0.0"),
               ("Niveau final avant appoint (repère)", "0.0"),
               ("Gasoil ajouté au plein (L) *", "0.0"),
               ("Autres ajouts (L)", "0.0"),
               ("Fuites / pertes estimées (L)", "0.0"),
               ("Énergie produite (kWh)", "0.0")]
    for i, (lib, fmt) in enumerate(mesures):
        r = 13 + i
        label(ws, f"B{r}", lib)
        input_cell(ws, f"D{r}", fmt=fmt)

    # Champs auto-calculés (colonne H13:H16)
    auto = [("Durée mesurée (h)", '=IF(OR(H7="",H8=""),"",H8-H7)', "0.00"),
            ("Conso corrigée (L)",
             '=IF(D15="","",D15+IF(D16="",0,D16)-IF(D17="",0,D17))', "0.00"),
            ("Conso moyenne (L/h)", '=IFERROR(H14/H13,"")', "0.00"),
            ("Conso spécifique (L/kWh)", '=IFERROR(H14/D18,"")', "0.000")]
    for i, (lib, formule, fmt) in enumerate(auto):
        r = 13 + i
        label(ws, f"F{r}", lib, italic=True)
        calc_cell(ws, f"H{r}", formule, fmt=fmt, bold=True,
                  font_color=BLUE_DARK)

    # ---- Section 4 : conditions et validation --------------------
    section(ws, "B20:H20", "4. CONDITIONS ET VALIDATION")
    quatri = [("État / conditions de l'essai", None),
              ("Observations / anomalies", None),
              ("Vérif. niveau identique ? (Oui/Non)", None),
              ("Validation responsable", None)]
    for i, (lib, fmt) in enumerate(quatri):
        r = 21 + i
        label(ws, f"B{r}", lib)
        input_cell(ws, f"D{r}", fmt=fmt)

    # Validations de données (listes Référentiels)
    validation_liste(ws, "$A$4:$A$6", "D7")    # Site
    validation_liste(ws, "$B$4:$B$6", "D8")    # Équipement
    validation_liste(ws, "$D$4:$D$7", "D9")    # Type de charge
    validation_liste(ws, "$C$4:$C$6", "D10")   # Opérateur
    validation_liste(ws, "$G$4:$G$5", "D23")   # Oui/Non

    # Note champs obligatoires
    note = ws["B27"]
    ws.merge_cells("B27:H27")
    note.value = ("* Champs obligatoires. Les boutons « Valider et ajouter » "
                  "et « Réinitialiser » sont créés par les macros VBA.")
    note.font = Font(size=9, italic=True, color="C00000")
    note.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # Emplacement réservé aux boutons macro (ligne 26)
    b = ws["B26"]
    b.value = "[ Boutons VBA : Valider et ajouter / Réinitialiser ]"
    b.font = Font(size=9, italic=True, color="808080")

    page_setup_landscape(ws)
    return ws


# ================================================================
# FEUILLE 6 — SAISIE_RAPIDE_EAU (formulaire)
# ================================================================

def creer_saisie_rapide_eau(wb):
    """Formulaire de saisie rapide d'un relevé débitmètre."""
    ws = wb.create_sheet("Saisie_Rapide_Eau")
    set_widths(ws, {"A": 2, "B": 40, "C": 3, "D": 28,
                    "E": 3, "F": 40, "G": 3, "H": 28})

    title(ws, "B2:H2", "SAISIE RAPIDE — RELEVÉ DÉBITMÈTRE", size=16, height=42)
    st = ws["B3"]
    ws.merge_cells("B3:H3")
    st.value = "Remplissez les champs puis cliquez sur « Valider et ajouter »."
    st.font = Font(size=10, italic=True, color="808080")
    st.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # ---- Section 1 : identification -----------------------------
    section(ws, "B5:D5", "1. IDENTIFICATION")
    idents = [("Date *", "dd/mm/yyyy"),
              ("Heure", "hh:mm"),
              ("Site *", None),
              ("Débitmètre *", None),
              ("Opérateur", None)]
    for i, (lib, fmt) in enumerate(idents):
        r = 6 + i
        label(ws, f"B{r}", lib)
        input_cell(ws, f"D{r}", fmt=fmt)
        ws.row_dimensions[r].height = 22

    # ---- Section 2 : relevé compteur ----------------------------
    section(ws, "F5:H5", "2. RELEVÉ COMPTEUR")
    compteurs = [("Index début", "0.0"),
                 ("Index fin", "0.0"),
                 ("Unité (m³ / L)", None),
                 ("Durée (h)", "0.00"),
                 ("Énergie utilisée (kWh)", "0.0")]
    for i, (lib, fmt) in enumerate(compteurs):
        r = 6 + i
        label(ws, f"F{r}", lib)
        input_cell(ws, f"H{r}", fmt=fmt)

    # ---- Champs auto-calculés ------------------------------------
    label(ws, "B12", "Volume calculé", italic=True, bold=True)
    calc_cell(ws, "D12", '=IF(OR(H6="",H7=""),"",H7-H6)', fmt="0.0",
              bold=True, font_color=BLUE_DARK)
    label(ws, "F12", "Débit moyen (m³/h)", italic=True, bold=True)
    calc_cell(ws, "H12", '=IFERROR(D12/H9,"")', fmt="0.00",
              bold=True, font_color=BLUE_DARK)

    # ---- Section 3 : observations et validation -------------------
    section(ws, "B15:H15", "3. OBSERVATIONS ET VALIDATION")
    label(ws, "B16", "Source d'énergie")
    input_cell(ws, "D16")
    label(ws, "B17", "Observations / anomalies")
    input_cell(ws, "D17")
    label(ws, "F16", "Validation responsable")
    input_cell(ws, "H16")

    # Validations de données
    validation_liste(ws, "$A$4:$A$6", "D8")    # Site
    validation_liste(ws, "$E$4:$E$6", "D9")    # Débitmètre
    validation_liste(ws, "$F$4:$F$5", "H8")    # Unité
    validation_liste(ws, "$H$4:$H$7", "D16")   # Source d'énergie

    note = ws["B20"]
    ws.merge_cells("B20:H20")
    note.value = ("* Champs obligatoires. Le volume est calculé si les deux "
                  "index sont renseignés.")
    note.font = Font(size=9, italic=True, color="C00000")
    note.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    b = ws["B19"]
    b.value = "[ Boutons VBA : Valider et ajouter / Réinitialiser ]"
    b.font = Font(size=9, italic=True, color="808080")

    page_setup_landscape(ws)
    return ws


# ================================================================
# FEUILLE 7 — SYNTHÈSE_GASOIL
# ================================================================

def creer_synthese_gasoil(wb):
    """Synthèse automatique gasoil : KPI, blocs SUMIF, données graphiques."""
    ws = wb.create_sheet("Synthèse_Gasoil")
    LR = 3 + NB_LIGNES_GASOIL  # 103

    title(ws, "A1:J1", "SYNTHÈSE AUTOMATIQUE — GASOIL", size=14, height=26)
    set_widths(ws, {"A": 30, "B": 3, "C": 16, "D": 3, "E": 14, "F": 14,
                    "G": 3, "H": 10, "I": 14, "J": 16})

    # ---- KPI (A:C, lignes 3 à 11) --------------------------------
    kpis = [
        ("Nombre d'essais",
         f"=COUNT(Saisie_Gasoil!B4:B{LR})", "0"),
        ("Total heures",
         f"=SUM(Saisie_Gasoil!J4:J{LR})", "0.0"),
        ("Total gasoil corrigé (L)",
         f"=SUM(Saisie_Gasoil!P4:P{LR})", "0.0"),
        ("Conso moyenne pondérée (L/h)",
         f"=IFERROR(C5/C4,\"\")", "0.00"),
        ("Énergie totale (kWh)",
         f"=SUM(Saisie_Gasoil!S4:S{LR})", "0.0"),
        ("Conso spécifique globale (L/kWh)",
         f"=IFERROR(C5/C7,\"\")", "0.000"),
        ("Coût total (FCFA)",
         "=ROUND(C5*PrixGasoil,0)", "#,##0"),
        ("Émissions CO₂ estimées (kg)",
         "=ROUND(C5*FacteurCO2,0)", "#,##0"),
        ("Fuite totale (L)",
         f"=SUM(Saisie_Gasoil!O4:O{LR})", "0.0"),
    ]
    for i, (lib, formule, fmt) in enumerate(kpis):
        r = 3 + i
        lc = ws.cell(row=r, column=1, value=lib)
        lc.font = Font(size=10)
        lc.border = thin_border()
        calc_cell(ws, f"C{r}", formule, fmt=fmt, bold=True,
                  font_color=BLUE_DARK)

    # ---- Détail par site (E3:F8) ---------------------------------
    section(ws, "E3:F3", "Détail par site")
    sites = REFERENTIELS["Sites"]
    for i, s in enumerate(sites):
        r = 4 + i
        c = ws.cell(row=r, column=5, value=s)
        c.border = thin_border()
        calc_cell(ws, f"F{r}",
                  f"=SUMIF(Saisie_Gasoil!$C$4:$C${LR},E{r},"
                  f"Saisie_Gasoil!$P$4:$P${LR})", fmt="#,##0.0")

    # ---- Détail par type de charge (E12:F17, pour camembert) -----
    section(ws, "E12:F12", "Détail par type de charge")
    types = REFERENTIELS["Types de charge"]
    for i, t in enumerate(types):
        r = 13 + i
        c = ws.cell(row=r, column=5, value=t)
        c.border = thin_border()
        calc_cell(ws, f"F{r}",
                  f"=SUMIF(Saisie_Gasoil!$E$4:$E${LR},E{r},"
                  f"Saisie_Gasoil!$P$4:$P${LR})", fmt="#,##0.0")

    # ---- Données graphiques : tendance (H3:J33) ------------------
    section(ws, "H3:J3", "Données graphiques — Tendance")
    hdr = ["N° essai", "Conso moyenne (L/h)", "Conso spécifique (L/kWh)"]
    header(ws, 4, 8, hdr)
    for i in range(30):
        r = 5 + i
        src = 4 + i  # ligne source dans Saisie_Gasoil
        calc_cell(ws, f"H{r}", f'=IF(Saisie_Gasoil!A{src}="","",'
                               f'Saisie_Gasoil!A{src})', fmt="0")
        calc_cell(ws, f"I{r}", f'=IF(Saisie_Gasoil!Q{src}="","",'
                               f'Saisie_Gasoil!Q{src})', fmt="0.00")
        calc_cell(ws, f"J{r}", f'=IF(Saisie_Gasoil!T{src}="","",'
                               f'Saisie_Gasoil!T{src})', fmt="0.000")
    return ws


# ================================================================
# FEUILLE 8 — SYNTHÈSE_EAU
# ================================================================

def creer_synthese_eau(wb):
    """Synthèse automatique production d'eau."""
    ws = wb.create_sheet("Synthèse_Eau")
    LR = 3 + NB_LIGNES_EAU  # 203

    title(ws, "A1:I1", "SYNTHÈSE AUTOMATIQUE — PRODUCTION D'EAU",
          size=14, height=26)
    set_widths(ws, {"A": 30, "B": 3, "C": 16, "D": 3, "E": 14, "F": 14,
                    "G": 3, "H": 10, "I": 14})

    kpis = [
        ("Nombre de relevés",
         f"=COUNT(Saisie_Eau!B4:B{LR})", "0"),
        ("Volume total (m³)",
         f"=SUM(Saisie_Eau!I4:I{LR})", "0.0"),
        ("Durée totale (h)",
         f"=SUM(Saisie_Eau!J4:J{LR})", "0.0"),
        ("Débit moyen global (m³/h)",
         '=IFERROR(C4/C5,"")', "0.00"),
        ("Énergie totale (kWh)",
         f"=SUM(Saisie_Eau!L4:L{LR})", "0.0"),
        ("Ratio gasoil / eau (L/m³)",
         f'=IFERROR(Synthèse_Gasoil!C5/C4,"")', "0.00"),
        ("Ratio kWh / m³",
         '=IFERROR(C6/C4,"")', "0.000"),
    ]
    for i, (lib, formule, fmt) in enumerate(kpis):
        r = 3 + i
        lc = ws.cell(row=r, column=1, value=lib)
        lc.font = Font(size=10)
        lc.border = thin_border()
        calc_cell(ws, f"C{r}", formule, fmt=fmt, bold=True,
                  font_color=BLUE_DARK)

    # Détail par site (E3:F8)
    section(ws, "E3:F3", "Détail par site")
    for i, s in enumerate(REFERENTIELS["Sites"]):
        r = 4 + i
        c = ws.cell(row=r, column=5, value=s)
        c.border = thin_border()
        calc_cell(ws, f"F{r}",
                  f"=SUMIF(Saisie_Eau!$D$4:$D${LR},E{r},"
                  f"Saisie_Eau!$I$4:$I${LR})", fmt="#,##0.0")

    # Données graphiques : débit (H3:I33)
    section(ws, "H3:I3", "Données graphiques — Débit")
    header(ws, 4, 8, ["N° relevé", "Débit moyen (m³/h)"])
    for i in range(30):
        r = 5 + i
        src = 4 + i
        calc_cell(ws, f"H{r}", f'=IF(Saisie_Eau!A{src}="","",'
                               f'Saisie_Eau!A{src})', fmt="0")
        calc_cell(ws, f"I{r}", f'=IF(Saisie_Eau!K{src}="","",'
                               f'Saisie_Eau!K{src})', fmt="0.00")
    return ws


# ================================================================
# FEUILLE 9 — TABLEAU_BORD (graphiques pro)
# ================================================================

def creer_tableau_bord(wb):
    """Tableau de bord : cartes KPI + 5 graphiques sans chevauchement."""
    ws = wb.create_sheet("Tableau_Bord")
    set_widths(ws, {"A": 4, "B": 22, "C": 22, "D": 22, "E": 22,
                    "F": 22, "G": 22, "H": 22})

    title(ws, "B2:H2", "TABLEAU DE BORD — GASOIL & PRODUCTION D'EAU",
          size=16, height=42)
    st = ws["B3"]
    ws.merge_cells("B3:H3")
    st.value = ("Vue synthétique de la consommation de gasoil et de la "
                "production d'eau — mise à jour automatique.")
    st.font = Font(size=10, italic=True, color="808080")
    st.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # ---- Cartes KPI (ligne 5 : en-têtes, ligne 6 : valeurs) -------
    cartes = [
        ("B", "Gasoil total (L)", "=Synthèse_Gasoil!C5", "#,##0"),
        ("C", "Conso moyenne (L/h)", "=Synthèse_Gasoil!C6", "0.00"),
        ("D", "Conso spécifique (L/kWh)", "=Synthèse_Gasoil!C8", "0.000"),
        ("E", "Coût total (FCFA)", "=Synthèse_Gasoil!C9", "#,##0"),
        ("F", "Volume eau (m³)", "=Synthèse_Eau!C4", "#,##0"),
        ("G", "Débit moyen (m³/h)", "=Synthèse_Eau!C6", "0.00"),
        ("H", "L gasoil / m³ eau", "=Synthèse_Eau!C8", "0.00"),
    ]
    for col, lib, formule, fmt in cartes:
        h = ws[f"{col}5"]
        h.value = lib
        h.font = Font(bold=True, color="FFFFFF", size=10)
        h.fill = fill(BLUE_MED)
        h.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        h.border = thin_border()
        v = ws[f"{col}6"]
        v.value = formule
        v.number_format = fmt
        v.font = Font(bold=True, color=BLUE_DARK, size=16)
        v.fill = fill(BLUE_PALE)
        v.alignment = Alignment(horizontal="center", vertical="center")
        v.border = thin_border()
    ws.row_dimensions[5].height = 30
    ws.row_dimensions[6].height = 38

    # ---- Section analyse graphique --------------------------------
    section(ws, "B9:H9", "ANALYSE GRAPHIQUE")

    # ================= GRAPHIQUE 1 : courbe L/h =====================
    ch1 = LineChart()
    ch1.title = "Évolution de la consommation moyenne (L/h)"
    ch1.style = 2
    ch1.height, ch1.width = 8, 16
    data = Reference(wb["Synthèse_Gasoil"], min_col=9, min_row=4, max_row=34)
    cats = Reference(wb["Synthèse_Gasoil"], min_col=8, min_row=5, max_row=34)
    ch1.add_data(data, titles_from_data=True)
    ch1.set_categories(cats)
    s1 = ch1.series[0]
    s1.smooth = True
    s1.graphicalProperties = GraphicalProperties(ln=LineProperties(
        solidFill=CHART_COLORS[0], w=28000))
    s1.marker = Marker(symbol="circle", size=6, spPr=GraphicalProperties(
        solidFill=CHART_COLORS[1]))
    ch1.x_axis.title = "N° d'essai"
    ch1.y_axis.title = "L/h"
    ch1.x_axis.delete = False
    ch1.y_axis.delete = False
    ch1.legend = None
    ws.add_chart(ch1, "B11")

    # ================= GRAPHIQUE 2 : barres gasoil/site =============
    ch2 = BarChart()
    ch2.type = "col"
    ch2.title = "Gasoil total par site (L)"
    ch2.style = 2
    ch2.height, ch2.width = 8, 14
    ch2.gapWidth = 60
    data2 = Reference(wb["Synthèse_Gasoil"], min_col=6, min_row=4, max_row=8)
    cats2 = Reference(wb["Synthèse_Gasoil"], min_col=5, min_row=5, max_row=8)
    ch2.add_data(data2, titles_from_data=False)
    ch2.set_categories(cats2)
    s2 = ch2.series[0]
    s2.graphicalProperties = GraphicalProperties(
        solidFill=CHART_COLORS[1], ln=LineProperties(solidFill=CHART_COLORS[0]))
    ch2.dataLabels = DataLabelList(showVal=True, numFmt="#,##0",
                                   dLblPos="outEnd")
    ch2.x_axis.title = "Site"
    ch2.y_axis.title = "Litres"
    ch2.x_axis.delete = False
    ch2.y_axis.delete = False
    ch2.legend = None
    ws.add_chart(ch2, "D11")

    # ================= GRAPHIQUE 3 : barres eau/site =================
    ch3 = BarChart()
    ch3.type = "col"
    ch3.title = "Volume d'eau produit par site (m³)"
    ch3.style = 2
    ch3.height, ch3.width = 8, 14
    ch3.gapWidth = 60
    data3 = Reference(wb["Synthèse_Eau"], min_col=6, min_row=4, max_row=8)
    cats3 = Reference(wb["Synthèse_Eau"], min_col=5, min_row=5, max_row=8)
    ch3.add_data(data3, titles_from_data=False)
    ch3.set_categories(cats3)
    s3 = ch3.series[0]
    s3.graphicalProperties = GraphicalProperties(
        solidFill="70AD47", ln=LineProperties(solidFill=DARK_GREEN))
    ch3.dataLabels = DataLabelList(showVal=True, numFmt="#,##0.0",
                                   dLblPos="outEnd")
    ch3.x_axis.title = "Site"
    ch3.y_axis.title = "m³"
    ch3.x_axis.delete = False
    ch3.y_axis.delete = False
    ch3.legend = None
    ws.add_chart(ch3, "F11")

    # ================= GRAPHIQUE 4 : camembert types de charge =======
    ch4 = PieChart()
    ch4.title = "Répartition gasoil par type de charge"
    ch4.style = 2
    ch4.height, ch4.width = 8, 14
    data4 = Reference(wb["Synthèse_Gasoil"], min_col=6, min_row=13,
                      max_row=16)
    cats4 = Reference(wb["Synthèse_Gasoil"], min_col=5, min_row=13,
                      max_row=16)
    ch4.add_data(data4, titles_from_data=False)
    ch4.set_categories(cats4)
    # Couleurs personnalisées par point de données
    pts = []
    for idx, col_hex in enumerate(CHART_COLORS[:4]):
        dp = DataPoint(idx=idx)
        dp.graphicalProperties = GraphicalProperties(solidFill=col_hex)
        pts.append(dp)
    ch4.series[0].data_points = pts
    ch4.dataLabels = DataLabelList(showPercent=True)
    ch4.legend.position = "r"
    ws.add_chart(ch4, "B27")

    # ================= GRAPHIQUE 5 : courbe débit eau ================
    ch5 = LineChart()
    ch5.title = "Évolution du débit d'eau (m³/h)"
    ch5.style = 2
    ch5.height, ch5.width = 8, 16
    data5 = Reference(wb["Synthèse_Eau"], min_col=9, min_row=4, max_row=34)
    cats5 = Reference(wb["Synthèse_Eau"], min_col=8, min_row=5, max_row=34)
    ch5.add_data(data5, titles_from_data=True)
    ch5.set_categories(cats5)
    s5 = ch5.series[0]
    s5.smooth = True
    s5.graphicalProperties = GraphicalProperties(ln=LineProperties(
        solidFill="70AD47", w=28000))
    s5.marker = Marker(symbol="diamond", size=7, spPr=GraphicalProperties(
        solidFill=DARK_GREEN))
    ch5.x_axis.title = "N° de relevé"
    ch5.y_axis.title = "m³/h"
    ch5.x_axis.delete = False
    ch5.y_axis.delete = False
    ch5.legend = None
    ws.add_chart(ch5, "E27")

    # Réserve pour les boutons de navigation macro (ligne 45+)
    nav = ws["B45"]
    nav.value = ("[ Boutons VBA : Saisie rapide gasoil / Saisie rapide eau / "
                 "Exporter en PDF ]")
    nav.font = Font(size=9, italic=True, color="808080")

    page_setup_landscape(ws)
    return ws


# ================================================================
# FEUILLE 10 — FICHE_TERRAIN
# ================================================================

def creer_fiche_terrain(wb):
    """Fiche terrain imprimable pour la méthode plein à plein + eau."""
    ws = wb.create_sheet("Fiche_Terrain")
    set_widths(ws, {"A": 26, "B": 20, "C": 4, "D": 26, "E": 20})

    title(ws, "A1:E1", "FICHE TERRAIN — MESURE PLEIN À PLEIN",
          size=14, height=28)

    # ---- Identification -------------------------------------------
    section(ws, "A3:E3", "1. CONTEXTE DE L'ESSAI")
    champs = [("Site", "A4", "B4"), ("Groupe / code", "A5", "B5"),
              ("Date", "D4", "E4"), ("Opérateur", "D5", "E5"),
              ("Puissance (kVA)", "A6", "B6"),
              ("Charge / équipement", "D6", "E6")]
    for lib, lab_coord, inp_coord in champs:
        label(ws, lab_coord, lib)
        input_cell(ws, inp_coord,
                   fmt="dd/mm/yyyy" if lib == "Date" else None)

    # ---- Compteurs --------------------------------------------------
    section(ws, "A8:E8", "2. COMPTEUR HORAIRE")
    label(ws, "A9", "Compteur début (h)")
    input_cell(ws, "B9", fmt="0.0")
    label(ws, "D9", "Compteur fin (h)")
    input_cell(ws, "E9", fmt="0.0")
    label(ws, "A10", "Durée (h) — auto")
    calc_cell(ws, "B10", '=IF(OR(B9="",E9=""),"",E9-B9)', fmt="0.00")

    # ---- Carburant ---------------------------------------------------
    section(ws, "A12:E12", "3. CARBURANT")
    label(ws, "A13", "Gasoil ajouté (L)")
    input_cell(ws, "B13", fmt="0.0")
    label(ws, "D13", "Énergie produite (kWh)")
    input_cell(ws, "E13", fmt="0.0")
    label(ws, "A14", "Fuites / autres (L)")
    input_cell(ws, "B14", fmt="0.0")
    label(ws, "A15", "Conso moyenne (L/h) — auto")
    calc_cell(ws, "B15", '=IFERROR((B13+B14)/B10,"")', fmt="0.00")

    # ---- Production d'eau --------------------------------------------
    section(ws, "A17:E17", "SUIVI DE LA PRODUCTION D'EAU — DÉBITMÈTRES")
    header(ws, 18, 1, ["Débitmètre", "Index début", "Index fin",
                       "Durée (h)", "Volume (m³)"])
    for r in range(19, 24):  # 5 lignes
        for col in ["A", "B", "C", "D"]:
            input_cell(ws, f"{col}{r}",
                       fmt="0.0" if col in ("B", "C", "D") else None)
        calc_cell(ws, f"E{r}",
                  f'=IF(OR(B{r}="",C{r}=""),"",C{r}-B{r})', fmt="0.0")

    # ---- Observations ---------------------------------------------------
    section(ws, "A25:E25", "OBSERVATIONS")
    ws.merge_cells("A26:E30")
    obs = ws["A26"]
    obs.fill = fill(INPUT_FILL)
    for row_cells in ws["A26:E30"]:
        for cc in row_cells:
            cc.border = thin_border()
            cc.fill = fill(INPUT_FILL)
    obs.alignment = Alignment(horizontal="left", vertical="top",
                              wrap_text=True)

    # ---- Signatures ---------------------------------------------------------
    section(ws, "A32:E32", "SIGNATURES")
    label(ws, "A33", "Opérateur :")
    label(ws, "D33", "Responsable :")
    for coord in ("B34", "E34"):
        ws[coord].border = Border(bottom=Side(style="thin", color="000000"))

    page_setup_landscape(ws)
    return ws


# ================================================================
# FEUILLE 11 — MODE_EMPLOI
# ================================================================

MODE_EMPLOI_ETAPES = [
    ("Préparer l'essai",
     "Réaliser le plein complet du réservoir (référence = niveau « plein ») "
     "et noter le compteur horaire de départ."),
    ("Exploiter le groupe",
     "Faire tourner le groupe électrogène dans les conditions réelles "
     "(charge, démarrage/arrêt) pendant la durée souhaitée."),
     ("Consigner chaque ajout",
      "Si du gasoil est ajouté en cours d'essai, le noter immédiatement "
      "dans « Autres ajouts » avec l'heure."),
    ("Refaire le plein (plein à plein)",
     "À la fin de l'essai, attendre 10 minutes, remettre le réservoir au "
     "même repère « plein » : le litre ajouté = consommation réelle."),
    ("Saisir dans Saisie_Gasoil",
     "Renseigner la ligne correspondante dans l'onglet Saisie_Gasoil : "
     "durée, conso corrigée, conso moyenne et spécifique sont calculées."),
    ("Utiliser le formulaire Saisie_Rapide_Gasoil",
     "Pour aller plus vite, remplir le formulaire de l'onglet "
     "« Saisie_Rapide_Gasoil » puis cliquer sur le bouton « Valider et "
     "ajouter » (macro) : la ligne est ajoutée automatiquement au tableau."),
    ("Relevés débitmètres",
     "Relever index début/fin et durée dans Saisie_Eau (ou via le "
     "formulaire Saisie_Rapide_Eau) : volume et débit moyen sont calculés."),
     ("Consulter le Tableau_Bord",
      "Les KPI et les 5 graphiques se mettent à jour automatiquement ; "
      "aucune action n'est requise."),
    ("Sécurité",
     "Manipuler le gasoil avec des gants, loin de toute flamme ; stocker "
     "les bidons dans un lieu ventilé ; ne jamais transvaser en présence "
     "d'enfants."),
    ("Sauvegarde mensuelle",
     "En fin de mois, enregistrer une copie datée du classeur "
     "(ex. suivi_2026-10.xlsm) et archiver la Fiche_Terrain imprimée."),
]


def creer_mode_emploi(wb):
    """Feuille Mode_Emploi : 10 étapes numérotées d'utilisation."""
    ws = wb.create_sheet("Mode_Emploi")
    set_widths(ws, {"A": 6, "B": 34, "C": 100})

    title(ws, "A1:C1", "MODE D'EMPLOI — GUIDE EN 10 ÉTAPES",
          size=14, height=28)

    for i, (titre, texte) in enumerate(MODE_EMPLOI_ETAPES, start=1):
        r = 3 + i * 2
        n = ws.cell(row=r, column=1, value=i)
        n.font = Font(bold=True, color="FFFFFF", size=12)
        n.fill = fill(BLUE_MED)
        n.alignment = Alignment(horizontal="center", vertical="center")
        n.border = thin_border()
        t = ws.cell(row=r, column=2, value=titre)
        t.font = Font(bold=True, color=BLUE_DARK, size=10)
        t.alignment = Alignment(vertical="center", wrap_text=True)
        t.border = thin_border()
        d = ws.cell(row=r, column=3, value=texte)
        d.font = Font(size=10)
        d.alignment = Alignment(vertical="center", wrap_text=True)
        d.border = thin_border()
        ws.row_dimensions[r].height = 34

    r_fin = 3 + len(MODE_EMPLOI_ETAPES) * 2 + 2
    note = ws.cell(row=r_fin, column=2,
                   value=("Astuce : la feuille « Saisie_Rapide_Gasoil » permet "
                          "d'enregistrer un essai en moins de 30 secondes via "
                          "les boutons macro."))
    note.font = Font(size=9, italic=True, color="808080")
    ws.merge_cells(start_row=r_fin, start_column=2,
                   end_row=r_fin, end_column=3)
    return ws


# ================================================================
# NOMS DÉFINIS
# ================================================================

def creer_noms_definis(wb):
    """Noms définis utilisés par les formules de synthèse."""
    wb.defined_names.add(DefinedName(
        "PrixGasoil", attr_text="Paramètres!$D$7"))
    wb.defined_names.add(DefinedName(
        "FacteurCO2", attr_text="Paramètres!$D$8"))


# ================================================================
# MODULES VBA
# ================================================================

VBA_MOD_ESSAIS = r'''
Attribute VB_Name = "modEssais"
Option Explicit

'==============================================================
' Formulaire Saisie_Rapide_Gasoil -> tableau tblGasoil
'==============================================================
Public Sub ValiderSaisieGasoil()
    Dim wsF As Worksheet, wsT As Worksheet
    Dim lo As ListObject, lr As ListRow
    Set wsF = ThisWorkbook.Worksheets("Saisie_Rapide_Gasoil")
    Set wsT = ThisWorkbook.Worksheets("Saisie_Gasoil")
    Set lo = wsT.ListObjects("tblGasoil")

    '--- Controle des champs obligatoires ---
    If IsEmpty(wsF.Range("D6").Value) Or wsF.Range("D6").Value = "" Then
        MsgBox "La date de l'essai est obligatoire.", vbExclamation: wsF.Range("D6").Select: Exit Sub
    End If
    If IsEmpty(wsF.Range("D7").Value) Or wsF.Range("D7").Value = "" Then
        MsgBox "Le site est obligatoire.", vbExclamation: wsF.Range("D7").Select: Exit Sub
    End If
    If IsEmpty(wsF.Range("D8").Value) Or wsF.Range("D8").Value = "" Then
        MsgBox "Le code equipement est obligatoire.", vbExclamation: wsF.Range("D8").Select: Exit Sub
    End If
    If IsEmpty(wsF.Range("D15").Value) Or wsF.Range("D15").Value = "" Then
        MsgBox "Le gasoil ajoute est obligatoire.", vbExclamation: wsF.Range("D15").Select: Exit Sub
    End If

    Application.ScreenUpdating = False

    '--- Ajout d'une ligne dans le tableau structure ---
    Set lr = lo.ListRows.Add
    With lr.Range
        .Cells(1, 2).Value = CDate(wsF.Range("D6").Value)          ' Date
        .Cells(1, 3).Value = wsF.Range("D7").Value                 ' Site
        .Cells(1, 4).Value = wsF.Range("D8").Value                 ' Equipement
        .Cells(1, 5).Value = wsF.Range("D9").Value                 ' Type de charge
        .Cells(1, 6).Value = wsF.Range("H6").Value                 ' Heure debut
        .Cells(1, 7).Value = wsF.Range("H7").Value                 ' Heure fin
        .Cells(1, 8).Value = wsF.Range("H8").Value                 ' Compteur debut
        .Cells(1, 9).Value = wsF.Range("H9").Value                 ' Compteur fin
        .Cells(1, 11).Value = wsF.Range("D13").Value               ' Niveau initial
        .Cells(1, 12).Value = wsF.Range("D14").Value               ' Niveau final
        .Cells(1, 13).Value = wsF.Range("D15").Value               ' Gasoil ajoute
        .Cells(1, 14).Value = wsF.Range("D16").Value               ' Autres ajouts
        .Cells(1, 15).Value = wsF.Range("D17").Value               ' Fuites
        .Cells(1, 18).Value = wsF.Range("H10").Value               ' Charge estimee
        .Cells(1, 19).Value = wsF.Range("D18").Value               ' Energie produite
        .Cells(1, 21).Value = wsF.Range("D21").Value               ' Etat / conditions
        .Cells(1, 22).Value = wsF.Range("D10").Value               ' Operateur
        .Cells(1, 23).Value = wsF.Range("D22").Value               ' Observations
        .Cells(1, 24).Value = wsF.Range("D23").Value               ' Verif. niveau
        .Cells(1, 25).Value = wsF.Range("D24").Value               ' Validation
    End With

    Application.ScreenUpdating = True
    MsgBox "Essai enregistre dans Saisie_Gasoil.", vbInformation
    Call RazSaisieGasoil
End Sub

Public Sub RazSaisieGasoil()
    Dim wsF As Worksheet
    Set wsF = ThisWorkbook.Worksheets("Saisie_Rapide_Gasoil")
    wsF.Range("D6:D10").ClearContents
    wsF.Range("H6:H10").ClearContents
    wsF.Range("D13:D18").ClearContents
    wsF.Range("D21:D24").ClearContents
    wsF.Range("D6").Select
End Sub

'==============================================================
' Formulaire Saisie_Rapide_Eau -> tableau tblEau
'==============================================================
Public Sub ValiderSaisieEau()
    Dim wsF As Worksheet, wsT As Worksheet
    Dim lo As ListObject, lr As ListRow
    Set wsF = ThisWorkbook.Worksheets("Saisie_Rapide_Eau")
    Set wsT = ThisWorkbook.Worksheets("Saisie_Eau")
    Set lo = wsT.ListObjects("tblEau")

    If IsEmpty(wsF.Range("D6").Value) Or wsF.Range("D6").Value = "" Then
        MsgBox "La date est obligatoire.", vbExclamation: wsF.Range("D6").Select: Exit Sub
    End If
    If IsEmpty(wsF.Range("D8").Value) Or wsF.Range("D8").Value = "" Then
        MsgBox "Le site est obligatoire.", vbExclamation: wsF.Range("D8").Select: Exit Sub
    End If
    If IsEmpty(wsF.Range("D9").Value) Or wsF.Range("D9").Value = "" Then
        MsgBox "Le debitmetre est obligatoire.", vbExclamation: wsF.Range("D9").Select: Exit Sub
    End If

    Application.ScreenUpdating = False
    Set lr = lo.ListRows.Add
    With lr.Range
        .Cells(1, 2).Value = CDate(wsF.Range("D6").Value)   ' Date
        .Cells(1, 3).Value = wsF.Range("D7").Value          ' Heure
        .Cells(1, 4).Value = wsF.Range("D8").Value          ' Site
        .Cells(1, 5).Value = wsF.Range("D9").Value          ' Debitometre
        .Cells(1, 6).Value = wsF.Range("H6").Value          ' Index debut
        .Cells(1, 7).Value = wsF.Range("H7").Value          ' Index fin
        .Cells(1, 8).Value = wsF.Range("H8").Value          ' Unite
        .Cells(1, 10).Value = wsF.Range("H9").Value         ' Duree
        .Cells(1, 12).Value = wsF.Range("H10").Value        ' Energie utilisee
        .Cells(1, 13).Value = wsF.Range("D16").Value        ' Source d'energie
        .Cells(1, 14).Value = wsF.Range("D17").Value        ' Observations
    End With
    Application.ScreenUpdating = True
    MsgBox "Releve enregistre dans Saisie_Eau.", vbInformation
    Call RazSaisieEau
End Sub

Public Sub RazSaisieEau()
    Dim wsF As Worksheet
    Set wsF = ThisWorkbook.Worksheets("Saisie_Rapide_Eau")
    wsF.Range("D6:D10").ClearContents
    wsF.Range("H6:H10").ClearContents
    wsF.Range("D16:D17").ClearContents
    wsF.Range("H16").ClearContents
    wsF.Range("D6").Select
End Sub

'==============================================================
' Ajout de lignes vides avec la date du jour
'==============================================================
Public Sub AjouterLigneGasoilVide()
    Dim lo As ListObject, lr As ListRow
    Set lo = ThisWorkbook.Worksheets("Saisie_Gasoil").ListObjects("tblGasoil")
    Set lr = lo.ListRows.Add
    lr.Range.Cells(1, 2).Value = Date
    lr.Range.Cells(1, 2).NumberFormat = "dd/mm/yyyy"
    MsgBox "Ligne vide ajoutee (date du jour).", vbInformation
End Sub

Public Sub AjouterLigneEauVide()
    Dim lo As ListObject, lr As ListRow
    Set lo = ThisWorkbook.Worksheets("Saisie_Eau").ListObjects("tblEau")
    Set lr = lo.ListRows.Add
    lr.Range.Cells(1, 2).Value = Date
    lr.Range.Cells(1, 2).NumberFormat = "dd/mm/yyyy"
    MsgBox "Ligne vide ajoutee (date du jour).", vbInformation
End Sub
'''

VBA_MOD_EXPORT = r'''
Attribute VB_Name = "modExport"
Option Explicit

Private Function DossierClasseur() As String
    DossierClasseur = ThisWorkbook.Path & Application.PathSeparator
End Function

'--- Export du tableau de bord en PDF date ---
Public Sub ExporterTableauBordPDF()
    Dim nom As String
    nom = DossierClasseur & "Tableau_de_Bord_" & Format(Date, "yyyymm-dd") & ".pdf"
    ThisWorkbook.Worksheets("Tableau_Bord").ExportAsFixedFormat _
        Type:=xlTypePDF, Filename:=nom, Quality:=xlQualityStandard
    MsgBox "Export realise :" & vbCrLf & nom, vbInformation
End Sub

'--- Export de la fiche terrain en PDF date ---
Public Sub ExporterFicheTerrainPDF()
    Dim nom As String
    nom = DossierClasseur & "Fiche_Terrain_" & Format(Date, "yyyymm-dd") & ".pdf"
    ThisWorkbook.Worksheets("Fiche_Terrain").ExportAsFixedFormat _
        Type:=xlTypePDF, Filename:=nom, Quality:=xlQualityStandard
    MsgBox "Export realise :" & vbCrLf & nom, vbInformation
End Sub

'--- Export complet du classeur en PDF ---
Public Sub ExporterClasseurCompletPDF()
    Dim nom As String
    nom = DossierClasseur & "Classeur_Complet_" & Format(Date, "yyyymm-dd") & ".pdf"
    ThisWorkbook.ExportAsFixedFormat _
        Type:=xlTypePDF, Filename:=nom, Quality:=xlQualityStandard
    MsgBox "Export realise :" & vbCrLf & nom, vbInformation
End Sub

'--- Recalcul complet ---
Public Sub ActualiserTout()
    Application.CalculateFull
    MsgBox "Calculs actualises.", vbInformation
End Sub
'''

VBA_MOD_NAVIGATION = r'''
Attribute VB_Name = "modNavigation"
Option Explicit

Private Sub Aller(nomFeuille As String)
    ThisWorkbook.Worksheets(nomFeuille).Activate
End Sub

Public Sub AllerTableauBord()
    Aller "Tableau_Bord"
End Sub

Public Sub AllerSaisieGasoil()
    Aller "Saisie_Gasoil"
End Sub

Public Sub AllerSaisieEau()
    Aller "Saisie_Eau"
End Sub

Public Sub AllerSaisieRapideGasoil()
    Aller "Saisie_Rapide_Gasoil"
End Sub

Public Sub AllerSaisieRapideEau()
    Aller "Saisie_Rapide_Eau"
End Sub

Public Sub AllerFicheTerrain()
    Aller "Fiche_Terrain"
End Sub
'''

VBA_THISWORKBOOK = r'''
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    ThisWorkbook.Worksheets("Tableau_Bord").Activate
    CreerBoutonsSaisieRapide
End Sub

Public Sub CreerBoutonsSaisieRapide()
    Dim ws As Worksheet, b As Button, rng As Range

    '--- Suppression des anciens boutons (eviter les doublons) ---
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        For Each b In ws.Buttons
            If Left(b.Name, 4) = "btn_" Then b.Delete
        Next b
        On Error GoTo 0
    Next ws

    '--- Saisie_Rapide_Gasoil : B26 et D26 ---
    Set ws = ThisWorkbook.Worksheets("Saisie_Rapide_Gasoil")
    Set rng = ws.Range("B26")
    Set b = ws.Buttons.Add(rng.Left, rng.Top, 180, 42)
    b.Name = "btn_ValiderGasoil": b.Text = "Valider et ajouter"
    b.OnAction = "ValiderSaisieGasoil"
    Set rng = ws.Range("D26")
    Set b = ws.Buttons.Add(rng.Left, rng.Top, 180, 42)
    b.Name = "btn_RazGasoil": b.Text = "Reinitialiser"
    b.OnAction = "RazSaisieGasoil"

    '--- Saisie_Rapide_Eau : B19 et D19 ---
    Set ws = ThisWorkbook.Worksheets("Saisie_Rapide_Eau")
    Set rng = ws.Range("B19")
    Set b = ws.Buttons.Add(rng.Left, rng.Top, 180, 42)
    b.Name = "btn_ValiderEau": b.Text = "Valider et ajouter"
    b.OnAction = "ValiderSaisieEau"
    Set rng = ws.Range("D19")
    Set b = ws.Buttons.Add(rng.Left, rng.Top, 180, 42)
    b.Name = "btn_RazEau": b.Text = "Reinitialiser"
    b.OnAction = "RazSaisieEau"

    '--- Tableau_Bord : navigation en B45, D45, F45 ---
    Set ws = ThisWorkbook.Worksheets("Tableau_Bord")
    Set rng = ws.Range("B45")
    Set b = ws.Buttons.Add(rng.Left, rng.Top, 180, 42)
    b.Name = "btn_NavGasoil": b.Text = "Saisie rapide gasoil"
    b.OnAction = "AllerSaisieRapideGasoil"
    Set rng = ws.Range("D45")
    Set b = ws.Buttons.Add(rng.Left, rng.Top, 180, 42)
    b.Name = "btn_NavEau": b.Text = "Saisie rapide eau"
    b.OnAction = "AllerSaisieRapideEau"
    Set rng = ws.Range("F45")
    Set b = ws.Buttons.Add(rng.Left, rng.Top, 180, 42)
    b.Name = "btn_ExportPDF": b.Text = "Exporter en PDF"
    b.OnAction = "ExporterTableauBordPDF"
End Sub
'''


def generer_modules_vba():
    """Écrit les 4 modules VBA dans le dossier ./vba/ et renvoie la liste."""
    os.makedirs(DOSSIER_VBA, exist_ok=True)
    fichiers = {
        "modEssais.bas": VBA_MOD_ESSAIS.strip("\n"),
        "modExport.bas": VBA_MOD_EXPORT.strip("\n"),
        "modNavigation.bas": VBA_MOD_NAVIGATION.strip("\n"),
        "ThisWorkbook.cls": VBA_THISWORKBOOK.strip("\n"),
    }
    chemins = []
    for nom, contenu in fichiers.items():
        p = os.path.join(DOSSIER_VBA, nom)
        with open(p, "w", encoding="utf-8-sig") as f:
            f.write(contenu)
        chemins.append(os.path.abspath(p))
        print(f"  [VBA] module généré : {p}")
    return chemins


# ================================================================
# INJECTION DES MACROS VIA PYWIN32
# ================================================================

def injecter_macros_vba(chemin_xlsx, chemin_xlsm, fichiers_vba):
    """Tente l'injection VBA via COM. Renvoie True si succès."""
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False, ("pywin32 indisponible (Windows + Excel requis).")

    excel = None
    wb = None
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(chemin_xlsx)
        vbc = wb.VBProject.VBComponents
        # ThisWorkbook : le module existe déjà -> on colle son code
        # (un .cls ne peut pas être importé deux fois sous ce nom).
        for f in fichiers_vba:
            if os.path.basename(f).lower() == "thisworkbook.cls":
                with open(f, "r", encoding="utf-8-sig") as fh:
                    code = "\n".join(
                        ligne for ligne in fh.read().splitlines()
                        if not ligne.lower().startswith("attribute "))
                vbc.Item("ThisWorkbook").CodeModule.AddFromString(code)
            else:
                vbc.Import(f)
        # Enregistrement au format .xlsm (52 = xlOpenXMLWorkbookMacroEnabled)
        wb.SaveAs(chemin_xlsm, FileFormat=52)
        wb.Close(SaveChanges=False)
        excel.Quit()
        return True, "Macros injectées automatiquement."
    except Exception as e:
        # Nettoyage robuste en cas d'échec
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        return False, f"Échec de l'injection ({type(e).__name__}: {e})."


def afficher_instructions_manuelles():
    """Instructions d'import manuel des modules VBA."""
    print("\n" + "=" * 64)
    print("IMPORT MANUEL DES MACROS (si l'injection a échoué)")
    print("=" * 64)
    print("""\
1. Ouvrir le fichier "%s" dans Excel.
2. Alt + F11 pour lancer l'éditeur VBA.
3. Menu Fichier > Importer un fichier... et importer, depuis
   le dossier "%s/" :
      - modEssais.bas
      - modExport.bas
      - modNavigation.bas
4. Double-cliquer sur "ThisWorkbook" dans l'explorateur de
   projets, puis coller le contenu du fichier "ThisWorkbook.cls"
   (remplacer le code existant s'il y en a).
   Astuce : il est aussi possible d'importer ThisWorkbook.cls
   directement ; vérifier alors qu'il ne reste qu'un seul
   objet ThisWorkbook.
5. Enregistrer sous > type "Classeur Excel avec prise en
   charge des macros (*.xlsm)".
6. Au premier ouverture : cliquer sur "Activer les macros".

NB : le classeur reste pleinement utilisable sans les macros
(formules, validations, graphiques fonctionnent en standalone).
""" % (FICHIER_XLSM, DOSSIER_VBA))


# ================================================================
# CONSTRUCTION DU CLASSEUR
# ================================================================

def construire_classeur():
    """Crée toutes les feuilles dans l'ordre demandé."""
    wb = Workbook()
    creer_parametres(wb)          # active -> "Paramètres"
    creer_referentiels(wb)
    creer_saisie_gasoil(wb)
    creer_saisie_eau(wb)
    creer_saisie_rapide_gasoil(wb)
    creer_saisie_rapide_eau(wb)
    creer_synthese_gasoil(wb)
    creer_synthese_eau(wb)
    creer_tableau_bord(wb)
    creer_fiche_terrain(wb)
    creer_mode_emploi(wb)
    creer_noms_definis(wb)

    # Feuille active à l'ouverture : Tableau_Bord
    wb.active = wb.sheetnames.index("Tableau_Bord")
    return wb


# ================================================================
# PROGRAMME PRINCIPAL
# ================================================================

def main():
    print("=" * 64)
    print("Génération du classeur de suivi gasoil & production d'eau v3")
    print("=" * 64)

    # 1. Construction du classeur
    print("\n[1/4] Construction des feuilles...")
    wb = construire_classeur()
    print(f"  {len(wb.sheetnames)} feuilles : {wb.sheetnames}")

    # Sauvegarde intermédiaire .xlsx (sans macros, toujours ouvrable)
    wb.save(FICHIER_SANS_MACROS)
    print(f"  Classeur écrit : {FICHIER_SANS_MACROS}")

    # 2. Génération des modules VBA dans ./vba/
    print("\n[2/4] Génération des modules VBA...")
    fichiers_vba = generer_modules_vba()

    # 3. Tentative d'injection automatique des macros
    print("\n[3/4] Tentative d'injection automatique (pywin32)...")
    chemin_xlsx = os.path.abspath(FICHIER_SANS_MACROS)
    chemin_xlsm = os.path.abspath(FICHIER_XLSM)
    ok, msg = injecter_macros_vba(chemin_xlsx, chemin_xlsm, fichiers_vba)
    if ok:
        print(f"  SUCCÈS : {msg}")
        # Le .xlsx intermédiaire n'est plus nécessaire
        try:
            os.remove(FICHIER_SANS_MACROS)
        except OSError:
            pass
    else:
        print(f"  AVERTISSEMENT : {msg}")
        # Sans pywin32 : on produit malgré tout un .xlsm valide
        # (openpyxl conserve les tables/formules ; les macros seront
        #  importées manuellement selon les instructions ci-dessous).
        try:
            wb.save(FICHIER_XLSM)
            os.remove(FICHIER_SANS_MACROS)
            print("  Fallback : .xlsm écrit sans macros "
                  "(ouvrable, import manuel possible).")
        except Exception:
            print(traceback.format_exc())
        afficher_instructions_manuelles()

    # 4. Chemin absolu du livrable final
    print("\n[4/4] Livrable final :")
    final = chemin_xlsm if os.path.exists(chemin_xlsm) \
        else os.path.abspath(FICHIER_SANS_MACROS)
    print("-" * 64)
    print(f"CHEMIN ABSOLU : {final}")
    print("-" * 64)
    print("Terminé. Modules VBA disponibles dans ./vba/.")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        print("ERREUR FATALE pendant la génération :")
        print(traceback.format_exc())
        sys.exit(1)
