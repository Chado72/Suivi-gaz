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