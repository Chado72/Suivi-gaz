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