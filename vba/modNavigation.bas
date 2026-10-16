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