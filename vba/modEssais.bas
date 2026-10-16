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