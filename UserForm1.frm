VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Formulaire :"
   ClientHeight    =   3840
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6870
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub UserForm_Initialize()
' Remplit la ComboBox avec les départements
    With ComboBox1
        .AddItem "75"
        .AddItem "77"
        .AddItem "78"
        .AddItem "91"
        .AddItem "92"
        .AddItem "93"
        .AddItem "94"
        .AddItem "95"
    End With
    
        
        ' Remplit la ComboBox2 avec une autre liste si besoin (par exemple : nombre de pièces ou autre critère)
    With ComboBox2
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .AddItem "6"
        .AddItem "7"
        .AddItem "8"
 
    End With
    
     With ComboBox3
        .AddItem "Appartement"
        .AddItem "Maison"
 
    End With
End Sub


Sub CommandButton1_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Feuil1") ' Feuille avec la base DVF

    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    
    ' Récupération des critères depuis le formulaire
    Dim surfaceSouhaitee As Double: surfaceSouhaitee = (TextBox1.Value)
    Dim dep As String: dep = (ComboBox1.Value)
    Dim nbPieces As Integer: nbPieces = (ComboBox2.Value)
    Dim logement As String: logement = (ComboBox3.Value)
    Dim ville As String: ville = UCase(Trim(TextBox6.Value))

    ' Plage de tolérance pour la surface (+/-20%)
    Dim surfaceMin As Double: surfaceMin = surfaceSouhaitee * 0.8
    Dim surfaceMax As Double: surfaceMax = surfaceSouhaitee * 1.2

    ' Variables de calcul
    Dim totalPrix As Double: totalPrix = 0
    Dim compteur As Long: compteur = 0
    Dim prixMoyen As Double


    Dim i As Long
    For i = 2 To lastRow
        ' Lecture des données dans les bonnes colonnes
        Dim prix As Variant: prix = ws.Cells(i, 1).Value           ' Col A
        Dim communeBase As String: communeBase = UCase(Trim(ws.Cells(i, 2).Value)) ' Col B
        Dim depBase As String: depBase = (ws.Cells(i, 3).Value)               ' Col C
        Dim surfaceBase As Double: surfaceBase = (ws.Cells(i, 6).Value)        ' Col F
        Dim piecesBase As Integer: piecesBase = (ws.Cells(i, 7).Value)         ' Col G

        ' Comparaison avec critères
        If depBase = dep And communeBase = ville And piecesBase = nbPieces And surfaceBase >= surfaceMin _
        And surfaceBase <= surfaceMax Then


            totalPrix = totalPrix + (prix)
            compteur = compteur + 1
        End If
    Next i


    ' Résultat
    If compteur > 0 Then
     
    Dim wsOut As Worksheet
    Set wsOut = ThisWorkbook.Sheets("DonnéesSaisies")

    Dim nextRow As Integer
    nextRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row + 1

    prixMoyen = totalPrix / compteur '

    ' Enregistrement des réponses dans DonnéesSaisies
    wsOut.Cells(nextRow, 1).Value = surfaceSouhaitee
    wsOut.Cells(nextRow, 2).Value = dep
    wsOut.Cells(nextRow, 3).Value = nbPieces
    wsOut.Cells(nextRow, 4).Value = logement
    wsOut.Cells(nextRow, 5).Value = ville
    wsOut.Cells(nextRow, 6).Value = prixMoyen

    MsgBox "Prix moyen estimé : " & Format(prixMoyen, "#,##0") & " €", vbInformation, "Estimation"
Else
    MsgBox "Aucun bien trouvé avec ces critères.", vbExclamation, "Pas de résultat"
End If

End Sub


