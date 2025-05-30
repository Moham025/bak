Attribute VB_Name = "Importe"
Sub CombinerCSV()
    Dim cheminDossier As String
    Dim fichierCSV As String
    Dim feuilleDestination As Worksheet
    Dim derniereLigne As Long
    Dim premiereLigne As Boolean
    Dim ligne As String
    Dim lignes() As String
    Dim i As Long

    ' Effacer toutes les valeurs de la feuille
    EffacerToutesValeurs
    
    ' 1. Sélectionner le dossier
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "Sélectionner un dossier contenant les fichiers CSV"
        If .Show = -1 Then
            cheminDossier = .SelectedItems(1) & "\"
        Else
            MsgBox "Aucun dossier sélectionné.", vbInformation
            Exit Sub
        End If
    End With

    ' 2. Définir la feuille de destination comme la feuille active
    Set feuilleDestination = ActiveSheet
    premiereLigne = True

    ' 3. Parcourir les fichiers CSV du dossier
    fichierCSV = Dir(cheminDossier & "*.csv")
    Do While fichierCSV <> ""
        ' 4. Lire chaque fichier CSV sans l'ouvrir
        Open cheminDossier & fichierCSV For Input As #1
        Do While Not EOF(1)
            Line Input #1, ligne
            
            ' Remplacer les virgules par des points pour gérer les symboles décimaux
            ligne = Replace(ligne, ",", ".")
            
            ' Diviser les lignes en colonnes en utilisant ";" comme séparateur
            lignes = Split(ligne, ";")
            
            derniereLigne = feuilleDestination.Cells(Rows.Count, 1).End(xlUp).Row + 1
            If premiereLigne Then
                For i = LBound(lignes) To UBound(lignes)
                    feuilleDestination.Cells(derniereLigne, i + 1).Value = lignes(i)
                Next i
                premiereLigne = False
            Else
                derniereLigne = feuilleDestination.Cells(Rows.Count, 1).End(xlUp).Row + 1
                For i = LBound(lignes) To UBound(lignes)
                    feuilleDestination.Cells(derniereLigne, i + 1).Value = lignes(i)
                Next i
            End If
        Loop
        Close #1
        ' 5. Passer au fichier CSV suivant
        fichierCSV = Dir
    Loop
    'MsgBox "Fichiers CSV combinés avec succès !", vbInformation
    Extraire
    Convertir
End Sub
Sub EffacerToutesValeurs()
    Dim ws As Worksheet

    ' Définir la feuille active
    Set ws = ActiveSheet

    ' Effacer toutes les valeurs de la feuille
    ws.Cells.ClearContents

    'MsgBox "Toutes les valeurs de la feuille ont été effacées.", vbInformation
End Sub
Sub Convertir()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long

    ' Définir la feuille active
    Set ws = ActiveSheet

    ' Trouver la dernière ligne utilisée
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = 5

    ' Convertir les valeurs des cellules en nombres et leur format en format nombre
    For i = 1 To lastRow
        For j = 5 To 7
            If IsNumeric(ws.Cells(i, j).Value) Then
                ws.Cells(i, j).Value = CDbl(ws.Cells(i, j).Value) ' Convertir la valeur en nombre
                ws.Cells(i, j).NumberFormat = "0" ' Convertir le format en nombre
            End If
        Next j
        For j = 8 To 9
            If IsNumeric(ws.Cells(i, j).Value) Then
                ws.Cells(i, j).Value = CDbl(ws.Cells(i, j).Value) ' Convertir la valeur en nombre
                ws.Cells(i, j).NumberFormat = "0.00" ' Convertir le format en nombre
            End If
        Next j
    Next i
End Sub
Sub Extraire()
'
' Extraire les expression A = , B = & C =
'

'
    Cells.Replace What:="A = ", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False _
        , FormulaVersion:=xlReplaceFormula2
    Cells.Replace What:="B = ", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False _
        , FormulaVersion:=xlReplaceFormula2
    Cells.Replace What:="C = ", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False _
        , FormulaVersion:=xlReplaceFormula2
End Sub

