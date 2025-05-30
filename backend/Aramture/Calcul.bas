Attribute VB_Name = "Calcul"
Sub CalculerValeursColonneE()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim v As Integer
    Dim r As Integer
    v = 11
    r = 5
    ' Définir la feuille active
    Set ws = ActiveSheet

    ' Trouver la dernière ligne utilisée dans la colonne A
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Parcourir chaque cellule de la colonne A
    For i = 1 To lastRow
        On Error Resume Next ' Ignorer les erreurs de type incompatible
        If ws.Cells(i, r).Value = 6 Then
            ws.Cells(i, v).Value = ws.Cells(i, r + 2).Value * (ws.Cells(i, r + 3).Value * 2 + ws.Cells(i, r + 4).Value * 2 + 0.05)
        ElseIf ws.Cells(i, r).Value = 8 Or ws.Cells(i, r).Value = 10 Or ws.Cells(i, r).Value = 12 Or ws.Cells(i, r).Value = 14 Then
            ws.Cells(i, v).Value = ws.Cells(i, r + 2).Value * ws.Cells(i, r + 3).Value
        End If
        On Error GoTo 0 ' Réactiver la gestion des erreurs
    Next i

    'MsgBox "Les valeurs de la colonne E ont été calculées avec succès.", vbInformation
End Sub
Sub SommesEtAffichageSimplifie()

  Dim ws As Worksheet
  Dim wa As Worksheet
  Dim lastRow As Long
  Dim i As Long
  Dim valeurs As Variant 'Tableau pour stocker les sommes
  Dim valeurs2 As Variant
  Dim valeurs3 As Variant
  Dim barre As Double
  Dim j As Long

  ' Définir la feuille active
  Set ws = ActiveSheet

  ' Trouver la dernière ligne utilisée dans la colonne A
  lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

  ' Dimensionner le tableau pour les sommes (indices 6, 8, 10, 12, 14)
  ReDim valeurs(6 To 14, 1 To 1) 'Une seule colonne pour les sommes

  ' Parcourir chaque cellule de la colonne A
  For i = 1 To lastRow

    Select Case ws.Cells(i, "E").Value
      Case 6, 8, 10, 12, 14 'Cas groupés
        valeurs(ws.Cells(i, "E").Value, 1) = valeurs(ws.Cells(i, "E").Value, 1) + ws.Cells(i, "K").Value
    End Select

  Next i

  ' Afficher les résultats (boucle pour plus d'efficacité)
   Set wa = ThisWorkbook.Sheets.Add
    wa.Name = "resultat"
valeurs3 = Array("Armature", "longeur en ml", "Nbr barre", "Prix/Tonne", "Nbr barre/tonne", "Prix par barre", "Prix")
n = 1
  For i = 0 To 6
  wa.Cells(i + 1, n).Value = valeurs3(i)
  Next i
  valeurs2 = Array(375, 210, 135, 93, 68)
  barre = Application.WorksheetFunction.Ceiling(valeurs(i, 1), 1)
m = 2 'Colonne de départ (G)
j = m
  For i = 6 To 14 Step 2 'Pas de 2 pour 6, 8, 10, 12, 14
    wa.Cells(1, j).Value = "HA" & i
    wa.Cells(2, j).Value = valeurs(i, 1)
    wa.Cells(3, j).Value = Application.WorksheetFunction.Ceiling(WorksheetFunction.Ceiling(valeurs(i, 1), 1) / 12, 1)
    wa.Cells(4, j).Value = 525000
    wa.Cells(5, j).Value = valeurs2(j - m)
    wa.Cells(6, j).Value = 525000 / valeurs2(j - m)
    wa.Cells(7, j).Value = WorksheetFunction.Ceiling(valeurs(i, 1) / 12, 1) * 525000 / valeurs2(j - m)
    j = j + 1 'Passer à la colonne suivante
  Next i

End Sub
