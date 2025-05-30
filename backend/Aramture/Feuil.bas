Attribute VB_Name = "Feuil"

Sub SupprimerFeuilleResultat()
    Dim ws As Worksheet
    Dim feuilleExiste As Boolean
    Dim feuilleExiste2 As Boolean

    ' Vérifier si la feuille "resultat" existe
    feuilleExiste = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "resultat" Then
            feuilleExiste = True
            Exit For
        End If
    Next ws

    ' Supprimer la feuille "resultat" si elle existe
    If feuilleExiste Then
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets("resultat").Delete
        Application.DisplayAlerts = True
        'MsgBox "La feuille 'resultat' a été supprimée.", vbInformation
    Else
        'MsgBox "La feuille 'resultat' n'existe pas.", vbInformation
    End If
End Sub
Sub CreerEtActiverFeuilleInfo()
    Dim ws As Worksheet

    ' Vérifier si la feuille "info" existe déjà
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("info")
    On Error GoTo 0

    ' Si la feuille "info" existe, la supprimer
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If

    ' Créer une nouvelle feuille appelée "info"
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "info"

    ' Activer la feuille "info"
    ws.Activate

    'MsgBox "La feuille 'info' a été créée et activée.", vbInformation
End Sub

