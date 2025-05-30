Attribute VB_Name = "Run"
Sub lance()
Dim ws As Worksheet
Set ws = ActiveSheet
'on suprime et in cree les feuille
Feuil.SupprimerFeuilleResultat
Feuil.CreerEtActiverFeuilleInfo
'combiner les tableau
Importe.CombinerCSV
'Faire le recapitulatif
Calcul.CalculerValeursColonneE
Calcul.SommesEtAffichageSimplifie
'mise en forme tableau recapitulatif
Forme.ConvertirEnTableau
End Sub
