Attribute VB_Name = "modele"
Option Base 0
Public sudoku(8, 8) As Integer, encore As Boolean, tab_verif(8) As Boolean, sudokuBool(8, 8) As Boolean
Sub mod_init()
'initialise le tableau avec les chiffres de la feuille
Dim i As Byte, j As Byte
For i = 0 To 8
    For j = 0 To 8
            sudoku(i, j) = vueGetValue(i, j)
    Next j
Next i
End Sub
Sub resoudre()
'dans chaque case du parcours elle détermine si un chiffre est solution évidente. Si non, on passe à la suivante
    Dim v As Byte
    encore = True
    Call reinitialiser_tab_verif
    Do While encore = True
'tant qu'il y a des cases vides en sudoku, la boucle de résolution continue
    encore = False
For i = 0 To 8
    For j = 0 To 8
        If sudoku(i, j) = 0 Then
            Call parcours_verif_colonne(j)
            Call parcours_verif_ligne(i)
            Call parcours_verif_cadre(ByVal i, ByVal j)
            Call resolution(i, j)
            Call reinitialiser_tab_verif
           
        End If
    Next j
Next i
Loop
For i = 0 To 8
    For j = 0 To 8
       Call vueSetValue(i, j, sudoku(i, j))
    Next j
Next i
End Sub
Sub resolutionFinale()
    Call mod_init
    Call resoudre
End Sub
Sub random()
    If verifPlein() = False Then
        Call easy
    End If

    Call mod_init
    Call permut_nombres
    Call permut_lignes
    Call permut_colonnes
    Call placer_chiffres
End Sub
Sub start()
    Worksheets("Game").Visible = True
    Worksheets("Start").Visible = False
    Call reinit_sudoku
    Call verifInput
End Sub
Sub startup()
    Worksheets("Start").Visible = True
    Worksheets("Game").Visible = False
    Call interface
End Sub

