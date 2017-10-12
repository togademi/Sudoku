Attribute VB_Name = "controleur"
Public Sub parcours_verif_colonne(j)
    Dim i As Byte
    For i = 0 To 8
        If sudoku(i, j) <> 0 Then
            tab_verif(sudoku(i, j) - 1) = False
        End If
    Next i
End Sub
Public Sub parcours_verif_ligne(i)
    Dim j As Byte
    For j = 0 To 8
        If sudoku(i, j) <> 0 Then
            tab_verif(sudoku(i, j) - 1) = False
        End If
    Next j
End Sub
Public Sub parcours_verif_cadre(ByVal i, ByVal j)
    Dim i0 As Byte, j0 As Byte
    i0 = i - (i Mod 3)
    j0 = j - (j Mod 3)
    For i = i0 To i0 + 2
        For j = j0 To j0 + 2
            If sudoku(i, j) <> 0 Then
                tab_verif(sudoku(i, j) - 1) = False
            End If
        Next j
    Next i
End Sub
Public Sub resolution(i, j)
    Dim a As Byte, c As Byte, z As Byte
    z = 0
    c = 0
    For a = 0 To 8
        If tab_verif(a) Then
            c = c + 1
            z = a
        End If
    Next a
    
    If c = 1 Then
        sudoku(i, j) = z + 1
    Else
        encore = True
    End If

End Sub
Public Sub reinitialiser_tab_verif()
'réinitialise le tableau de vérification
    Dim a As Byte
    For a = 0 To 8
        tab_verif(a) = True
    Next a
End Sub
Public Sub permut_nombres()
'permute les nombres de 1 à 9 dans un nouveau tableau. Le tableau final permuté est tabpermut_nombres
    Dim tabpermut_nombres(8) As Byte
    Dim tabnombres(8) As Byte
    Dim n As Byte, i As Byte
    For i = 0 To 8
        tabpermut_nombres(i) = 0
        tabnombres(i) = i + 1
    Next i
    For i = 0 To 8
        n = Int(Rnd * 9)
        Do While tabpermut_nombres(n) <> 0
            n = Int(Rnd * 9)
        Loop
        tabpermut_nombres(n) = tabnombres(i)
    Next i
    For i = 0 To 8
        For j = 0 To 8
            If sudoku(i, j) <> 0 Then
                sudoku(i, j) = tabpermut_nombres(sudoku(i, j) - 1)
            End If
        Next j
    Next i
End Sub
Public Sub permut_lignes()
'Permute les lignes par paquets de 3
    Dim i As Byte, j As Byte, n As Byte, c As Byte
    'Premier paquet
    For i = 0 To 2
        n = Int(Rnd * 3)
        For j = 0 To 8
            c = sudoku(i, j)
            sudoku(i, j) = sudoku(n, j)
            sudoku(n, j) = c
        Next j
    Next i
    'Deuxième paquet
    For i = 3 To 5
        n = Int(Rnd * 3) + 3
        For j = 0 To 8
            c = sudoku(i, j)
            sudoku(i, j) = sudoku(n, j)
            sudoku(n, j) = c
        Next j
    Next i
    'Troisième paquet
    For i = 6 To 8
        n = Int(Rnd * 3) + 6
        For j = 0 To 8
            c = sudoku(i, j)
            sudoku(i, j) = sudoku(n, j)
            sudoku(n, j) = c
        Next j
    Next i
End Sub
Public Sub permut_colonnes()
    'Permute les colonnes par paquets de 3
    Dim i As Byte, j As Byte, n As Byte, c As Byte
    'Premier paquet
    For j = 0 To 2
        n = Int(Rnd * 3)
        For i = 0 To 8
            c = sudoku(i, j)
            sudoku(i, j) = sudoku(i, n)
            sudoku(i, n) = c
        Next i
    Next j
    'Deuxième paquet
    For j = 3 To 5
        n = Int(Rnd * 3) + 3
        For i = 0 To 8
            c = sudoku(i, j)
            sudoku(i, j) = sudoku(i, n)
            sudoku(i, n) = c
        Next i
    Next j
    'Troisième paquet
    For j = 6 To 8
        n = Int(Rnd * 3) + 6
        For i = 0 To 8
            c = sudoku(i, j)
            sudoku(i, j) = sudoku(i, n)
            sudoku(i, n) = c
        Next i
    Next j
End Sub
Public Sub verifInput()
    Application.OnKey "1", "keyUn"
    Application.OnKey "2", "keyDeux"
    Application.OnKey "3", "keyTrois"
    Application.OnKey "4", "keyQuatre"
    Application.OnKey "5", "keyCinque"
    Application.OnKey "6", "keySix"
    Application.OnKey "7", "keySept"
    Application.OnKey "8", "keyHuit"
    Application.OnKey "9", "keyNeuf"
End Sub
Sub keyUn()
    Dim b As Boolean
    b = verifInput2(1)
End Sub
Sub keyDeux()
    Dim b As Boolean
    b = verifInput2(2)
End Sub
Sub keyTrois()
    Dim b As Boolean
    b = verifInput2(3)
End Sub
Sub keyQuatre()
    Dim b As Boolean
    b = verifInput2(4)
End Sub
Sub keyCinque()
    Dim b As Boolean
    b = verifInput2(5)
End Sub
Sub keySix()
    Dim b As Boolean
    b = verifInput2(6)
End Sub
Sub keySept()
    Dim b As Boolean
    b = verifInput2(7)
End Sub
Sub keyHuit()
    Dim b As Boolean
    b = verifInput2(8)
End Sub
Sub keyNeuf()
    Dim b As Boolean
    b = verifInput2(9)
End Sub
Function verifInput2(x As Byte) As Boolean
   If x >= 1 And x <= 9 Then
        'ActiveCell.Font.ColorIndex = 11
         verifInput2 = True
    Else
        'ActiveCell.Font.ColorIndex = 3
        verifInput2 = False
    End If
    ActiveCell.Value = x
End Function
