Attribute VB_Name = "vue"
Public Const ioffset As Integer = 5
Public Const joffset As Integer = 2
Public Sub interface()
    For i = 0 To 8
        For j = 0 To 8
            Cells(i + ioffset, j + joffset).Font.Size = 24
            Cells(i + ioffset, j + joffset).ColumnWidth = 10
            Cells(i + ioffset, j + joffset).RowHeight = 50
            Cells(i + ioffset, j + joffset).HorizontalAlignment = xlCenter
            Cells(i + ioffset, j + joffset).VerticalAlignment = xlCenter
            Cells(i + ioffset, j + joffset).NumberFormat = "0;;;@"
        Next j
    Next i
End Sub
Public Function vueGetValue(i, j)
'à l'initialisation prend le chiffre qui est dans la case
    If IsEmpty(Cells(i + ioffset, j + joffset)) Or Cells(i + ioffset, j + joffset).Value = 0 Then
        vueGetValue = 0
    Else
        vueGetValue = Cells(i + ioffset, j + joffset).Value
    End If
End Function
Public Function vueSetValue(i, j, v)
'met sur la case de la feuille le chiffre qui correspond après résolution
     If IsEmpty(Cells(i + ioffset, j + joffset)) Or Cells(i + ioffset, j + joffset).Value = 0 Then
        Cells(i + ioffset, j + joffset).Value = v
        Cells(i + ioffset, j + joffset).Font.ColorIndex = 10
     End If
End Function
Public Sub reinit_sudoku()
'remet le tableau de la feuille à vide
    Dim i As Byte, j As Byte
    For i = 0 To 8
        For j = 0 To 8
            Cells(i + ioffset, j + joffset).ClearContents
            Cells(i + ioffset, j + joffset).Font.ColorIndex = 1
        Next j
    Next i
End Sub
Public Function verif_user(i, j) As Boolean
'renvoie true si l'entier rentré par l'utilisateur est entre 1 et 9. Si non, renvoie false
    If (Cells(i + ioffset, j + joffset).Value <= 9 & Cells(i + ioffset, j + joffset).Value >= 1) Or IsEmpty(Cells(i + ioffset, j + joffset)) Then
      verif_user = True
    Else
        verif_user = False
    End If
End Function
Public Sub placer_chiffres()
    Dim i As Byte, j As Byte
    For i = 0 To 8
        For j = 0 To 8
            Cells(i + ioffset, j + joffset).Value = sudoku(i, j)
        Next j
    Next i
End Sub
Public Sub easy()
'affiche un sudoku facile prédéterminé
    Cells(0 + ioffset, 2 + joffset).Value = 1
    Cells(0 + ioffset, 5 + joffset).Value = 7
    Cells(0 + ioffset, 7 + joffset).Value = 5
    Cells(0 + ioffset, 8 + joffset).Value = 2
    Cells(1 + ioffset, 0 + joffset).Value = 6
    Cells(1 + ioffset, 3 + joffset).Value = 3
    Cells(1 + ioffset, 5 + joffset).Value = 8
    Cells(1 + ioffset, 6 + joffset).Value = 7
    Cells(1 + ioffset, 8 + joffset).Value = 9
    Cells(2 + ioffset, 0 + joffset).Value = 5
    Cells(2 + ioffset, 5 + joffset).Value = 2
    Cells(2 + ioffset, 6 + joffset).Value = 4
    Cells(2 + ioffset, 7 + joffset).Value = 3
    Cells(2 + ioffset, 8 + joffset).Value = 6
    Cells(3 + ioffset, 1 + joffset).Value = 3
    Cells(3 + ioffset, 2 + joffset).Value = 6
    Cells(3 + ioffset, 3 + joffset).Value = 8
    Cells(3 + ioffset, 8 + joffset).Value = 4
    Cells(4 + ioffset, 0 + joffset).Value = 2
    Cells(4 + ioffset, 1 + joffset).Value = 7
    Cells(4 + ioffset, 2 + joffset).Value = 4
    Cells(4 + ioffset, 5 + joffset).Value = 6
    Cells(4 + ioffset, 7 + joffset).Value = 9
    Cells(5 + ioffset, 7 + joffset).Value = 7
    Cells(5 + ioffset, 8 + joffset).Value = 3
    Cells(6 + ioffset, 3 + joffset).Value = 5
    Cells(6 + ioffset, 4 + joffset).Value = 4
    Cells(6 + ioffset, 5 + joffset).Value = 3
    Cells(6 + ioffset, 8 + joffset).Value = 7
    Cells(7 + ioffset, 1 + joffset).Value = 2
    Cells(7 + ioffset, 7 + joffset).Value = 6
    Cells(8 + ioffset, 0 + joffset).Value = 7
    Cells(8 + ioffset, 3 + joffset).Value = 6
End Sub
Sub showAll()
    Worksheets("Start").Visible = True
    Worksheets("Game").Visible = True
End Sub
Public Function verifPlein() As Boolean
    verifPlein = False
    For i = 0 To 2
        For j = 0 To 2
            If IsEmpty(Cells(i + ioffset, j + joffset)) = False Then
                verifPlein = True
            End If
        Next j
    Next i
End Function
