Private Sub CommandButton1_Click()
'Заполнение массива положительными и отрицательными числами'
For i = 1 To 30
Cells(1, i) = Int((100 * Rnd) - 50)
Next i
End Sub

Private Sub CommandButton2_Click()
'нахождение номера третьего положительного элемента массива'
For i = 1 To 30
If Cells(1, i) > 0 Then
k = k + 1
If k = 3 Then
Exit For
End If
End If
Next i
MsgBox (i)
End Sub

Private Sub CommandButton3_Click()
'закрытие формы'
UserForm1.Hide
End Sub