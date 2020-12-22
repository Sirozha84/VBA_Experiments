Attribute VB_Name = "Main"
Dim Shift As Long       'Значение сдвига буфера
Dim RowMax As Long      'Максимальное количество строк
Dim ColMax As Long      'Максимальное количество колонок
Dim Changed As Boolean  'Изменения во время последнего прогона
Dim Lenght As Long      'Общая длина по задаче
Dim StartX As Long      'Начальная точка цепочки
Dim StartY As Long      'Начальная точка цепочки
Dim chains As Collection 'Найденные цепочки

Sub Main()
    Dim i As Long
    Dim j As Long
    
    'Подготовка буфера и очистка поля
    i = 1
    Do While Cells(1, i) = "#"
        i = i + 1
    Loop
    RowMax = i - 1
    i = 1
    Do While Cells(1, i) = "#"
        i = i + 1
    Loop
    ColMax = i - 1
    Shift = ColMax + 5
    For i = 1 To ColMax
        For j = 1 To RowMax
            b = ""
            If Cells(i, j) = "#" Then b = "#"
            Cells(i, j + Shift) = b
        Next
    Next
    Set ramka = Range(Cells(1, 1), Cells(RowMax, ColMax))
    ramka.ClearFormats
    ramka.HorizontalAlignment = xlCenter
    
    Do
        Changed = False
        For i = 2 To RowMax - 1
            For j = 2 To ColMax - 1
                If Cells(i, j) <> "" And Cells(i, j + Shift) = "" Then
                    Set chains = New Collection
                    Cells(i, j).Interior.Color = vbRed
                    Lenght = Cells(i, j)
                    
                    StartX = i
                    StartY = j
                    DoEvents
                    Application.ScreenUpdating = False
                    find i, j, Lenght, "", ""
                     
                    If chains.Count = 1 Then
                        SetChain chains(1)
                        Changed = True
                    Else
                        Cells(i, j).Interior.Pattern = xlNone
                    End If
                    Application.ScreenUpdating = True
                    DoEvents
                End If
            Next
        Next
    Loop While Changed
End Sub

Sub find(ByVal x As Long, ByVal y As Long, ByVal l As Long, ByVal chain As String, ByVal n As String)
    
    may = False
    If n = "" And Cells(x, y + Shift) = "" Then may = True
    If Cells(x, y) = "" And Cells(x, y + Shift) = "" Then may = True
    If l = 1 And Cells(x, y) = Lenght And Cells(x, y + Shift) = "" Then may = True
    If Not may Then Exit Sub
    
    Cells(x, y + Shift) = "*"
    chain = chain + n
    l = l - 1
    
    If l = 0 Then
        'Цепочка кончилась
        If Cells(x, y) = Lenght Then chains.Add chain
        Cells(x, y + Shift) = ""
        Exit Sub
    End If
    
    find x - 1, y, l, chain, "1"
    find x, y + 1, l, chain, "2"
    find x + 1, y, l, chain, "3"
    find x, y - 1, l, chain, "4"
    
    'Варианты перебрали, откатываемся на шаг назад
    Cells(x, y + Shift) = ""
    
End Sub

'Рисоваине цепочки в буфере
Sub SetChain(chain As String)
    x = StartX
    y = StartY
    Cells(x, y + Shift) = "#"
    pixel x, y
    For i = 1 To Len(chain)
        n = Mid(chain, i, 1)
        If n = "1" Then
            x = x - 1
            pixel x, y
            Cells(x, y).Borders(xlEdgeBottom).Weight = 1
        End If
        If n = "2" Then
            y = y + 1
            pixel x, y
            Cells(x, y).Borders(xlEdgeLeft).Weight = 1
        End If
        If n = "3" Then
            x = x + 1
            pixel x, y
            Cells(x, y).Borders(xlEdgeTop).Weight = 1
        End If
        If n = "4" Then
            y = y - 1
            pixel x, y
            Cells(x, y).Borders(xlEdgeRight).Weight = 1
        End If
        Cells(x, y + Shift) = "#"
    Next
End Sub

Sub pixel(ByVal x As Integer, ByVal y As Integer)
    Cells(x, y).Borders.Weight = 4
    Cells(x, y).Interior.Color = vbGreen
End Sub
