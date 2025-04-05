Attribute VB_Name = "Main"
Dim RowMax As Long      'Максимальное количество строк
Dim ColMax As Long      'Максимальное количество колонок
Dim Changed As Boolean  'Изменения во время последнего прогона
Dim Lenght As Long      'Общая длина по задаче
Dim StartX As Long      'Начальная точка цепочки
Dim StartY As Long      'Начальная точка цепочки
Dim chains As Collection 'Найденные цепочки
Dim m() As Integer

Sub Main()
    Dim i As Long
    Dim j As Long
    

    ClearCells
    
    Do
        Changed = False
        For i = 2 To RowMax - 1
            For j = 2 To ColMax - 1
                If Cells(i, j) <> "" And m(i, j) = 0 Then
                    Set chains = New Collection
                    Cells(i, j).Interior.Color = vbRed
                    Lenght = Cells(i, j)
                    StartX = i
                    StartY = j
                    DoEvents
                    find i, j, Lenght, "", ""
                    If chains.Count = 1 Then
                        SetChain chains(1)
                        Changed = True
                    Else
                        Cells(i, j).Interior.Pattern = xlNone
                    End If
                End If
            Next
        Next
    Loop While Changed
    
End Sub

Sub find(ByVal x As Long, ByVal y As Long, ByVal l As Long, ByVal chain As String, ByVal n As String)
    
    may = False
    If n = "" And m(x, y) = 0 Then may = True
    If Cells(x, y) = "" And m(x, y) = 0 Then may = True
    If l = 1 And Cells(x, y) = Lenght And m(x, y) = 0 Then may = True
    If Not may Then Exit Sub
    
    m(x, y) = -2
    chain = chain + n
    l = l - 1
    
    If l = 0 Then
        'Цепочка кончилась
        If Cells(x, y) = Lenght Then chains.Add chain
        m(x, y) = 0
        Exit Sub
    End If
    
    find x - 1, y, l, chain, "1"
    find x, y + 1, l, chain, "2"
    find x + 1, y, l, chain, "3"
    find x, y - 1, l, chain, "4"
    
    'Варианты перебрали, откатываемся на шаг назад
    m(x, y) = 0
    DoEvents
End Sub

'Рисоваине цепочки в буфере
Sub SetChain(chain As String)
    x = StartX
    y = StartY
    m(x, y) = -1
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
        m(x, y) = -1
    Next
End Sub

Sub pixel(ByVal x As Integer, ByVal y As Integer)
    Cells(x, y).Borders.Weight = 4
    Cells(x, y).Interior.Color = vbGreen
End Sub

Sub ExitButton()
    End
End Sub

Sub ClearCells()
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
    ReDim m(1 To RowMax, 1 To ColMax)
    For i = 1 To ColMax
        For j = 1 To RowMax
            b = 0
            If Cells(i, j) = "#" Then b = -1
            m(i, j) = b
        Next
    Next
    Set ramka = Range(Cells(1, 1), Cells(RowMax, ColMax))
    ramka.ClearFormats
    ramka.HorizontalAlignment = xlCenter
    For i = 2 To RowMax - 1
        For j = 2 To ColMax - 1
            If (i + j) Mod 2 = 0 Then Cells(i, j).Interior.Color = RGB(200, 230, 255)
        Next
    Next
End Sub