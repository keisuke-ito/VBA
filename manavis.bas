  Sub extraction()
    If Worksheets(1).name <> "抽出" Then
        Worksheets.Add(Before:=Worksheets(1)).name = "抽出"
    End If
    Set s1 = Worksheets(1)
    s1.Cells.Clear
 
    Dim s_name '  Student name
    Dim s_time '  Attend
    Dim s_number ' Manabis ID
    Dim init_row_name, init_row_time, init_column_cell, init_col_last, color1, color2
    Dim i, c, d, d_f, sDATE, sLast, Y, M
    Dim step As Integer
    Dim name As Integer
    Dim id As Integer
    Dim time As Integer
    

    ' Reflect Extracted data to Schedule Sheet
    Y = Application.InputBox("対象年月", " 現在、西暦何年ですか?", Type:=1)
    M = Worksheets(2).Range("A3").Value
    
    sDATE = DateSerial(Y, M, 1)
    '   翌月1日の前日の取得
    sLast = Format(DateAdd("d", -1, DateAdd("m", 1, sDATE)), "d")
    

    ' Make Sutudent Attendance Table
    init_column_cell = 2
    init_col_last = 26
    init_row_name = 2
    init_row_id = 3
    init_row_time = 4
    color1 = RGB(204, 204, 255)
    color2 = RGB(255, 255, 255)
    
    s1.Range(Cells(init_row_name, 1), Cells(init_row_name + 3 * (sLast - 1) + 2, init_col_last)).Borders.LineStyle = True
    s1.Range(Cells(init_row_name, 1), Cells(init_row_name + 3 * (sLast - 1) + 2, init_col_last)).HorizontalAlignment = xlCenter
    s1.Range(Cells(init_row_name, 1), Cells(init_row_name + 3 * (sLast - 1) + 2, init_col_last)).RowHeight = 17
    s1.Range("A:A").ColumnWidth = 14
    s1.Range("B:Z").ColumnWidth = 12
    
    For d = 1 To sLast
        c = 0
        step = 3 * (d - 1)
        step2 = 6 * (d - 1)
        '  Input day, id and class time into each rows
        s1.Cells(init_row_name + step, 1).Value = d
        s1.Cells(init_row_id + step, 1).Value = "マナビス生番号"
        s1.Cells(init_row_time + step, 1).Value = "Time"
        If d <= sLast / 2 Then
            s1.Range(Cells(init_row_name + step2, 1), Cells(init_row_time + step2, init_col_last)).Interior.Color = color1
            s1.Range(Cells(init_row_name + 3 + step2, 1), Cells(init_row_time + 3 + step2, init_col_last)).Interior.Color = color2
        
        ElseIf d = 31 Then
            s1.Range(Cells(init_row_name + 90, 1), Cells(init_row_time + 90, init_col_last)).Interior.Color = color1
        
        ElseIf d = 29 And sLast = 29 Then
            s1.Range(Cells(init_row_name + 84, 1), Cells(init_row_time + 84, init_col_last)).Interior.Color = color1
            
        End If
        
        '  Operation in each sheets
        For i = 2 To Sheets.Count
            '  Extract name, id and time in each sheets.
            
            s_name = Worksheets(i).Range("L3").Value
            s_id = Worksheets(i).Range("F3").Value
            s_time = Worksheets(i).Cells(6 + 4 * (d - 1), 5).Value
            ' if time cell has something, input data to sheet1
            If Len(s_time) > 3 Then
                s1.Cells(init_row_name + step, init_column_cell + c).Value = s_name
                s1.Cells(init_row_id + step, init_column_cell + c).Value = s_id
                s1.Cells(init_row_time + step, init_column_cell + c).Value = s_time
                c = c + 1
            End If
        Next
        
    Next
    

    '  Sort data
    step = 0
    For i = 1 To sLast
        step = 3 * (i - 1)
        name = init_row_name + step
        id = init_row_id + step
        time = init_row_time + step
        With Sheets(1)
            .Range(.Cells(name, 2), .Cells(time, 26)).Sort key1:=.Cells(time, 3), _
                order1:=xlAscending, _
                Orientation:=xlSortRows
        End With
    Next
    

    ' Preview for Print
    If ActiveWindow.View <> xlPageBreakPreview Then
        ActiveWindow.View = xlPageBreakPreview
        ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
        ' Set ActiveSheet.VPageBreaks(1).Location = Range(Cells(init_row_name, init_col_last))
        ActiveSheet.PageSetup.PrintArea = Range(Cells(init_row_name, 1), Cells(init_row_time + 3 * (sLast - 1), init_col_last)).Address
    End If
 End Sub
