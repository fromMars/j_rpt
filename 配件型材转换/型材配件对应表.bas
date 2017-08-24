Attribute VB_Name = "Ä£¿é1"
Public Function get_accessory(ByVal ws_no, ByVal cnt)
    x = DoEvents()
    max_row = Worksheets(ws_no).UsedRange.Rows.Count
    Dim a(5) As Variant
    'cnt = 1
    
    For i = 2 To max_row
        sn_row = 0
        a_sys = Worksheets(ws_no).Range("A" & i).Value
        a_prof = Worksheets(ws_no).Range("B" & i).Value
        'MsgBox (a_prof)
        If Len(a_prof) = 3 Then
            a_prof = "0" & a_prof
            'MsgBox (a_prof)
        End If
        
        a(0) = Worksheets(ws_no).Range("D" & i).Value
        a(1) = Worksheets(ws_no).Range("F" & i).Value
        a(2) = Worksheets(ws_no).Range("H" & i).Value
        a(3) = Worksheets(ws_no).Range("J" & i).Value
        a(4) = Worksheets(ws_no).Range("L" & i).Value
        
        On Error Resume Next
        sn_row = WorksheetFunction.Match(a_prof, Worksheets("Sheet1").Range("B:B"), 0)
        xx = Worksheets("Sheet1").Range("A" & sn_row).Value
        
        w_cnt = -1
        While sn_row > 0 And (a_sys <> xx) And w_cnt < 10000
            sn_row = sn_row + 1
            tmp_row = sn_row
            x1 = DoEvents()
            On Error Resume Next
            sn_row = WorksheetFunction.Match(a_prof, Worksheets("Sheet1").Range("B" & sn_row & ":B" & max_row), 0) + tmp_row - 1
            xx = Worksheets("Sheet1").Range("A" & sn_row).Value
            w_cnt = w_cnt + 1
            If a_sys = xx And Worksheets("Sheet1").Range("B" & sn_row) = a_prof Then
                w_cnt = -1
            End If
        Wend
        
        If (sn_row > 0) And (a_sys = xx) And w_cnt = -1 Then
            sn_col = 0
            For j = 0 To 5
                On Error Resume Next
                sn_col = WorksheetFunction.Match(a(j), Worksheets("Sheet1").Range("A" & sn_row & ":BZ" & sn_row), 0)
                If sn_col > 0 Then              'exclude same accessories
                    
                Else
                    max_col = Worksheets("Sheet1").Range("BZ" & sn_row).End(xlToLeft).Column
                    Worksheets("Sheet1").Cells(sn_row, max_col + 1) = a(j)
                End If
            Next
        Else
            Worksheets("Sheet1").Range("A" & cnt) = a_sys
            Worksheets("Sheet1").Range("B" & cnt) = a_prof
            Worksheets("Sheet1").Range("C" & cnt) = a(0)
            Worksheets("Sheet1").Range("D" & cnt) = a(1)
            Worksheets("Sheet1").Range("E" & cnt) = a(2)
            Worksheets("Sheet1").Range("F" & cnt) = a(3)
            Worksheets("Sheet1").Range("G" & cnt) = a(4)
            cnt = cnt + 1
        End If
    Next
    
    get_accessory = cnt
End Function

Public Function reverse()
    max_row = Worksheets("Sheet1").UsedRange.Rows.Count
    max_col = 26        'Worksheets("Sheet1").UsedRange.Columns.Count
    cnt = 1
    
    For r = 1 To max_row
        sys_name = Worksheets("Sheet1").Cells(r, 1)
        prof_name = Worksheets("Sheet1").Cells(r, 2)
        'profile_name = sys_name & " " & prof_name
        profile_name = sys_name & prof_name
        
        On Error Resume Next
        t_row = WorksheetFunction.Match(profile_name, Worksheets("Sheet3").Range("E:E"), 0)
        If t_row > 0 Then
            profile_name = Worksheets("Sheet3").Cells(t_row, 4)
        End If
        
        
        For c = 3 To max_col
            curr_acc = ""
            sn_row = 0
            x = DoEvents()
            curr_acc = Worksheets("Sheet1").Cells(r, c).Value
'            If curr_acc = "0502410" Then
'                MsgBox (curr_acc)
'            End If
            If curr_acc <> "" Then
                On Error Resume Next
                sn_row = WorksheetFunction.Match(curr_acc, Range("A:A"), 0)
'                MsgBox (curr_acc & ": " & Range("A:A")(sn_row).Value)
                
                If sn_row > 0 Then
                    m_col = Worksheets("Sheet2").Range("IV" & sn_row).End(xlToLeft).Column
                    Worksheets("Sheet2").Cells(sn_row, m_col + 1).Value = profile_name
                Else
                    Worksheets("Sheet2").Range("A" & cnt).Value = curr_acc
                    Worksheets("Sheet2").Cells(cnt, 2).Value = profile_name
                    cnt = cnt + 1
                End If
            End If
        Next
    Next
    
End Function

Sub accessorys_profile()
'For i = 1 To Worksheets.Count
'    MsgBox (Worksheets(i).Name & ": " & i)
'Next

'from here -->
Worksheets("Sheet1").Cells.Clear
Worksheets("Sheet1").Range("B1:BZ10000").NumberFormat = "@"
Worksheets("Sheet2").Cells.Clear
Worksheets("Sheet2").Range("A1:BZ10000").NumberFormat = "@"

a = get_accessory(2, 1)
a = get_accessory(3, 1)
a = get_accessory(4, 1)
'--->here ends

b = reverse

MsgBox ("Finished!")

'Dim myColl As New Collection
'Dim A(1) As String
'A(0) = "b"
'A(1) = "c"
'myColl.Add A(), "a"
'
'MsgBox myColl("a")(1)
End Sub
