Attribute VB_Name = "模块1"
Sub 宏1()
Attribute 宏1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 宏1 宏
' 调整最后一行高度
'
For i = 1 To Worksheets.Count
    Worksheets(i).Activate
    For Each r In Worksheets(i).UsedRange.Rows
        If r.Columns(1) = "部门主管审核确认（签名）：" Then
            r.Select
            r.EntireRow.AutoFit
        End If
    Next
Next
End Sub
