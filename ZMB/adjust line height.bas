Attribute VB_Name = "ģ��1"
Sub ��1()
Attribute ��1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��1 ��
' �������һ�и߶�
'
For i = 1 To Worksheets.Count
    Worksheets(i).Activate
    For Each r In Worksheets(i).UsedRange.Rows
        If r.Columns(1) = "�����������ȷ�ϣ�ǩ������" Then
            r.Select
            r.EntireRow.AutoFit
        End If
    Next
Next
End Sub
