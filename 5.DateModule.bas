Attribute VB_Name = "Module1"
Sub DateModule()
    Dim ws As Worksheet
    Dim targetDate As String
    
    ' �Է��� ��¥ ����
    targetDate = "2025. 4. 14."
    
    ' ��� ��Ʈ�� �ݺ��ϸ� Ȯ��
    For Each ws In ThisWorkbook.Sheets
        If InStr(ws.Name, "-") > 0 Then ' ��Ʈ �̸��� "-"�� ���Ե� ��츸 ����
            ws.Range("E27").Value = targetDate
        End If
    Next ws
    
    MsgBox "��� �ش� ��Ʈ�� ��¥ �Է� �Ϸ�!", vbInformation
End Sub

