Attribute VB_Name = "Module3"
Sub ClearCellsInSheets()
    Dim ws As Worksheet

    ' ��� ��Ʈ�� �ݺ��ϸ� ������ ���� �����
    For Each ws In ThisWorkbook.Sheets
        If InStr(ws.Name, "-") > 0 Then ' �̸��� "-"�� ���Ե� ��Ʈ�� ����
            ws.Range("J18:J20").ClearContents ' J18:J20 ������ ���� ����
        End If
    Next ws
End Sub

