Attribute VB_Name = "Module1"
Sub CopyAndFillEvaluationSheets()
    Dim wsSource As Worksheet
    Dim wsTemplate As Worksheet
    Dim wsNew As Worksheet
    Dim lastRow As Integer
    Dim i As Integer
    Dim newSheetName As String
    
    ' ������� ��Ʈ �� ��ǥ ���ø� ��Ʈ ����
    Set wsSource = ThisWorkbook.Sheets("�������")
    Set wsTemplate = ThisWorkbook.Sheets("��ǥ��ǥ")
    
    ' ������� ��Ʈ�� ������ ������ �� ã��
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' �����͸� ����Ͽ� ��ǥ��ǥ ���� �� ����
    For i = 4 To lastRow
        ' ��ǥ ��Ʈ ����
        wsTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Set wsNew = ActiveSheet
        
        ' �� ��Ʈ �̸� ���� (��: ��ǥ��ǥ 1-2, 1-3 ...)
        newSheetName = "��ǥ��ǥ " & (i - 3) & "-1"
        wsNew.Name = newSheetName
        
        ' ������ �Է�
        wsNew.Range("C7").Value = wsSource.Cells(i, 4).Value ' D�� ������
        wsNew.Range("C5").Value = wsSource.Cells(i, 2).Value ' B�� ������
        wsNew.Range("H6").Value = wsSource.Cells(i, 3).Value ' C�� ������
        wsNew.Range("H7").Value = wsSource.Cells(i, 1).Value ' A�� ������
    Next i
End Sub

