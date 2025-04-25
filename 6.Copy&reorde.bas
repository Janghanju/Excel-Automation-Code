Attribute VB_Name = "Module1"
Sub DuplicateSheetsCorrectOrder()
    Dim ws As Worksheet
    Dim newSheet As Worksheet
    Dim sheetName As String
    Dim i As Integer, j As Integer
    Dim lastSheetNumber As Integer
    Dim copyCount As Integer
    Dim mainIndex As Integer
    Dim subIndex As Integer
    Dim positionIndex As Integer
    
    ' ������ ��Ʈ ��ȣ ���
    lastSheetNumber = ThisWorkbook.Sheets.Count
    
    ' �� ��ǥ��ǥ ��Ʈ �ݺ�
    For mainIndex = 1 To lastSheetNumber - 4
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets("��ǥ��ǥ " & mainIndex & "-1")
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            ' �⺻ ���� ���� ����
            copyCount = 4
            
            ' H7 ���� '������ȭ��' ���� ���� Ȯ���Ͽ� �߰� ����
            If InStr(ws.Range("H7").Value, "������ȭ��") > 0 Then
                copyCount = 6
            End If
            
            ' ���� ��Ʈ �ٷ� �ڿ� ������ ��ġ ����
            positionIndex = ws.Index
            
            ' �ݺ� ���� �� ������ �̸� ����
            For subIndex = 2 To (copyCount + 1)
                ws.Copy After:=ThisWorkbook.Sheets(positionIndex + subIndex - 2)
                Set newSheet = ActiveSheet
                
                ' ��Ʈ �̸� ����
                sheetName = "��ǥ��ǥ " & mainIndex & "-" & subIndex
                On Error Resume Next
                newSheet.Name = sheetName
                On Error GoTo 0
            Next subIndex
        End If
    Next mainIndex
End Sub

