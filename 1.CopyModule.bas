Attribute VB_Name = "Module1"
Sub DuplicateSheets()
    Dim ws As Worksheet
    Dim newSheet As Worksheet
    Dim sheetName As String
    Dim i As Integer
    Dim lastSheetNumber As Integer

    ' ������ ��Ʈ ��ȣ ���
    lastSheetNumber = ThisWorkbook.Sheets.Count - 2 ' ��ũ�� �� �� ��Ʈ �������� ���� ��Ʈ ����

    ' ù ��° �ݺ�: -1
    For i = 1 To lastSheetNumber
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(i))
        On Error GoTo 0

        If Not ws Is Nothing Then
            ' ���� ��Ʈ �̸� ����
            ws.Name = i & "-1"
        End If
    Next i

    ' �� ��° �ݺ�: -2
    For i = 1 To lastSheetNumber
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(i & "-1")
        On Error GoTo 0

        If Not ws Is Nothing Then
            ' ���纻 ���� �� �̸� ����
            ws.Copy After:=ws
            Set newSheet = ActiveSheet
            sheetName = i & "-2"
            On Error Resume Next
            newSheet.Name = sheetName
            On Error GoTo 0
        End If
    Next i

    ' �� ��° �ݺ�: -3
    For i = 1 To lastSheetNumber
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(i & "-2")
        On Error GoTo 0

        If Not ws Is Nothing Then
            ' ���纻 ���� �� �̸� ����
            ws.Copy After:=ws
            Set newSheet = ActiveSheet
            sheetName = i & "-3"
            On Error Resume Next
            newSheet.Name = sheetName
            On Error GoTo 0
        End If
    Next i
End Sub

