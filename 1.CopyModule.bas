Attribute VB_Name = "Module1"
Sub DuplicateSheets()
    Dim ws As Worksheet
    Dim newSheet As Worksheet
    Dim sheetName As String
    Dim i As Integer
    Dim lastSheetNumber As Integer

    ' 마지막 시트 번호 계산
    lastSheetNumber = ThisWorkbook.Sheets.Count - 2 ' 워크북 내 총 시트 개수에서 기존 시트 제외

    ' 첫 번째 반복: -1
    For i = 1 To lastSheetNumber
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(i))
        On Error GoTo 0

        If Not ws Is Nothing Then
            ' 원본 시트 이름 수정
            ws.Name = i & "-1"
        End If
    Next i

    ' 두 번째 반복: -2
    For i = 1 To lastSheetNumber
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(i & "-1")
        On Error GoTo 0

        If Not ws Is Nothing Then
            ' 복사본 생성 및 이름 수정
            ws.Copy After:=ws
            Set newSheet = ActiveSheet
            sheetName = i & "-2"
            On Error Resume Next
            newSheet.Name = sheetName
            On Error GoTo 0
        End If
    Next i

    ' 세 번째 반복: -3
    For i = 1 To lastSheetNumber
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(i & "-2")
        On Error GoTo 0

        If Not ws Is Nothing Then
            ' 복사본 생성 및 이름 수정
            ws.Copy After:=ws
            Set newSheet = ActiveSheet
            sheetName = i & "-3"
            On Error Resume Next
            newSheet.Name = sheetName
            On Error GoTo 0
        End If
    Next i
End Sub

