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
    
    ' 마지막 시트 번호 계산
    lastSheetNumber = ThisWorkbook.Sheets.Count
    
    ' 각 발표평가표 시트 반복
    For mainIndex = 1 To lastSheetNumber - 4
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets("발표평가표 " & mainIndex & "-1")
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            ' 기본 복사 개수 설정
            copyCount = 4
            
            ' H7 셀에 '에너지화공' 포함 여부 확인하여 추가 복사
            If InStr(ws.Range("H7").Value, "에너지화공") > 0 Then
                copyCount = 6
            End If
            
            ' 기준 시트 바로 뒤에 복사할 위치 설정
            positionIndex = ws.Index
            
            ' 반복 복사 및 순차적 이름 설정
            For subIndex = 2 To (copyCount + 1)
                ws.Copy After:=ThisWorkbook.Sheets(positionIndex + subIndex - 2)
                Set newSheet = ActiveSheet
                
                ' 시트 이름 변경
                sheetName = "발표평가표 " & mainIndex & "-" & subIndex
                On Error Resume Next
                newSheet.Name = sheetName
                On Error GoTo 0
            Next subIndex
        End If
    Next mainIndex
End Sub

