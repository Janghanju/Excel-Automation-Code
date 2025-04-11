Attribute VB_Name = "Module1"
Sub DateModule()
    Dim ws As Worksheet
    Dim targetDate As String
    
    ' 입력할 날짜 설정
    targetDate = "2025. 4. 14."
    
    ' 모든 시트를 반복하며 확인
    For Each ws In ThisWorkbook.Sheets
        If InStr(ws.Name, "-") > 0 Then ' 시트 이름에 "-"가 포함된 경우만 실행
            ws.Range("E27").Value = targetDate
        End If
    Next ws
    
    MsgBox "모든 해당 시트에 날짜 입력 완료!", vbInformation
End Sub

