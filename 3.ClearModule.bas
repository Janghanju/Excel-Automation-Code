Attribute VB_Name = "Module3"
Sub ClearCellsInSheets()
    Dim ws As Worksheet

    ' 모든 시트를 반복하며 지정된 범위 지우기
    For Each ws In ThisWorkbook.Sheets
        If InStr(ws.Name, "-") > 0 Then ' 이름에 "-"가 포함된 시트만 적용
            ws.Range("J18:J20").ClearContents ' J18:J20 범위의 내용 삭제
        End If
    Next ws
End Sub

