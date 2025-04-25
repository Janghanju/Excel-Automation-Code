Attribute VB_Name = "Module1"
Sub CopyAndFillEvaluationSheets()
    Dim wsSource As Worksheet
    Dim wsTemplate As Worksheet
    Dim wsNew As Worksheet
    Dim lastRow As Integer
    Dim i As Integer
    Dim newSheetName As String
    
    ' 기업정보 시트 및 평가표 템플릿 시트 설정
    Set wsSource = ThisWorkbook.Sheets("기업정보")
    Set wsTemplate = ThisWorkbook.Sheets("발표평가표")
    
    ' 기업정보 시트의 마지막 데이터 행 찾기
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' 데이터를 사용하여 발표평가표 복사 및 수정
    For i = 4 To lastRow
        ' 평가표 시트 복사
        wsTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Set wsNew = ActiveSheet
        
        ' 새 시트 이름 설정 (예: 발표평가표 1-2, 1-3 ...)
        newSheetName = "발표평가표 " & (i - 3) & "-1"
        wsNew.Name = newSheetName
        
        ' 데이터 입력
        wsNew.Range("C7").Value = wsSource.Cells(i, 4).Value ' D열 데이터
        wsNew.Range("C5").Value = wsSource.Cells(i, 2).Value ' B열 데이터
        wsNew.Range("H6").Value = wsSource.Cells(i, 3).Value ' C열 데이터
        wsNew.Range("H7").Value = wsSource.Cells(i, 1).Value ' A열 데이터
    Next i
End Sub

