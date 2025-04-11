Attribute VB_Name = "Module1"
Sub ReorderSheetsByPattern()
    Dim ws As Worksheet
    Dim sheetList As Object
    Dim i As Integer, j As Integer
    Dim sheetName As String
    Dim parts As Variant
    Dim maxGroup As Integer

    ' 시트 목록을 저장할 Dictionary 생성
    Set sheetList = CreateObject("Scripting.Dictionary")

    ' 모든 시트를 검사하여 '-'이 포함된 시트만 정리
    For Each ws In ThisWorkbook.Sheets
        sheetName = ws.Name
        If InStr(sheetName, "-") > 0 Then
            parts = Split(sheetName, "-") ' 예: "1-1" → {"1", "1"}
            If UBound(parts) = 1 Then
                If Not sheetList.exists(parts(0)) Then
                    sheetList.Add parts(0), CreateObject("Scripting.Dictionary")
                End If
                sheetList(parts(0)).Add parts(1), ws
                If Val(parts(1)) > maxGroup Then maxGroup = Val(parts(1)) ' 최대 그룹 번호 저장
            End If
        End If
    Next ws

    ' 시트 이동 (1-1, 2-1, ... → 1-2, 2-2, ... → 1-3, 2-3, ...)
    For j = 1 To maxGroup ' -1, -2, -3 순서
        For i = 1 To sheetList.Count ' 그룹별 이동
            If sheetList.exists(CStr(i)) Then
                If sheetList(CStr(i)).exists(CStr(j)) Then
                    Set ws = sheetList(CStr(i))(CStr(j))
                    ws.Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                End If
            End If
        Next i
    Next j

    MsgBox "시트 정렬 완료!", vbInformation
End Sub

