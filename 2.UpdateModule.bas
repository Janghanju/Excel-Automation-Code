Attribute VB_Name = "Module2"
Sub UpdateReviewerInfo()
    Dim ws As Worksheet
    Dim reviewerNames As Variant
    Dim reviewerAffiliations As Variant
    Dim subIndex As Integer

    ' 평가위원 정보 배열 설정
    reviewerNames = Array("", "이름1", "이름2", "이름3") ' 이름
    reviewerAffiliations = Array("", "소속1", "소속2", "소속3") ' 소속

    ' 모든 시트를 반복하며 업데이트
    For Each ws In ThisWorkbook.Sheets
        If InStr(ws.Name, "-") > 0 Then ' 시트 이름에 "-"가 포함된 경우에만 적용
            ' 서브 번호 추출
            subIndex = CInt(Split(ws.Name, "-")(1)) - 1 ' 서브 번호 (1, 2, 3 -> 0부터 시작)

            ' 배열에서 값 선택하여 입력
            ws.Range("수정할 셀위치").Value = reviewerAffiliations(subIndex Mod UBound(reviewerAffiliations) + 1) ' 소속 입력
            ws.Range("수정할 셀위치").Value = reviewerNames(subIndex Mod UBound(reviewerNames) + 1) ' 이름 입력
        End If
    Next ws
End Sub



