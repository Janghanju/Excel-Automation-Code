Attribute VB_Name = "Module1"
Sub FillCommitteeDataCyclic()
    Dim committeeWS As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long, rowIndex As Long
    Dim deptName As String
    Dim candidateIndex As Long
    Dim matchCount As Long
    Dim deptCounter As Object
    Dim candidateCount As Long, effectiveCandidateIndex As Long
    Dim foundCandidate As Boolean

    ' "위원" 시트 설정 (분과: D열, 이름: F열, 소속: K열; 데이터는 10행부터)
    On Error Resume Next
    Set committeeWS = ThisWorkbook.Worksheets("위원")
    On Error GoTo 0
    If committeeWS Is Nothing Then
        MsgBox "'위원' 시트가 존재하지 않습니다.", vbCritical
        Exit Sub
    End If

    ' "위원" 시트에서 데이터가 있는 마지막 행 (D열 기준)
    lastRow = committeeWS.Cells(committeeWS.Rows.Count, "D").End(xlUp).Row

    ' 각 분과별 위원 할당 숫자를 기록할 Dictionary 생성
    Set deptCounter = CreateObject("Scripting.Dictionary")
    
    ' 모든 워크시트를 순회 (발표평가표 시트: 이름에 "-"가 포함된 시트)
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "-") > 0 Then
            ' 대상 시트의 분과명은 H7 셀에서 읽음 (여분의 공백 제거)
            deptName = Trim(ws.Range("H7").Value)
            If deptName = "" Then GoTo NextSheet
            
            ' 해당 분과에 대해 지금까지 몇 번째 후보를 할당했는지를 기록·누적
            If Not deptCounter.exists(deptName) Then
                deptCounter.Add deptName, 1
            Else
                deptCounter(deptName) = deptCounter(deptName) + 1
            End If
            candidateIndex = deptCounter(deptName)  ' 이번 시트가 그 분과에서 몇 번째 시트인지를 의미
            
            ' "위원" 시트에서 해당 분과 후보 총수를 센다.
            candidateCount = 0
            For rowIndex = 10 To lastRow
                If Trim(committeeWS.Cells(rowIndex, "D").Value) = deptName Then
                    candidateCount = candidateCount + 1
                End If
            Next rowIndex
            
            If candidateCount > 0 Then
                ' 후보 수보다 시트 수가 많으면 모듈로 연산으로 순환시킴
                effectiveCandidateIndex = ((candidateIndex - 1) Mod candidateCount) + 1
                
                matchCount = 0
                foundCandidate = False
                For rowIndex = 10 To lastRow
                    If Trim(committeeWS.Cells(rowIndex, "D").Value) = deptName Then
                        matchCount = matchCount + 1
                        If matchCount = effectiveCandidateIndex Then
                            ws.Range("H28").Value = committeeWS.Cells(rowIndex, "F").Value  ' 위원 이름
                            ws.Range("C28").Value = committeeWS.Cells(rowIndex, "K").Value  ' 소속
                            foundCandidate = True
                            Exit For
                        End If
                    End If
                Next rowIndex
                
                ' 할당할 후보를 찾지 못하면 셀 클리어
                If Not foundCandidate Then
                    On Error Resume Next
                    ws.Range("H28").ClearContents
                    ws.Range("C28").ClearContents
                    On Error GoTo 0
                End If
            Else
                ' 해당 분과에 위원 후보가 없다면 셀 클리어
                On Error Resume Next
                ws.Range("H28").ClearContents
                ws.Range("C28").ClearContents
                On Error GoTo 0
            End If
        End If
NextSheet:
    Next ws
    
    MsgBox "위원 데이터 입력 완료!", vbInformation
End Sub

