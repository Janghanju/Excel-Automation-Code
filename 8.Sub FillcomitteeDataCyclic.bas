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

    ' "����" ��Ʈ ���� (�а�: D��, �̸�: F��, �Ҽ�: K��; �����ʹ� 10�����)
    On Error Resume Next
    Set committeeWS = ThisWorkbook.Worksheets("����")
    On Error GoTo 0
    If committeeWS Is Nothing Then
        MsgBox "'����' ��Ʈ�� �������� �ʽ��ϴ�.", vbCritical
        Exit Sub
    End If

    ' "����" ��Ʈ���� �����Ͱ� �ִ� ������ �� (D�� ����)
    lastRow = committeeWS.Cells(committeeWS.Rows.Count, "D").End(xlUp).Row

    ' �� �а��� ���� �Ҵ� ���ڸ� ����� Dictionary ����
    Set deptCounter = CreateObject("Scripting.Dictionary")
    
    ' ��� ��ũ��Ʈ�� ��ȸ (��ǥ��ǥ ��Ʈ: �̸��� "-"�� ���Ե� ��Ʈ)
    For Each ws In ThisWorkbook.Worksheets
        If InStr(ws.Name, "-") > 0 Then
            ' ��� ��Ʈ�� �а����� H7 ������ ���� (������ ���� ����)
            deptName = Trim(ws.Range("H7").Value)
            If deptName = "" Then GoTo NextSheet
            
            ' �ش� �а��� ���� ���ݱ��� �� ��° �ĺ��� �Ҵ��ߴ����� ��ϡ�����
            If Not deptCounter.exists(deptName) Then
                deptCounter.Add deptName, 1
            Else
                deptCounter(deptName) = deptCounter(deptName) + 1
            End If
            candidateIndex = deptCounter(deptName)  ' �̹� ��Ʈ�� �� �а����� �� ��° ��Ʈ������ �ǹ�
            
            ' "����" ��Ʈ���� �ش� �а� �ĺ� �Ѽ��� ����.
            candidateCount = 0
            For rowIndex = 10 To lastRow
                If Trim(committeeWS.Cells(rowIndex, "D").Value) = deptName Then
                    candidateCount = candidateCount + 1
                End If
            Next rowIndex
            
            If candidateCount > 0 Then
                ' �ĺ� ������ ��Ʈ ���� ������ ���� �������� ��ȯ��Ŵ
                effectiveCandidateIndex = ((candidateIndex - 1) Mod candidateCount) + 1
                
                matchCount = 0
                foundCandidate = False
                For rowIndex = 10 To lastRow
                    If Trim(committeeWS.Cells(rowIndex, "D").Value) = deptName Then
                        matchCount = matchCount + 1
                        If matchCount = effectiveCandidateIndex Then
                            ws.Range("H28").Value = committeeWS.Cells(rowIndex, "F").Value  ' ���� �̸�
                            ws.Range("C28").Value = committeeWS.Cells(rowIndex, "K").Value  ' �Ҽ�
                            foundCandidate = True
                            Exit For
                        End If
                    End If
                Next rowIndex
                
                ' �Ҵ��� �ĺ��� ã�� ���ϸ� �� Ŭ����
                If Not foundCandidate Then
                    On Error Resume Next
                    ws.Range("H28").ClearContents
                    ws.Range("C28").ClearContents
                    On Error GoTo 0
                End If
            Else
                ' �ش� �а��� ���� �ĺ��� ���ٸ� �� Ŭ����
                On Error Resume Next
                ws.Range("H28").ClearContents
                ws.Range("C28").ClearContents
                On Error GoTo 0
            End If
        End If
NextSheet:
    Next ws
    
    MsgBox "���� ������ �Է� �Ϸ�!", vbInformation
End Sub

