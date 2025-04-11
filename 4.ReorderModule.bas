Attribute VB_Name = "Module1"
Sub ReorderSheetsByPattern()
    Dim ws As Worksheet
    Dim sheetList As Object
    Dim i As Integer, j As Integer
    Dim sheetName As String
    Dim parts As Variant
    Dim maxGroup As Integer

    ' ��Ʈ ����� ������ Dictionary ����
    Set sheetList = CreateObject("Scripting.Dictionary")

    ' ��� ��Ʈ�� �˻��Ͽ� '-'�� ���Ե� ��Ʈ�� ����
    For Each ws In ThisWorkbook.Sheets
        sheetName = ws.Name
        If InStr(sheetName, "-") > 0 Then
            parts = Split(sheetName, "-") ' ��: "1-1" �� {"1", "1"}
            If UBound(parts) = 1 Then
                If Not sheetList.exists(parts(0)) Then
                    sheetList.Add parts(0), CreateObject("Scripting.Dictionary")
                End If
                sheetList(parts(0)).Add parts(1), ws
                If Val(parts(1)) > maxGroup Then maxGroup = Val(parts(1)) ' �ִ� �׷� ��ȣ ����
            End If
        End If
    Next ws

    ' ��Ʈ �̵� (1-1, 2-1, ... �� 1-2, 2-2, ... �� 1-3, 2-3, ...)
    For j = 1 To maxGroup ' -1, -2, -3 ����
        For i = 1 To sheetList.Count ' �׷캰 �̵�
            If sheetList.exists(CStr(i)) Then
                If sheetList(CStr(i)).exists(CStr(j)) Then
                    Set ws = sheetList(CStr(i))(CStr(j))
                    ws.Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                End If
            End If
        Next i
    Next j

    MsgBox "��Ʈ ���� �Ϸ�!", vbInformation
End Sub

