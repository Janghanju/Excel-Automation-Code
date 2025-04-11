Attribute VB_Name = "Module2"
Sub UpdateReviewerInfo()
    Dim ws As Worksheet
    Dim reviewerNames As Variant
    Dim reviewerAffiliations As Variant
    Dim subIndex As Integer

    ' ������ ���� �迭 ����
    reviewerNames = Array("", "�̸�1", "�̸�2", "�̸�3") ' �̸�
    reviewerAffiliations = Array("", "�Ҽ�1", "�Ҽ�2", "�Ҽ�3") ' �Ҽ�

    ' ��� ��Ʈ�� �ݺ��ϸ� ������Ʈ
    For Each ws In ThisWorkbook.Sheets
        If InStr(ws.Name, "-") > 0 Then ' ��Ʈ �̸��� "-"�� ���Ե� ��쿡�� ����
            ' ���� ��ȣ ����
            subIndex = CInt(Split(ws.Name, "-")(1)) - 1 ' ���� ��ȣ (1, 2, 3 -> 0���� ����)

            ' �迭���� �� �����Ͽ� �Է�
            ws.Range("������ ����ġ").Value = reviewerAffiliations(subIndex Mod UBound(reviewerAffiliations) + 1) ' �Ҽ� �Է�
            ws.Range("������ ����ġ").Value = reviewerNames(subIndex Mod UBound(reviewerNames) + 1) ' �̸� �Է�
        End If
    Next ws
End Sub



