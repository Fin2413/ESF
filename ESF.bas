Attribute VB_Name = "Module1"
Sub ExtractESFNumbers()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim startPos As Long
    Dim endPos As Long
    Dim extractedText As String
    
    ' ���������� ������ �� �������� ����
    Set ws = ActiveSheet
    ' ������� �������� ������� D (����� ������� 4)
    Set rng = ws.Range("D1:D" & ws.Cells(ws.Rows.Count, 4).End(xlUp).Row)
    
    ' ������������ ������ ������
    For Each cell In rng
        If Not IsEmpty(cell.Value) Then
            ' ���� ������ ������ "ESF-"
            startPos = InStr(cell.Value, "ESF-")
            
            If startPos > 0 Then
                ' ���� ����� ������ (������ ����������� ������ ����� ������)
                endPos = InStr(startPos, cell.Value, ")")
                If endPos > 0 Then
                    ' ��������� ����� ����� "ESF-" � ����������� �������
                    extractedText = Mid(cell.Value, startPos, endPos - startPos)
                Else
                    ' ���� ����������� ������ �� �������, ����� ����� �� ����� ������
                    extractedText = Mid(cell.Value, startPos)
                End If
                
                ' �������� ���������� ������ �� ��������� �����
                cell.Value = extractedText
            Else
                ' ���� "ESF-" �� ������, ������� ������
                cell.Value = ""
            End If
        End If
    Next cell

    MsgBox "��������� ���������!", vbInformation
End Sub

