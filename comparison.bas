Attribute VB_Name = "�绰��ȶ�ģ��"
Sub ����ظ�������()
    If Len(Range("b3")) = 0 And Len(Range("c3")) = 0 And Len(Range("d3")) = 0 And Len(Range("e3")) = 0 And Len(Range("f3")) = 0 Then
        MsgBox "����ȷ�����ļ���"
        Exit Sub
    End If
    Dim dictionary, preg, arraySource, arrayContent, i As Integer, currentRow As Integer, countRow As Integer, errorRow As Integer, repeatRange As String, nameAddress As String, isErrors
    Set dictionary = CreateObject("Scripting.Dictionary")
    Set preg = CreateObject("VBSCRIPT.REGEXP")
    With preg
        .Global = True
        .IgnoreCase = True
        .Pattern = "\d+\-\d+"
    End With
    countRow = [b65536].End(xlUp).Row
    '��ֵ 9��10 �� 13 ���Էֱ�ת��Ϊ�Ʊ�������з��ͻس���
    Cells.Replace What:=" ", Replacement:=""
    Cells.Replace What:=Chr(9), Replacement:=""
    Cells.Replace What:=Chr(10), Replacement:=""
    Cells.Replace What:=Chr(13), Replacement:=""
    arraySource = Range("b3:f" & countRow)
    ReDim arrayContent(1 To countRow, 1 To 9)
    Columns.Font.ColorIndex = 0
    isErrors = Array(False, False, False, False, False, False, False, False, False)
    arrayTableTitle = Array("�绰1�ظ���", "�绰2�ظ���", "�绰3�ظ���", "����+��ַ�ظ���", "�绰1λ��", "�绰2λ��", "�绰3λ��", "�����Ƿ�Ϊ��", "�����绰��Ϊ��")
    For i = 1 To UBound(arraySource)
        currentRow = i + 2
        '�绰1
        arraySource(i, 3) = CStr(arraySource(i, 3))
        If Len(arraySource(i, 3)) > 0 Then
            '�ظ���
            If dictionary.Exists(arraySource(i, 3)) Then
                repeatRange = dictionary(arraySource(i, 3))
                arrayContent(i, 1) = "�ظ�����" & repeatRange & "��Ԫ��"
                errorRow = Right(repeatRange, Len(repeatRange) - 1)
                arrayContent(errorRow - 2, ��һ���ظ�����(repeatRange)) = "�ظ�����D" & currentRow & "��Ԫ��"
                Rows(errorRow).Font.ColorIndex = 3
                Rows(currentRow).Font.ColorIndex = 3
            Else
                dictionary(CStr(arraySource(i, 3))) = "D" & currentRow
            End If
            '�ж�λ��
            If Len(arraySource(i, 3)) = 7 And IsNumeric(arraySource(i, 3)) = True Then
                arrayContent(i, 5) = "�̶��绰"
            ElseIf preg.Test(arraySource(i, 3)) = True Then
                arrayContent(i, 5) = "�̶��绰"
            ElseIf Len(arraySource(i, 3)) = 11 And IsNumeric(arraySource(i, 3)) = True Then
                arrayContent(i, 5) = "�ƶ��绰"
            Else
                arrayContent(i, 5) = "λ���������ڷ�����"
                Rows(currentRow).Font.ColorIndex = 17
            End If
        End If
        '�绰2
        arraySource(i, 4) = CStr(arraySource(i, 4))
        If Len(arraySource(i, 4)) > 0 Then
            '�ظ���
            If dictionary.Exists(arraySource(i, 4)) Then
                repeatRange = dictionary(arraySource(i, 4))
                arrayContent(i, 2) = "�ظ�����" & repeatRange & "��Ԫ��"
                errorRow = Right(repeatRange, Len(repeatRange) - 1)
                arrayContent(errorRow - 2, ��һ���ظ�����(repeatRange)) = "�ظ�����E" & currentRow & "��Ԫ��"
                Rows(errorRow).Font.ColorIndex = 4
                Rows(currentRow).Font.ColorIndex = 4
            Else
                dictionary(arraySource(i, 4)) = "E" & currentRow
            End If
            '�ж�λ��
            If Len(arraySource(i, 3)) = 7 And IsNumeric(arraySource(i, 3)) = True Then
                arrayContent(i, 5) = "�̶��绰"
            ElseIf preg.Test(arraySource(i, 3)) = True Then
                arrayContent(i, 5) = "�̶��绰"
            ElseIf Len(arraySource(i, 3)) = 11 And IsNumeric(arraySource(i, 3)) = True Then
                arrayContent(i, 5) = "�ƶ��绰"
            Else
                arrayContent(i, 5) = "λ���������ڷ�����"
                Rows(currentRow).Font.ColorIndex = 17
            End If
        End If
        '�绰3
        arraySource(i, 5) = CStr(arraySource(i, 5))
        If Len(arraySource(i, 5)) > 0 Then
            '�ظ���
            If dictionary.Exists(arraySource(i, 5)) Then
                repeatRange = dictionary(arraySource(i, 5))
                arrayContent(i, 3) = "�ظ�����" & repeatRange & "��Ԫ��"
                errorRow = Right(repeatRange, Len(repeatRange) - 1)
                arrayContent(errorRow - 2, ��һ���ظ�����(repeatRange)) = "�ظ�����F" & currentRow & "��Ԫ��"
                Rows(errorRow).Font.ColorIndex = 5
                Rows(currentRow).Font.ColorIndex = 5
            Else
                dictionary(arraySource(i, 5)) = "F" & currentRow
            End If
            '�ж�λ��
            If Len(arraySource(i, 3)) = 7 And IsNumeric(arraySource(i, 3)) = True Then
                arrayContent(i, 5) = "�̶��绰"
            ElseIf preg.Test(arraySource(i, 3)) = True Then
                arrayContent(i, 5) = "�̶��绰"
            ElseIf Len(arraySource(i, 3)) = 11 And IsNumeric(arraySource(i, 3)) = True Then
                arrayContent(i, 5) = "�ƶ��绰"
            Else
                arrayContent(i, 5) = "λ���������ڷ�����"
                Rows(currentRow).Font.ColorIndex = 17
            End If
        End If
        '����+��ַ�ظ���
        nameAddress = arraySource(i, 1) & arraySource(i, 2)
        If dictionary.Exists(nameAddress) Then
            repeatRange = dictionary(nameAddress)
            arrayContent(i, 4) = "�ظ�����" & repeatRange & "��"
            Rows(repeatRange).Font.ColorIndex = 9
            Rows(currentRow).Font.ColorIndex = 9
        Else
            dictionary(nameAddress) = currentRow
        End If
        '�����Ƿ�Ϊ��
        If Len(arraySource(i, 1)) < 1 Then
            arrayContent(i, 8) = "��"
            Rows(currentRow).Font.ColorIndex = 13
        End If
        '�����绰��Ϊ��
        If Len(arraySource(i, 3)) < 1 And Len(arraySource(i, 4)) < 1 And Len(arraySource(i, 5)) < 1 Then
            arrayContent(i, 9) = "��"
            Rows(currentRow).Font.ColorIndex = 11
        End If
        If isErrors(0) = True Or Len(arrayContent(i, 1)) > 0 Then isErrors(0) = True
        If isErrors(1) = True Or Len(arrayContent(i, 2)) > 0 Then isErrors(1) = True
        If isErrors(2) = True Or Len(arrayContent(i, 3)) > 0 Then isErrors(2) = True
        If isErrors(3) = True Or Len(arrayContent(i, 4)) > 0 Then isErrors(3) = True
        If isErrors(4) = True Or Len(arrayContent(i, 5)) <> 4 And Len(arrayContent(i, 5)) <> 0 Then isErrors(4) = True
        If isErrors(6) = True Or Len(arrayContent(i, 6)) <> 4 And Len(arrayContent(i, 6)) <> 0 Then isErrors(5) = True
        If isErrors(5) = True Or Len(arrayContent(i, 7)) <> 4 And Len(arrayContent(i, 7)) <> 0 Then isErrors(6) = True
        If isErrors(7) = True Or Len(arrayContent(i, 8)) > 0 Then isErrors(7) = True
        If isErrors(8) = True Or Len(arrayContent(i, 9)) > 0 Then isErrors(8) = True
    Next i
    Range("j2:r" & countRow) = ""
    Range("j2:r2") = arrayTableTitle
    Range("j3:r" & countRow) = arrayContent
    If isErrors(0) = False And isErrors(1) = False And isErrors(2) = False And isErrors(3) = False And isErrors(4) = False And isErrors(5) = False And isErrors(6) = False And isErrors(7) = False And isErrors(8) = False Then
        MsgBox "л��л��ȫ���޴���"
    Else
        If isErrors(0) = True Then MsgBox "��Ǹ���绰1���ظ��:("
        If isErrors(1) = True Then MsgBox "��Ǹ���绰2���ظ��:("
        If isErrors(2) = True Then MsgBox "��Ǹ���绰3���ظ��:("
        If isErrors(3) = True Then MsgBox "��Ǹ���������ַ���ظ���:("
        If isErrors(4) = True Then MsgBox "��Ǹ���绰1�д���:("
        If isErrors(5) = True Then MsgBox "��Ǹ���绰2�д���:("
        If isErrors(6) = True Then MsgBox "��Ǹ���绰3�д���:("
        If isErrors(7) = True Then MsgBox "��Ǹ����������:("
        If isErrors(8) = True Then MsgBox "��Ǹ�������绰��Ϊ�գ�:("
    End If
End Sub
Function ��һ���ظ�����(s As String)
    Dim errorColumn As String, contentColumn As Integer, errorRange As String
    errorColumn = Left(s, 1)
    Select Case errorColumn
        Case "D": contentColumn = 1
        Case "E": contentColumn = 2
        Case "F": contentColumn = 3
        Case Else: contentColumn = 0
    End Select
    ��һ���ظ����� = contentColumn
End Function
Sub ��������������()
    Dim i As Integer, countRow As Integer, arraySource, arrayContent, x As Integer, y As Integer
    countRow = Worksheets("����").[b65536].End(xlUp).Row
    If Worksheets.Count < 13 Then
        shequ = Array("����", "����", "���", "����Ӫ", "����", "��ɽ��", "��ɽ��", "�Ϻ�", "��ͨ", "�Ļ�Է", "ع��", "ع��")
        For i = 0 To i <= 11
            Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = shequ(i)
        Next i
    End If
    arrayTableTitle = Array("���", "��������", "��ͥ��ַ", "�绰1", "�绰2", "�绰3", "�ְ���", "����", "��ע")
    arraySource = Worksheets("����").Range("b3:i" & countRow)
    ReDim arrayContent(1 To 3000, 1 To 9)
    x = 1
    For i = 1 To UBound(arraySource)
        If arraySource(i, 7) <> arraySource(i + 1, 7) Then
            arrayContent(x, 1) = x
            arrayContent(x, 2) = arraySource(i, 1)
            arrayContent(x, 3) = arraySource(i, 2)
            arrayContent(x, 4) = arraySource(i, 3)
            arrayContent(x, 5) = arraySource(i, 4)
            arrayContent(x, 6) = arraySource(i, 5)
            arrayContent(x, 7) = arraySource(i, 6)
            arrayContent(x, 8) = arraySource(i, 7)
            arrayContent(x, 9) = arraySource(i, 8)
            With Worksheets(arraySource(i, 7))
                .Cells.Clear
                .Range("a1") = "������������绰��ǼǱ�" & arraySource(i, 7) & "������"
                .Range("a1:i1").MergeCells = True
                .Range("a2:i2") = arrayTableTitle
                .Range("a3:i" & x + 2) = arrayContent
                .Range("a2:i" & x + 2).Borders(xlEdgeTop).LineStyle = xlContinuous
                .Range("a2:i" & x + 2).Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Range("a2:i" & x + 2).Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Range("a2:i" & x + 2).Borders(xlEdgeRight).LineStyle = xlContinuous
                .Range("a2:i" & x + 2).Borders(xlInsideVertical).LineStyle = xlContinuous
                .Range("a2:i" & x + 2).Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .Cells.Columns.AutoFit
            End With
            ReDim arrayContent(1 To 3000, 1 To 9)
            x = 1
        Else
            arrayContent(x, 1) = x
            arrayContent(x, 2) = arraySource(i, 1)
            arrayContent(x, 3) = arraySource(i, 2)
            arrayContent(x, 4) = arraySource(i, 3)
            arrayContent(x, 5) = arraySource(i, 4)
            arrayContent(x, 6) = arraySource(i, 5)
            arrayContent(x, 7) = arraySource(i, 6)
            arrayContent(x, 8) = arraySource(i, 7)
            arrayContent(x, 9) = arraySource(i, 8)
            x = x + 1
            If countRow - 2 - i = 1 Then
                arrayContent(x, 1) = x
                arrayContent(x, 2) = arraySource(i + 1, 1)
                arrayContent(x, 3) = arraySource(i + 1, 2)
                arrayContent(x, 4) = arraySource(i + 1, 3)
                arrayContent(x, 5) = arraySource(i + 1, 4)
                arrayContent(x, 6) = arraySource(i + 1, 5)
                arrayContent(x, 7) = arraySource(i + 1, 6)
                arrayContent(x, 8) = arraySource(i + 1, 7)
                arrayContent(x, 9) = arraySource(i + 1, 8)
                With Worksheets(arraySource(i + 1, 7))
                    .Cells.Clear
                    .Range("a1") = "������������绰��ǼǱ�" & arraySource(i + 1, 7) & "������"
                    .Range("a1:i1").MergeCells = True
                    .Range("a2:i2") = arrayTableTitle
                    .Range("a3:i" & x + 2) = arrayContent
                    .Range("a2:i" & x + 2).Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Range("a2:i" & x + 2).Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Range("a2:i" & x + 2).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Range("a2:i" & x + 2).Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Range("a2:i" & x + 2).Borders(xlInsideVertical).LineStyle = xlContinuous
                    .Range("a2:i" & x + 2).Borders(xlInsideHorizontal).LineStyle = xlContinuous
                    .Cells.Columns.AutoFit
                End With
                Exit For
            End If
        End If
    Next i
    MsgBox "���Ƴɹ���"
End Sub
