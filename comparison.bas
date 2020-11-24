Attribute VB_Name = "电话库比对模块"
Sub 检查重复及错误()
    If Len(Range("b3")) = 0 And Len(Range("c3")) = 0 And Len(Range("d3")) = 0 And Len(Range("e3")) = 0 And Len(Range("f3")) = 0 Then
        MsgBox "请正确导入文件！"
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
    '数值 9、10 和 13 可以分别转换为制表符、换行符和回车符
    Cells.Replace What:=" ", Replacement:=""
    Cells.Replace What:=Chr(9), Replacement:=""
    Cells.Replace What:=Chr(10), Replacement:=""
    Cells.Replace What:=Chr(13), Replacement:=""
    arraySource = Range("b3:f" & countRow)
    ReDim arrayContent(1 To countRow, 1 To 9)
    Columns.Font.ColorIndex = 0
    isErrors = Array(False, False, False, False, False, False, False, False, False)
    arrayTableTitle = Array("电话1重复数", "电话2重复数", "电话3重复数", "姓名+地址重复数", "电话1位数", "电话2位数", "电话3位数", "姓名是否为空", "三个电话都为空")
    For i = 1 To UBound(arraySource)
        currentRow = i + 2
        '电话1
        arraySource(i, 3) = CStr(arraySource(i, 3))
        If Len(arraySource(i, 3)) > 0 Then
            '重复项
            If dictionary.Exists(arraySource(i, 3)) Then
                repeatRange = dictionary(arraySource(i, 3))
                arrayContent(i, 1) = "重复项在" & repeatRange & "单元格"
                errorRow = Right(repeatRange, Len(repeatRange) - 1)
                arrayContent(errorRow - 2, 第一次重复列数(repeatRange)) = "重复项在D" & currentRow & "单元格"
                Rows(errorRow).Font.ColorIndex = 3
                Rows(currentRow).Font.ColorIndex = 3
            Else
                dictionary(CStr(arraySource(i, 3))) = "D" & currentRow
            End If
            '判断位数
            If Len(arraySource(i, 3)) = 7 And IsNumeric(arraySource(i, 3)) = True Then
                arrayContent(i, 5) = "固定电话"
            ElseIf preg.Test(arraySource(i, 3)) = True Then
                arrayContent(i, 5) = "固定电话"
            ElseIf Len(arraySource(i, 3)) = 11 And IsNumeric(arraySource(i, 3)) = True Then
                arrayContent(i, 5) = "移动电话"
            Else
                arrayContent(i, 5) = "位数错误或存在非数字"
                Rows(currentRow).Font.ColorIndex = 17
            End If
        End If
        '电话2
        arraySource(i, 4) = CStr(arraySource(i, 4))
        If Len(arraySource(i, 4)) > 0 Then
            '重复项
            If dictionary.Exists(arraySource(i, 4)) Then
                repeatRange = dictionary(arraySource(i, 4))
                arrayContent(i, 2) = "重复项在" & repeatRange & "单元格"
                errorRow = Right(repeatRange, Len(repeatRange) - 1)
                arrayContent(errorRow - 2, 第一次重复列数(repeatRange)) = "重复项在E" & currentRow & "单元格"
                Rows(errorRow).Font.ColorIndex = 4
                Rows(currentRow).Font.ColorIndex = 4
            Else
                dictionary(arraySource(i, 4)) = "E" & currentRow
            End If
            '判断位数
            If Len(arraySource(i, 3)) = 7 And IsNumeric(arraySource(i, 3)) = True Then
                arrayContent(i, 5) = "固定电话"
            ElseIf preg.Test(arraySource(i, 3)) = True Then
                arrayContent(i, 5) = "固定电话"
            ElseIf Len(arraySource(i, 3)) = 11 And IsNumeric(arraySource(i, 3)) = True Then
                arrayContent(i, 5) = "移动电话"
            Else
                arrayContent(i, 5) = "位数错误或存在非数字"
                Rows(currentRow).Font.ColorIndex = 17
            End If
        End If
        '电话3
        arraySource(i, 5) = CStr(arraySource(i, 5))
        If Len(arraySource(i, 5)) > 0 Then
            '重复项
            If dictionary.Exists(arraySource(i, 5)) Then
                repeatRange = dictionary(arraySource(i, 5))
                arrayContent(i, 3) = "重复项在" & repeatRange & "单元格"
                errorRow = Right(repeatRange, Len(repeatRange) - 1)
                arrayContent(errorRow - 2, 第一次重复列数(repeatRange)) = "重复项在F" & currentRow & "单元格"
                Rows(errorRow).Font.ColorIndex = 5
                Rows(currentRow).Font.ColorIndex = 5
            Else
                dictionary(arraySource(i, 5)) = "F" & currentRow
            End If
            '判断位数
            If Len(arraySource(i, 3)) = 7 And IsNumeric(arraySource(i, 3)) = True Then
                arrayContent(i, 5) = "固定电话"
            ElseIf preg.Test(arraySource(i, 3)) = True Then
                arrayContent(i, 5) = "固定电话"
            ElseIf Len(arraySource(i, 3)) = 11 And IsNumeric(arraySource(i, 3)) = True Then
                arrayContent(i, 5) = "移动电话"
            Else
                arrayContent(i, 5) = "位数错误或存在非数字"
                Rows(currentRow).Font.ColorIndex = 17
            End If
        End If
        '姓名+地址重复数
        nameAddress = arraySource(i, 1) & arraySource(i, 2)
        If dictionary.Exists(nameAddress) Then
            repeatRange = dictionary(nameAddress)
            arrayContent(i, 4) = "重复项在" & repeatRange & "行"
            Rows(repeatRange).Font.ColorIndex = 9
            Rows(currentRow).Font.ColorIndex = 9
        Else
            dictionary(nameAddress) = currentRow
        End If
        '姓名是否为空
        If Len(arraySource(i, 1)) < 1 Then
            arrayContent(i, 8) = "空"
            Rows(currentRow).Font.ColorIndex = 13
        End If
        '三个电话都为空
        If Len(arraySource(i, 3)) < 1 And Len(arraySource(i, 4)) < 1 And Len(arraySource(i, 5)) < 1 Then
            arrayContent(i, 9) = "空"
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
        MsgBox "谢天谢地全部无错误！"
    Else
        If isErrors(0) = True Then MsgBox "抱歉！电话1有重复项！:("
        If isErrors(1) = True Then MsgBox "抱歉！电话2有重复项！:("
        If isErrors(2) = True Then MsgBox "抱歉！电话3有重复项！:("
        If isErrors(3) = True Then MsgBox "抱歉！姓名与地址有重复！:("
        If isErrors(4) = True Then MsgBox "抱歉！电话1有错误！:("
        If isErrors(5) = True Then MsgBox "抱歉！电话2有错误！:("
        If isErrors(6) = True Then MsgBox "抱歉！电话3有错误！:("
        If isErrors(7) = True Then MsgBox "抱歉！无姓名！:("
        If isErrors(8) = True Then MsgBox "抱歉！三个电话都为空！:("
    End If
End Sub
Function 第一次重复列数(s As String)
    Dim errorColumn As String, contentColumn As Integer, errorRange As String
    errorColumn = Left(s, 1)
    Select Case errorColumn
        Case "D": contentColumn = 1
        Case "E": contentColumn = 2
        Case "F": contentColumn = 3
        Case Else: contentColumn = 0
    End Select
    第一次重复列数 = contentColumn
End Function
Sub 复制社区工作表()
    Dim i As Integer, countRow As Integer, arraySource, arrayContent, x As Integer, y As Integer
    countRow = Worksheets("汇总").[b65536].End(xlUp).Row
    If Worksheets.Count < 13 Then
        shequ = Array("大海阳", "德新", "翡翠", "海防营", "海港", "黄山北", "黄山南", "南洪", "南通", "文化苑", "毓东", "毓西")
        For i = 0 To i <= 11
            Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = shequ(i)
        Next i
    End If
    arrayTableTitle = Array("序号", "户主姓名", "家庭地址", "电话1", "电话2", "电话3", "分包人", "社区", "备注")
    arraySource = Worksheets("汇总").Range("b3:i" & countRow)
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
                .Range("a1") = "“评社区”活动电话库登记表（" & arraySource(i, 7) & "社区）"
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
                    .Range("a1") = "“评社区”活动电话库登记表（" & arraySource(i + 1, 7) & "社区）"
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
    MsgBox "复制成功！"
End Sub
