Attribute VB_Name = "FloatPictures"
Option Explicit

' Selection only. The user invokes WPS's context menu and saves manually.
Public Sub ConvertAllPictures(Optional ByVal control As Object)
    On Error GoTo Failed
    Dim ws As Worksheet, formulas As Range, c As Range, chosen As Range
    Dim seen As Object, matcher As Object, area As Range, address As String, selectedCount As Long
    If Application.ActiveWorkbook Is Nothing Then Exit Sub
    If Not TypeOf Application.ActiveSheet Is Worksheet Then Exit Sub
    Set ws = Application.ActiveSheet
    Set seen = CreateObject("Scripting.Dictionary")
    Set matcher = CreateObject("VBScript.RegExp")
    matcher.Pattern = "^\s*=\s*@?(_xlfn\.)?DISPIMG\s*\("
    matcher.IgnoreCase = True
    On Error Resume Next
    Set formulas = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo Failed
    If Not formulas Is Nothing Then
        For Each c In formulas.Cells
            If matcher.Test(CStr(c.Formula)) Then
                Set area = c.MergeArea
                address = area.Address
                If Not seen.Exists(address) Then
                    seen.Add address, True
                    If chosen Is Nothing Then
                        Set chosen = area
                    Else
                        Set chosen = Application.Union(chosen, area)
                    End If
                    selectedCount = selectedCount + 1
                End If
            End If
        Next c
    End If
    If chosen Is Nothing Then
        ws.Cells(1, 1).Select
        MsgBox "当前工作表没有可选的嵌入图片单元格。", vbInformation, "选择图片"
        Exit Sub
    End If
    chosen.Select
    Application.StatusBar = "已选中 " & selectedCount & " 个嵌入图片单元格。请右键选择切换为浮动图片；程序未转换、未保存。"
    Exit Sub
Failed:
    MsgBox "未能选中图片：" & Err.Description, vbExclamation, "选择图片"
End Sub

Private Function ImageTarget(ByVal sh As Shape, ByVal ws As Worksheet) As Range
    Dim anchor As Range, j As Long
    Set anchor = sh.TopLeftCell
    ' WPS can report the preceding cell at a rounded boundary.
    For j = 1 To 4
        If anchor.Height > 0 And sh.Top >= anchor.Top + anchor.Height - 0.25 Then
            Set anchor = ws.Cells(anchor.Row + 1, anchor.Column)
        Else
            Exit For
        End If
    Next j
    For j = 1 To 4
        If anchor.Width > 0 And sh.Left >= anchor.Left + anchor.Width - 0.25 Then
            Set anchor = ws.Cells(anchor.Row, anchor.Column + 1)
        Else
            Exit For
        End If
    Next j
    Set ImageTarget = anchor.MergeArea.Cells(1, 1)
End Function

Private Function Occupied(ByVal target As Range) As Boolean
    Occupied = True
    If target.HasFormula Then Exit Function
    If IsError(target.Value2) Then Exit Function
    If IsNull(target.Value2) Then Exit Function
    Occupied = Len(CStr(target.Value2)) > 0
End Function

Public Sub EmbedAllPictures(Optional ByVal control As Object)
    On Error GoTo Failed
    Dim ws As Worksheet, sh As Shape, target As Range, address As String
    Dim targets As Object, candidates As Object, key As Variant
    Dim names() As Variant, selectedCount As Long, skipped As Long, duplicates As Long
    If Application.ActiveWorkbook Is Nothing Then Exit Sub
    If Not TypeOf Application.ActiveSheet Is Worksheet Then Exit Sub
    Set ws = Application.ActiveSheet
    Set targets = CreateObject("Scripting.Dictionary")
    Set candidates = CreateObject("Scripting.Dictionary")
    For Each sh In ws.Shapes
        If sh.Type = 13 Or sh.Type = 11 Then
            Set target = ImageTarget(sh, ws)
            If Occupied(target) Then
                skipped = skipped + 1
            Else
                address = target.Address
                candidates.Add sh.Name, address
                If targets.Exists(address) Then
                    targets(address) = targets(address) + 1
                Else
                    targets.Add address, 1
                End If
            End If
        End If
    Next sh
    If candidates.Count > 0 Then ReDim names(1 To candidates.Count)
    For Each key In candidates.Keys
        If targets(candidates(key)) = 1 Then
            selectedCount = selectedCount + 1
            names(selectedCount) = CStr(key)
        Else
            duplicates = duplicates + 1
        End If
    Next key
    If selectedCount = 0 Then
        ' Clear any prior object selection so it cannot be mistaken for this result.
        ws.Cells(1, 1).Select
        MsgBox "当前工作表没有可选的浮动图片。已跳过 " & skipped & " 张目标有内容、" & duplicates & " 张目标重叠的图片。", vbInformation, "选择图片"
        Exit Sub
    End If
    ReDim Preserve names(1 To selectedCount)
    ws.Shapes.Range(names).Select
    Application.StatusBar = "已选中 " & selectedCount & " 张浮动图片；跳过 " & skipped & " 张目标有内容、" & duplicates & " 张目标重叠的图片。请右键嵌入；程序未转换、未保存。"
    Exit Sub
Failed:
    MsgBox "未能选中图片：" & Err.Description, vbExclamation, "选择图片"
End Sub
