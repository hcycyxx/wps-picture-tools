Attribute VB_Name = "AmountWords"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal fmt As Long, ByVal memory As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal flags As Long, ByVal bytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal memory As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal memory As LongPtr) As Long
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal memory As LongPtr) As LongPtr
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As LongPtr, ByVal source As LongPtr, ByVal bytes As LongPtr)
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal fmt As Long, ByVal memory As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal flags As Long, ByVal bytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal memory As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal memory As Long) As Long
    Private Declare Function GlobalFree Lib "kernel32" (ByVal memory As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As Long, ByVal source As Long, ByVal bytes As Long)
#End If

Private mLastStatus As String
Private mPreviousStatus As Variant

Public Sub AmountUSD(Optional ByVal control As Object)
    CopyAmountWords "USD"
End Sub

Public Sub AmountRMB(Optional ByVal control As Object)
    CopyAmountWords "CNY"
End Sub

Public Sub AmountRMBEnglish(Optional ByVal control As Object)
    CopyAmountWords "CNY_EN"
End Sub

' Convert in Decimal arithmetic; VBA Round uses banker's rounding, which is
' unsuitable for the expected financial half-up rounding at the cent boundary.
Private Sub SplitAmount(ByVal amount As Variant, ByRef whole As Variant, ByRef cents As Long, ByRef negative As Boolean)
    Dim n As Variant, pennies As Variant
    If IsError(amount) Or IsEmpty(amount) Or IsNull(amount) Then Err.Raise vbObjectError + 1300, , "请选择有数字金额的单元格。"
    If VarType(amount) = vbBoolean Or Not IsNumeric(amount) Then Err.Raise vbObjectError + 1301, , "选中的内容不是数字金额。"
    n = CDec(amount)
    If Abs(n) > CDec("1000000000000") Then Err.Raise vbObjectError + 1302, , "金额绝对值不能超过 999,999,999,999.99。"
    pennies = Fix(Abs(n) * CDec(100) + CDec(0.5))
    If pennies > CDec("99999999999999") Then Err.Raise vbObjectError + 1302, , "四舍五入后的金额绝对值不能超过 999,999,999,999.99。"
    whole = Fix(pennies / CDec(100))
    cents = CLng(pennies - whole * CDec(100))
    negative = n < 0 And pennies > 0
End Sub

Private Function EnglishSmall(ByVal n As Long) As String
    Dim ones As Variant, tens As Variant
    ones = Split("|ONE|TWO|THREE|FOUR|FIVE|SIX|SEVEN|EIGHT|NINE|TEN|ELEVEN|TWELVE|THIRTEEN|FOURTEEN|FIFTEEN|SIXTEEN|SEVENTEEN|EIGHTEEN|NINETEEN", "|")
    tens = Split("||TWENTY|THIRTY|FORTY|FIFTY|SIXTY|SEVENTY|EIGHTY|NINETY", "|")
    If n < 20 Then
        EnglishSmall = ones(n)
    Else
        EnglishSmall = tens(n \ 10)
        If n Mod 10 > 0 Then EnglishSmall = EnglishSmall & "-" & ones(n Mod 10)
    End If
End Function

Private Function EnglishGroup(ByVal n As Long) As String
    If n >= 100 Then
        EnglishGroup = EnglishSmall(n \ 100) & " HUNDRED"
        If n Mod 100 > 0 Then EnglishGroup = EnglishGroup & " AND " & EnglishSmall(n Mod 100)
    Else
        EnglishGroup = EnglishSmall(n)
    End If
End Function

Private Function EnglishWhole(ByVal n As Variant) As String
    Dim group As Long, i As Long, divisor As Variant, remaining As Variant
    Dim scales As Variant, result As String
    If n = 0 Then EnglishWhole = "ZERO": Exit Function
    scales = Array("", " THOUSAND", " MILLION", " BILLION")
    remaining = n
    divisor = CDec(1000000000)
    For i = 3 To 0 Step -1
        group = CLng(Fix(remaining / divisor))
        remaining = remaining - CDec(group) * divisor
        If group > 0 Then
            If Len(result) > 0 Then
                If i = 0 And group < 100 Then result = result & " AND " Else result = result & " "
            End If
            result = result & EnglishGroup(group) & scales(i)
        End If
        divisor = divisor / CDec(1000)
    Next i
    EnglishWhole = result
End Function

Public Function KynkoUSDWords(ByVal amount As Variant) As String
    Dim whole As Variant, cents As Long, negative As Boolean, result As String
    SplitAmount amount, whole, cents, negative
    result = "SAY US DOLLARS "
    If negative Then result = result & "MINUS "
    result = result & EnglishWhole(whole)
    If cents > 0 Then
        result = result & " AND " & EnglishSmall(cents) & " CENT"
        If cents <> 1 Then result = result & "S"
    End If
    KynkoUSDWords = result & " ONLY"
End Function

Public Function KynkoCNYEnglishWords(ByVal amount As Variant) As String
    KynkoCNYEnglishWords = Replace(KynkoUSDWords(amount), "SAY US DOLLARS ", "SAY CNY ", 1, 1, vbBinaryCompare)
End Function

Private Function ChineseDigit(ByVal n As Long) As String
    ChineseDigit = Mid$("零壹贰叁肆伍陆柒捌玖", n + 1, 1)
End Function

Private Function ChineseGroup(ByVal n As Long) As String
    Dim result As String, pendingZero As Boolean, i As Long, d As Long
    Dim powers As Variant, units As Variant
    powers = Array(1, 10, 100, 1000)
    units = Array("", "拾", "佰", "仟")
    For i = 3 To 0 Step -1
        d = (n \ powers(i)) Mod 10
        If d > 0 Then
            If pendingZero Then result = result & "零"
            result = result & ChineseDigit(d) & units(i)
            pendingZero = False
        ElseIf Len(result) > 0 Then
            pendingZero = True
        End If
    Next i
    ChineseGroup = result
End Function

Private Function ChineseWhole(ByVal n As Variant) As String
    Dim group As Long, i As Long, divisor As Variant, remaining As Variant
    Dim scales As Variant, result As String, skipped As Boolean
    If n = 0 Then ChineseWhole = "零": Exit Function
    scales = Array("", "万", "亿")
    remaining = n
    divisor = CDec(100000000)
    For i = 2 To 0 Step -1
        group = CLng(Fix(remaining / divisor))
        remaining = remaining - CDec(group) * divisor
        If group > 0 Then
            If Len(result) > 0 And (group < 1000 Or skipped) Then result = result & "零"
            result = result & ChineseGroup(group) & scales(i)
            skipped = False
        ElseIf Len(result) > 0 Then
            skipped = True
        End If
        divisor = divisor / CDec(10000)
    Next i
    ChineseWhole = result
End Function

Public Function KynkoRMBWords(ByVal amount As Variant) As String
    Dim whole As Variant, cents As Long, negative As Boolean, result As String
    Dim jiao As Long, fen As Long
    SplitAmount amount, whole, cents, negative
    If negative Then result = "负"
    result = result & "人民币" & ChineseWhole(whole) & "元"
    If cents = 0 Then
        result = result & "整"
    Else
        jiao = cents \ 10: fen = cents Mod 10
        If jiao > 0 Then result = result & ChineseDigit(jiao) & "角"
        If fen > 0 Then
            If jiao = 0 And whole > 0 Then result = result & "零"
            result = result & ChineseDigit(fen) & "分"
        End If
    End If
    KynkoRMBWords = result
End Function

Private Sub CopyUnicodeText(ByVal text As String)
    On Error GoTo Failed
#If VBA7 Then
    Dim memory As LongPtr, pointer As LongPtr
#Else
    Dim memory As Long, pointer As Long
#End If
    Dim opened As Boolean, problem As Long, description As String
    memory = GlobalAlloc(&H42, LenB(text) + 2)
    If memory = 0 Then Err.Raise vbObjectError + 1320, , "无法分配剪贴板内存。"
    pointer = GlobalLock(memory)
    If pointer = 0 Then Err.Raise vbObjectError + 1321, , "无法访问剪贴板内存。"
    CopyMemory pointer, StrPtr(text), LenB(text)
    GlobalUnlock memory
    pointer = 0
    If OpenClipboard(Application.hwnd) = 0 Then Err.Raise vbObjectError + 1322, , "剪贴板正在被其他程序使用，请稍后再点一次。"
    opened = True
    If EmptyClipboard() = 0 Then Err.Raise vbObjectError + 1323, , "无法更新剪贴板，请稍后重试。"
    If SetClipboardData(13, memory) = 0 Then Err.Raise vbObjectError + 1324, , "复制失败，请稍后再点一次。"
    memory = 0
    CloseClipboard
    Exit Sub
Failed:
    problem = Err.Number: description = Err.Description
    On Error Resume Next
    If pointer <> 0 Then GlobalUnlock memory
    If opened Then CloseClipboard
    If memory <> 0 Then GlobalFree memory
    On Error GoTo 0
    Err.Raise problem, , description
End Sub

Public Sub ClearAmountCopyStatus()
    On Error Resume Next
    If CStr(Application.StatusBar) = mLastStatus Then Application.StatusBar = mPreviousStatus
End Sub

Private Sub CopyAmountWords(ByVal currencyCode As String)
    On Error GoTo Failed
    Dim chosen As Range, cell As Range, words As String, title As String
    title = "美元大写": If currencyCode = "CNY" Then title = "人民币大写"
    If currencyCode = "CNY_EN" Then title = "人民币英文大写"
    If Application.ActiveWorkbook Is Nothing Then Exit Sub
    If Application.ActiveWorkbook.IsAddin Then Exit Sub
    If TypeName(Application.Selection) <> "Range" Then Err.Raise vbObjectError + 1312, , "请先选中表格里的数字金额，再点按钮。"
    Set chosen = Application.Selection
    If chosen.Areas.Count <> 1 Then Err.Raise vbObjectError + 1310, , "请只选择一个金额单元格。"
    Set cell = chosen.Cells(1, 1)
    If chosen.Rows.Count > 1 Or chosen.Columns.Count > 1 Then
        If chosen.Address <> cell.MergeArea.Address Then Err.Raise vbObjectError + 1311, , "请只选择一个金额单元格，不要选择整行、整列或多个单元格。"
    End If
    Select Case currencyCode
        Case "USD": words = KynkoUSDWords(cell.Value2)
        Case "CNY_EN": words = KynkoCNYEnglishWords(cell.Value2)
        Case Else: words = KynkoRMBWords(cell.Value2)
    End Select
    CopyUnicodeText words
    On Error Resume Next
    If CStr(Application.StatusBar) <> mLastStatus Then mPreviousStatus = Application.StatusBar
    mLastStatus = "金额已复制：" & Left$(words, 160) & "（Ctrl+V 粘贴）"
    Application.StatusBar = mLastStatus
    Application.OnTime Now + TimeSerial(0, 0, 5), "'" & ThisWorkbook.Name & "'!ClearAmountCopyStatus"
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, title
End Sub
