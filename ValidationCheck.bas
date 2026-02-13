Attribute VB_Name = "ValidationCheck"
Option Explicit
Option Base 1

' OKなら ""、NGならエラーメッセージを返す
Public Function apiValidate(ByVal value As String, _
                            ByVal rules As Variant, _
                            Optional ByVal fieldLabel As String = "") As String
    Dim i As Long
    For i = LBound(rules) To UBound(rules)
        Dim rule As String: rule = Trim$(rules(i))
        If rule <> "" Then
            Dim msg As String: msg = ValidateByRule(value, rule, fieldLabel)
            If msg <> "" Then
                apiValidate = msg
                Exit Function
            End If
        End If
    Next i
End Function

' 1ルール適用：OKなら ""、NGならメッセージ
Private Function ValidateByRule(ByVal value As String, _
                                ByVal rule As String, _
                                ByVal fieldLabel As String) As String
    Select Case LCase$(rule)
        Case "required"
            If Not Rule_Required(value) Then
                ValidateByRule = getErrMsg(fieldLabel, "入力必須です。")
            End If
        Case "digits6"
            If Not Rule_DigitsN(value, 6) Then
                ValidateByRule = getErrMsg(fieldLabel, "6桁の数字で入力してください。")
            End If
        Case "numeric"
            If Not Rule_Numeric(value) Then
                ValidateByRule = getErrMsg(fieldLabel, "半角数字で入力してください。")
            End If
        Case "yymmdd"
            If Not Rule_YYMMDD(value) Then
                ValidateByRule = getErrMsg(fieldLabel, "YYMMDD形式で入力してください。")
            End If
        Case Else
            '未定義ルールはひとまず無視
            ValidateByRule = "未定義ルール: " & rule
    End Select
End Function

Private Function getErrMsg(ByVal fieldLabel As String, _
                           ByVal body As String) As String
    If Len(fieldLabel) > 0 Then
        getErrMsg = fieldLabel & " Error ：" & body
    Else
        getErrMsg = body
    End If
End Function

' ---- ルール関数群（できるだけ純粋に） ----
Private Function Rule_Required(ByVal s As String) As Boolean
    Rule_Required = (Trim$(s) <> "")
End Function
Private Function Rule_Numeric(ByVal s As String) As Boolean
    Rule_Numeric = IsNumeric(s)
End Function
Private Function Rule_DigitsN(ByVal s As String, _
                              ByVal n As Long) As Boolean
    If Len(s) <> n Then Exit Function
    Rule_DigitsN = (s Like String(n, "#"))
End Function
Private Function Rule_YYMMDD(ByVal s As String) As Boolean
    ' 前提：digits6済みでもよいが、単体でも安全に判定できるようにしておく
    If Len(s) <> 6 Then Exit Function
    If Not (s Like "######") Then Exit Function

    Dim yy As Long, mm As Long, dd As Long
    yy = CInt(Left$(s, 2))
    mm = CInt(Mid$(s, 3, 2))
    dd = CInt(Right$(s, 2))

    If mm < 1 Or mm > 12 Then Exit Function
    If dd < 1 Then Exit Function

    Dim maxD As Long
    maxD = Day(DateSerial(2000 + yy, mm + 1, 0))
    Rule_YYMMDD = (dd <= maxD)
End Function
