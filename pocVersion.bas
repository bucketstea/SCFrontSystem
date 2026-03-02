Attribute VB_Name = "pocVersion"
Option Explicit
Option Base 1

'移行期間用の切り戻し処理モジュール
'安定版リリース時に削除予定

Private OriginSh As Worksheet

Public Function isRollbacked() As Boolean
    Dim Sh As Worksheet
    For Each Sh In ThisWorkbook.Sheets
        If InStr(Sh.name, "入力シートrev") > 0 Then
            isRollbacked = True
            Exit Function
        End If
    Next Sh
    
    isRollbacked = False
End Function

Public Sub rollback()
    Call DispMod.init
    
    '原版シート特定
    Call checkOriginSh
    
    'データ転記
    Call rewriteTable
    
    'Home画面表示機能をOFF
    On Error Resume Next
    Application.CommandBars("Cell").Controls("Home画面を表示").Delete
    On Error GoTo 0
    
    'シート名置換
    Call swapSh
    
    '全フォームオブジェクトをクリア
    Call DispMod.clearingForms
    
    '完了通知
    MsgBox "原版の入力シートを復帰しました。以降、入力フォームではなく、従来の方法でシートへ入力してください。" & vbCrLf & _
           "(入力フォームは後日改修します。)", _
           vbExclamation, "ロールバック済み"
End Sub
Private Sub checkOriginSh()
    Dim Sh As Worksheet
    For Each Sh In ThisWorkbook.Sheets
        If InStr(Sh.name, "origin") > 0 Then
            Set OriginSh = ThisWorkbook.Sheets("origin")
            Exit For
        End If
    Next Sh
    If OriginSh Is Nothing Then
        MsgBox "originシートがないためロールバック処理ができません。", vbCritical, "ロールバック中止"
        End
    End If
End Sub
Private Sub rewriteTable()
    'シート内容を取得
    With InputSh
        Dim newArrFinalR As Long: newArrFinalR = .Cells(Rows.Count, COL_DATE).End(xlUp).Row
        Dim newArr As Variant
        newArr = .Range(.Cells(1, COL_A), .Cells(newArrFinalR, COL_LAST)).Value2
    End With
    With OriginSh
        Dim oldArrFinalR As Long: oldArrFinalR = OriginSh.Cells(Rows.Count, COL_DATE).End(xlUp).Row
        Dim oldArr As Variant
        oldArr = OriginSh.Range(.Cells(1, COL_A), .Cells(oldArrFinalR, COL_LAST)).Formula2
    End With
    
    '新規追加するデータを取得
    Dim addStartRow As Long: addStartRow = UBound(oldArr, 1) + 1
    Dim addFinalRow As Long: addFinalRow = UBound(newArr, 1)
    If addStartRow > addFinalRow Then
        MsgBox "追加分データなし", vbCritical, "中止"
        End
    End If
    Dim addArr As Variant: addArr = extract2dArrayFrom2dArray(newArr, addStartRow, addFinalRow)
    
    '時刻表示と0始まり番号対策
    addArr = formattingArr(addArr)
    
    '新規追加分を反映した原版配列を生成
    Dim resultArr As Variant: resultArr = appendArray(oldArr, addArr)
    
    With OriginSh
        .Unprotect "042595"
        .Range(.Cells(1, COL_A), .Cells(UBound(resultArr, 1), COL_LAST)).value = resultArr
        .Protect "042595"
    End With
End Sub
Private Function formattingArr(arr) As Variant
    Dim resultArr As Variant
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        arr(i, COL_TEL) = CStr(arr(i, COL_TEL))
        arr(i, COL_TIME) = Format(arr(i, COL_TIME), "hh:mm")
    Next i
    formattingArr = arr
End Function
Private Sub swapSh()
    InputSh.name = "入力シートrev"
    InputSh.Visible = xlSheetHidden
    OriginSh.name = "入力シート"
    OriginSh.Visible = xlSheetVisible
End Sub
