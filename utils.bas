Attribute VB_Name = "utils"
Option Explicit
'///////////////////////////////////////////////////////////
'(汎用Function)自動計算等をストップ
'///////////////////////////////////////////////////////////
Public Sub appSet()
'処理中に、描画など余計なものを省略して高速化
    With Application
        .ScreenUpdating = False '描画を省略
        .Calculation = xlCalculationManual '手動計算
        .DisplayAlerts = False '警告を省略。
        .EnableEvents = False 'DisplayAlertsよりこちらを設定した方が良い？
    End With
End Sub

'///////////////////////////////////////////////////////////
'(汎用Function)自動計算等を再開
'///////////////////////////////////////////////////////////
Public Sub appReset()
'appSetをリセット
    With Application
        .ScreenUpdating = True '描画する
        .Calculation = xlCalculationAutomatic '自動計算
        .DisplayAlerts = True '警告を行う
        .EnableEvents = True
    End With
End Sub

'///////////////////////////////////////////////////////////
'(汎用Function)ファイルピッカー
'///////////////////////////////////////////////////////////
Public Function GeneralSelectFile(ByVal defaultPath As String, Optional ex As String) As Variant
    
    Dim Filter1 As String
    Dim Filter2 As String
    
    'exのオプション引数を受けて、それが"CSV"ならCSVでフィルタリングする
    If ex = "CSV" Then
        Filter1 = "CSVファイル"
        Filter2 = "*.csv"
    Else
        Filter1 = "MSエクセルファイル"
        Filter2 = "*.xls*"
    End If
        
    Dim SelectFileMessage As String
    SelectFileMessage = "複数ファイルはコントロールキーを押しながら選択してください"
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = True     '複数選択可
        .InitialFileName = defaultPath      '初期表示フォルダ
        .Filters.Add Filter1, Filter2, 1 '拡張子フィルター
        .Title = SelectFileMessage   'ダイアログタイトル"
        
        '◆選択後処理
        If .Show = True Then
            Dim Files As Variant
            ReDim Files(1 To 1)
            Dim SelectedItem As Variant
            
            For Each SelectedItem In .SelectedItems
                Files(UBound(Files)) = SelectedItem
                If UBound(Files) < .SelectedItems.Count Then
                    '配列を拡張
                    ReDim Preserve Files(1 To UBound(Files) + 1)
                End If
            Next SelectedItem
        Else
            .Execute    '処理(filedialogを初期化)
        End If
    End With
    
    GeneralSelectFile = Files
    
End Function

'///////////////////////////////////////////////////////////
'(汎用Function)連想配列内のValueの最大値を取得する
'///////////////////////////////////////////////////////////
Public Function GetMaxValueFromDictionary(ByVal TargetDict, ByVal ValueIndexNo As Long)
    Dim key As Variant
    Dim keys As Variant
    Dim arr As Variant
    Dim maxValue As Long
    
    keys = TargetDict.keys
    maxValue = TargetDict(keys(0))(ValueIndexNo) '最大値を仮置き
    
    '全てのKeyを比較にかけて最大値を更新する
    For Each key In keys
        arr = TargetDict(key)
        If arr(ValueIndexNo) > maxValue Then
            maxValue = arr(ValueIndexNo)
        End If
    Next key
    
    'return
    GetMaxValueFromDictionary = maxValue
End Function

'///////////////////////////////////////////////////////////
'(汎用Function)配列チェック
'背景：VBAには配列型がなく、Variant型を使うが、
'これが動的に型付けされるため配列にならないことがある。
'するとその後たとえばForループでLBound(arr)やarr(i)などを記述すると型エラーになる。
'
'対策：配列として使う意図だったVariant型変数への代入が単一値だったために
'配列にならなかった変数を配列化する
'///////////////////////////////////////////////////////////
Public Function VariantToArray(ByVal arr As Variant)
    If Not IsArray(arr) Then
        Dim tempArr As Variant
        ReDim tempArr(1 To 1)
        tempArr(1) = arr
        arr = tempArr
    End If
    '配列であれば何もせず、そのまま返す
    
    'return
    VariantToArray = arr
End Function

'///////////////////////////////////////////////////////////
'(汎用Function)特定のファイル名が今既に開いているブックにあるかどうか判定する
'///////////////////////////////////////////////////////////
Public Function isWorkbookOpen(ByVal fileName As String) As Boolean
    Dim wb As Workbook
    On Error Resume Next
    For Each wb In Application.Workbooks
        If StrComp(wb.name, fileName, vbTextCompare) = 0 Then
            'return
            isWorkbookOpen = True
            Exit Function
        End If
    Next wb
    
    'return
    isWorkbookOpen = False
End Function

'///////////////////////////////////////////////////////////
'//2次元配列を縦方向（行）に結合する関数
'// 引数：
'//   baseArr - 元の配列（Empty なら addArr をそのまま返す）
'//   addArr  - 追加する配列（列数が baseArr と同じである必要あり）
'// 戻り値：
'//   baseArr の末尾に addArr を連結した新しい配列
'//
'//例：
'//   baseArr = | A | B |
'//             | C | D |
'//
'//   addArr  = | E | F |
'//
'// ⇒ return  = | A | B |
'//             | C | D |
'//             | E | F |
'///////////////////////////////////////////////////////////
Public Function appendArray(ByVal baseArr As Variant, _
                            ByVal addArr As Variant) As Variant
    '//base が空なら add をそのまま返す
    If IsEmpty(baseArr) Then
        appendArray = addArr
        Exit Function
    End If
    
    Dim baseRows As Long: baseRows = UBound(baseArr, 1)
    Dim addRows  As Long:  addRows = UBound(addArr, 1)
    Dim colCount As Long: colCount = UBound(baseArr, 2)
    
    Dim mergedArr() As Variant
    ReDim mergedArr(1 To baseRows + addRows, 1 To colCount)
    
    Dim i As Long, j As Long
    
    '//baseをコピー
    For i = 1 To baseRows
        For j = 1 To colCount
            mergedArr(i, j) = baseArr(i, j)
        Next j
    Next i
    
    '//addをコピー（後半に追加）
    For i = 1 To addRows
        For j = 1 To colCount
            mergedArr(baseRows + i, j) = addArr(i, j)
        Next j
    Next i
    
    '//return
    appendArray = mergedArr
End Function

'単要素をレコードとして持っているCollectionを1次元配列に変換する関数
Public Function convertCollectionTo1dArray(ByVal records As Collection) As Variant
    Dim resultArr() As Variant: ReDim resultArr(1 To records.Count)
    Dim resultCt As Long: resultCt = 1
    
    Dim record As Variant
    For Each record In records
        resultArr(resultCt) = record '通常の変数の場合
'        Set resultArr(resultCt) = record 'オブジェクトの場合
        resultCt = resultCt + 1
    Next record
    
    'return
    convertCollectionTo1dArray = resultArr
End Function
'1次元配列をレコードとして持っているCollectionを2次元配列に変換する関数
Public Function convertCollectionToArray(ByVal records As Collection)
    Dim columnCt As Long: columnCt = UBound(records(1))
    Dim resultArr() As Variant: ReDim resultArr(1 To records.Count, 1 To columnCt)
    Dim resultCt As Long: resultCt = 1
    
    Dim record As Variant
    For Each record In records
        Dim recArr() As Variant: recArr = record
        Dim columnIndex As Long
        For columnIndex = 1 To columnCt
            resultArr(resultCt, columnIndex) = recArr(columnIndex)
        Next columnIndex
        resultCt = resultCt + 1
    Next record
    'return
    convertCollectionToArray = resultArr
End Function
'2次元配列を、1次元配列レコードを持つCollectionに変換する関数
Public Function convertArrayToCollection(ByVal arr As Variant) As Collection
    Dim ResultRecords As New Collection
    
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Dim record As Variant: ReDim record(1 To UBound(arr, 2))
        Dim j As Long
        For j = LBound(arr, 2) To UBound(arr, 2)
            record(j) = arr(i, j)
        Next j
        ResultRecords.Add record
    Next i
    'return
    Set convertArrayToCollection = ResultRecords
End Function
'2次元配列のある行からある行までを抜き出す汎用関数
Public Function extract2dArrayFrom2dArray(ByVal arr As Variant, _
                                          ByVal startRow As Long, _
                                          Optional ByVal finalRow As Long = 0) As Variant
    If finalRow = 0 Then finalRow = startRow
    
    Dim resultArr() As Variant
    ReDim resultArr(1 To finalRow - (startRow - 1), LBound(arr, 2) To UBound(arr, 2))
    Dim resultCt As Long: resultCt = resultCt + 1
    
    Dim i As Long
    For i = startRow To finalRow
        Dim j As Long
        For j = LBound(arr, 2) To UBound(arr, 2)
            resultArr(resultCt, j) = arr(i, j)
        Next j
        resultCt = resultCt + 1
    Next i
    
    'return
    extract2dArrayFrom2dArray = resultArr
End Function

'=============================================
' 2次元配列を指定列で高速ソート（クイックソート）
'
' 呼び出し方：
'   Sort2DByColumn arr, 2      ' 2列目をキーに昇順ソート
'
' 引数：
'   arr       : 2次元配列 (行×列)
'   sortCol   : ソート基準にする列番号（第2次元のインデックス）
'   各Optional: 再帰用なので、Call時は不使用
'=============================================
Public Function quickSort2dByColumn(ByRef arr As Variant, _
                                    ByVal sortCol As Long, _
                                    Optional ByVal order As String = "ASC", _
                                    Optional ByVal firstRow As Long = 0, _
                                    Optional ByVal lastRow As Long = 0, _
                                    Optional ByVal cL As Long = 0, _
                                    Optional ByVal cU As Long = 0) As Variant
    Dim rL As Long, rU As Long
    Dim i As Long, j As Long
    Dim pivot As Variant
    Dim tmp As Variant
    Dim col As Long
    Dim isRoot As Boolean

    '--------------------------------------
    ' この呼び出しが最上位かどうか判定
    '--------------------------------------
    isRoot = (firstRow = 0 And lastRow = 0)

    '--------------------------------------
    ' 初回呼び出しで初期化（検証もここで）
    '--------------------------------------
    If isRoot Then
        ' arr が空
        If IsEmpty(arr) Then
            Exit Function
        End If
        
        ' 2次元配列かどうか（LBound/UBoundの例外対策）
        On Error Resume Next
        rL = LBound(arr, 1): rU = UBound(arr, 1)
        cL = LBound(arr, 2): cU = UBound(arr, 2)
        If Err.Number <> 0 Then
            Err.Clear
            Exit Function
        End If
        On Error GoTo 0
        
        ' 要素数不足 / 列範囲外
        If rU - rL < 1 Then quickSort2dByColumn = arr: Exit Function
        If sortCol < cL Or sortCol > cU Then quickSort2dByColumn = arr: Exit Function
        
        ' ソート範囲設定
        firstRow = rL
        lastRow = rU
    End If

    '--------------------------------------
    ' クイックソート本体
    '--------------------------------------
    i = firstRow
    j = lastRow
    pivot = arr((firstRow + lastRow) \ 2, sortCol)

    Do While i <= j
        If UCase(order) = "ASC" Then
            ' 昇順：pivot より小さい間は i を進める
            Do While arr(i, sortCol) < pivot
                i = i + 1
            Loop
            ' 昇順：pivot より大きい間は j を戻す
            Do While arr(j, sortCol) > pivot
                j = j - 1
            Loop
        ElseIf UCase(order) = "DESC" Then
            ' 降順：pivot より大きい間は i を進める
            Do While arr(i, sortCol) > pivot
                i = i + 1
            Loop
            ' 降順：pivot より小さい間は j を戻す
            Do While arr(j, sortCol) < pivot
                j = j - 1
            Loop
        End If
        
        If i <= j Then
            If i <> j Then
                ' 行丸ごと入れ替え
                For col = cL To cU
                    tmp = arr(i, col)
                    arr(i, col) = arr(j, col)
                    arr(j, col) = tmp
                Next col
            End If
            i = i + 1
            j = j - 1
        End If
    Loop

    '--------------------------------------
    ' 左側の再帰
    '--------------------------------------
    If firstRow < j Then
        Call quickSort2dByColumn(arr, sortCol, order, firstRow, j, cL, cU)
    End If

    '--------------------------------------
    ' 右側の再帰
    '--------------------------------------
    If i < lastRow Then
        Call quickSort2dByColumn(arr, sortCol, order, i, lastRow, cL, cU)
    End If

    '--------------------------------------
    ' 最上位呼び出しのときだけ戻り値へ設定
    '--------------------------------------
    If isRoot Then
        quickSort2dByColumn = arr
    End If
End Function

'=============================================
'2つの1次元配列の同一性チェック
'=============================================
Public Function isSame1dArr(arrA As Variant, arrB As Variant) As Boolean
    isSame1dArr = False
    If LBound(arrA) <> LBound(arrB) Then Exit Function
    If UBound(arrA) <> UBound(arrB) Then Exit Function
    Dim i As Long
    For i = LBound(arrA) To UBound(arrB)
        If arrA(i) <> arrB(i) Then Exit Function
    Next i
    isSame1dArr = True
End Function

Public Function extractUnique(ByVal arr As Variant, ByVal idCol As Long) As Variant
    Dim dic As Object: Set dic = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If Not dic.exists(arr(i, idCol)) Then
            Dim rec As Variant: ReDim rec(LBound(arr, 2) To UBound(arr, 2))
            Dim j As Long
            For j = LBound(arr, 2) To UBound(arr, 2)
                 rec(j) = arr(i, j)
            Next j
            dic.Add arr(i, idCol), rec
        End If
    Next i
    
    Dim resultArr As Variant
    ReDim resultArr(LBound(arr, 1) To dic.Count, LBound(arr, 2) To UBound(arr, 2))
    
    Dim resultCt As Long: resultCt = 1
    Dim id As Variant
    For Each id In dic.keys
        For j = LBound(arr, 2) To UBound(arr, 2)
            resultArr(resultCt, j) = dic(id)(j)
        Next j
        resultCt = resultCt + 1
    Next id
    
    extractUnique = resultArr
End Function
