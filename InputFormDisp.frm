VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputFormDisp 
   Caption         =   "データ入力"
   ClientHeight    =   9012.001
   ClientLeft      =   -3168
   ClientTop       =   -12936
   ClientWidth     =   22824
   OleObjectBlob   =   "InputFormDisp.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "InputFormDisp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Base 1

'/////////////////////////////////////////////////////////// Windows API の宣言
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" _
    (ByVal hInstance As LongPtr, ByVal lpCursorName As Long) As LongPtr
Private Declare PtrSafe Function SetCursor Lib "user32" _
    (ByVal hCursor As LongPtr) As LongPtr
Private Const IDC_HAND = 32649
'///////////////////////////////////////////////////////////

Private isEditMode As Boolean
Private generalid As String
Private preLoadData As Variant
Private previousDisp As Object

'///////////////////////////////////////////////////////////
'初期化
'///////////////////////////////////////////////////////////
Private Sub UserForm_Initialize()
    'UIの見た目（サイズや位置などの見た目要素）を設定
    Call UiConfig_InputFormDisp.configUiDesign(Me)
    Call UiConfig_InputFormDisp.configComboBox(Me) 'ComboBox設定
End Sub

'///////////////////////////////////////////////////////////
'再表示更新
'///////////////////////////////////////////////////////////
Public Sub reloadInputs(ByVal targetDate As Variant, _
                        ByVal listCt As Long, _
                        Optional ByVal targetName As String = "", _
                        Optional ByVal targetTel As String = "")
    
    Call UiConfig_InputFormDisp.configUiDesign(Me)
    Call clearAllValue
    
    isEditMode = False
    ImageDelete.Visible = False
    Me.Caption = "新規データ入力画面"
    
    Dim shLastRow As Long: shLastRow = InputSh.Cells(InputSh.Rows.Count, 1).End(xlUp).Row
    generalid = shLastRow + 1
    LabelEditId.Caption = "ID: " & listCt + 1 & "_" & generalid
    
    TextBoxDate.Text = Format(targetDate, "yymmdd")
    If targetName <> "" Then TextBoxName.Text = targetName
    If targetTel <> "" Then TextBoxTel.Text = targetTel
    
    'プライマリアクション(新規追加)ボタン
    CommandButtonSave.Caption = "Add"
    CommandButtonSave.BackColor = RGB(68, 114, 196)
    
    '入力状態プリロード
    preLoadData = collectInputData()
    
    'プレースホルダ設定
    Call switchPlaceholder(Me)
    
    ComboBoxAd.SetFocus
End Sub
Public Sub editInputs(ByVal listRow As Variant, _
                      ByVal dailyId As Long, _
                      ByVal id As String)
    Call UiConfig_InputFormDisp.configUiDesign(Me)
    Call clearAllValue
    
    generalid = id
    isEditMode = True
    Me.Caption = "既存データ編集画面"
    
    '選択レコードを入力欄に反映する
    Call currentRecordToInputs(listRow, dailyId, generalid)
    
    'プライマリアクション(変更保存)ボタン
    CommandButtonSave.Caption = "Save changes!"
    CommandButtonSave.BackColor = RGB(237, 125, 49)
    
    '入力状態プリロード
    preLoadData = collectInputData()
    
    'プレースホルダ設定
    Call switchPlaceholder(Me)
    
    ComboBoxAd.SetFocus
End Sub

'ListViewのレコードに基づいて入力欄各項目にデータを反映する
Public Sub currentRecordToInputs(ByVal listRow As Variant, ByVal dailyId As Long, ByVal generalid As String)
    With Me
        .ImageDelete.Visible = True
        .LabelEditId.Caption = "ID: " & dailyId & "_" & generalid
        .TextBoxDate.Text = listRow(COL_DATE)
        .ComboBoxAd.Text = listRow(COL_ROOT)
        .ComboBoxType.Text = listRow(COL_TYPE)
        .TextBoxTime.Text = listRow(COL_TIME)
        .TextBoxName.Text = listRow(COL_CUSTM)
        .TextBoxTel.Text = CStr(listRow(COL_TEL)) '0始まり崩れ防止
        .TextBoxNG.Text = listRow(COL_NG)
        .TextBoxNotes.Text = listRow(COL_NOTE)
        .TextBoxCast.Text = listRow(COL_STAFF)
        .TextBoxCourse.Text = listRow(COL_COURS)
        .TextBoxService.Text = listRow(COL_SERV)
        .TextBoxOP.Text = listRow(COL_OP)
        .TextBoxDestination.Text = listRow(COL_DEST)
        .TextBoxExpand.Text = listRow(COL_EXPAN)
        .TextBoxSales.Text = listRow(COL_SALES)
        .TextBoxCost.Text = listRow(COL_CCOST)
        .TextBoxProfit.Text = listRow(COL_PROFI)
        .TextBoxQB.Text = listRow(COL_QBACK)
        .TextBoxSB.Text = listRow(COL_SBACK)
    End With
End Sub
'入力内容を1次元配列にまとめる
Private Function collectInputData() As Variant
    Dim Data(1 To COL_LAST) As Variant ' 列数に応じて変更
    
    Data(COL_A) = generalid
    Data(COL_COUNT) = ""
    Data(COL_DATE) = TextBoxDate.Text
    Data(COL_NEW) = ""
    Data(COL_ROOT) = ComboBoxAd.Text
    Data(COL_TYPE) = ComboBoxType.Text
    Data(COL_STAFF) = TextBoxCast.Text
    Data(COL_CUSTM) = TextBoxName.Text
    Data(COL_TEL) = CStr(TextBoxTel.Text) '0始まり崩れの防止
    Data(COL_NG) = TextBoxNG.Text
    Data(COL_NOTE) = TextBoxNotes.Text
    Data(COL_DEST) = TextBoxDestination.Text
    Data(COL_SERV) = TextBoxService.Text
    Data(COL_COURS) = TextBoxCourse.Text
    Data(COL_EXPAN) = TextBoxExpand.Text
    Data(COL_OP) = TextBoxOP.Text
    Data(COL_TIME) = TextBoxTime.Text
    Data(COL_SALES) = TextBoxSales.Text
    Data(COL_CCOST) = TextBoxCost.Text
    Data(COL_PROFI) = TextBoxProfit.Text
    Data(COL_QBACK) = TextBoxQB.Text
    Data(COL_SBACK) = TextBoxSB.Text
    
    collectInputData = Data
End Function

'///////////////////////////////////////////////////////////
'遷移系ボタン操作
'///////////////////////////////////////////////////////////
Private Sub LabelBack_Click()
    Dim currentData As Variant: currentData = collectInputData()
    If Not isSame1dArr(preLoadData, currentData) Then
        If (MsgBox("You have changes that haven't been saved yet." & vbCrLf & _
                   "Would you like to back to the previous screen?", vbExclamation + vbYesNo, "Caution") = vbNo) Then
            Exit Sub
        End If
    End If
    
    Me.Hide
    goBack
End Sub
Private Sub ImageDelete_Click()
    If (MsgBox("Delete the data?", vbExclamation + vbYesNo, "Caution") = vbYes) Then
        
        Dim targetrow As Long: targetrow = searchRowIndex(generalid)
        
        '_del付与(削除レコード化する)
        Call delToExcel(targetrow)
        
        Dim selectedDate As Date: selectedDate = CDate(parseYymmdd(Me.TextBoxDate.value))
        
        Me.Hide
        FrontDataDisp.FrontDataUpdate selectedDate
        goBack
    End If
End Sub

'新規追加/変更保存ボタン
Private Sub CommandButtonSave_Click()
    '入力値を検証（必要であれば）
    If Not validateInputs() Then Exit Sub
    
    '入力内容を1次元配列化
    Dim newRow As Variant
    newRow = collectInputData()
    
    'Excelに1行追記
    Dim targetrow As Long: targetrow = searchRowIndex(generalid)
    Call writeToExcel(newRow, targetrow)
    
    Dim selectedDate As Date: selectedDate = CDate(parseYymmdd(Me.TextBoxDate.value))
    
    Call FrontDataDisp.FrontDataUpdate(selectedDate) 'Excelから読込直して受付画面のListView更新
    Call CustomerDetailDisp.setupDetail(newRow(COL_CUSTM), newRow(COL_TEL)) '対象レコード渡して受付画面のListView更新
    Call HistoryDataDisp.api_searchHistory(HistoryDataDisp.TextBoxFree.Text)
    
    '入力画面を閉じ、前の画面に戻る
    Me.Hide
    goBack
End Sub
Private Function searchRowIndex(ByVal editID As String) As Long
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("入力シート")
    
    If Not isEditMode Then
        searchRowIndex = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        Exit Function
    End If
    
    Dim arr As Variant: arr = ws.Columns(1).value
    Dim i As Long
    For i = UBound(arr, 1) To LBound(arr, 1) Step -1
        If arr(i, 1) = editID Then
            searchRowIndex = i
            Exit Function
        End If
    Next i
End Function

'検索ボタン
Private Sub CommandButtonSearchName_Click()
    If TextBoxName.Text = "" Then Exit Sub
    Call searchNameOrTel("name")
End Sub
Private Sub CommandButtonSearchTel_Click()
    If TextBoxTel.Text = "" Then Exit Sub
    Call searchNameOrTel("tel")
End Sub
Private Sub searchNameOrTel(ByVal target As String)
    Dim searchStr As String
    Dim searchCol As Long
    Dim targetBox As Object
    Dim targetDic As Long
    Select Case LCase(target)
        Case "name"
            searchStr = TextBoxName.Text
            searchCol = COL_CUSTM
            Set targetBox = TextBoxTel
            targetDic = 3 'DictionaryのValue配列のインデックス
        Case "tel"
            searchStr = TextBoxTel.Text
            searchCol = COL_TEL
            Set targetBox = TextBoxName
            targetDic = 2 'DictionaryのValue配列のインデックス
    End Select
    
    Dim resultArr As Variant: resultArr = searchAfromB(searchStr, searchCol)
    If IsEmpty(resultArr) Then
        MsgBox "該当なし", vbOKOnly, "検索結果"
        Exit Sub
    End If
    
    'NameとTelでDic化
    ' key = generalId | value = name, tel
    Dim customDic As Object: Set customDic = summaryByCustom(resultArr)
    
    'hit1件なら入力欄に即反映(非遷移アクション)
    Dim keys As Variant: keys = customDic.keys
    If customDic.Count = 1 Then
        TextBoxName.Text = customDic(keys(LBound(keys)))(2)
        TextBoxTel.Text = customDic(keys(LBound(keys)))(3)
        
        'プレースホルダ切替
        Call switchPlaceholder(Me)
    End If
    'hit2件以上なら候補選択画面ListViewを表示
    If customDic.Count >= 2 Then
        Call CustomerSearchDisp.setupScreen(customDic)
        CustomerSearchDisp.Show vbModal
        
        'プレースホルダ切替
        Call switchPlaceholder(Me)
    End If
End Sub

'入力/バリデーションチェック
Private Function validateInputs() As Boolean
    Dim result As Boolean: result = True
    
    '全ての入力欄をチェックする
    '日付
    Dim errDate As String: errDate = apiValidate(TextBoxDate.Text, Array("required", "yymmdd"))
    If errDate = "" Then
        LabelErrorDate.Visible = False
    Else
        LabelErrorDate.Caption = errDate
        LabelErrorDate.Visible = True
        result = False
    End If
    
    'AD_ここはコンボボックスなので専用処理
    If ComboBoxAd.ListIndex = 0 Then
        LabelErrorAd.Visible = True
        result = False
    Else
        LabelErrorAd.Visible = False
    End If
    'Type_ここはコンボボックスなので専用処理
    If ComboBoxType.ListIndex = 0 Then
        LabelErrorType.Visible = True
        result = False
    Else
        LabelErrorAd.Visible = False
    End If
    
    'Name
    Dim errName As String: errName = apiValidate(TextBoxName.Text, Array("required"))
    If errName = "" Then
        LabelErrorName.Visible = False
    Else
        LabelErrorName.Caption = errName
        LabelErrorName.Visible = True
        result = False
    End If
    
    'TEL
    Dim errTel As String: errTel = apiValidate(TextBoxTel.Text, Array("required"))
    If errTel = "" Then
        LabelErrorTel.Visible = False
    Else
        LabelErrorTel.Caption = errTel
        LabelErrorTel.Visible = True
        result = False
    End If
    
    'Sales
    Dim errSales As String: errSales = apiValidate(TextBoxSales.Text, Array("required", "numeric"))
    If errSales = "" Then
        LabelErrorSales.Visible = False
    Else
        LabelErrorSales.Caption = errSales
        LabelErrorSales.Visible = True
        result = False
    End If
    
    'CaCost
    Dim errCaCost As String: errCaCost = apiValidate(TextBoxCost.Text, Array("required", "numeric"))
    If errCaCost = "" Then
        LabelErrorCost.Visible = False
    Else
        LabelErrorCost.Caption = errCaCost
        LabelErrorCost.Visible = True
        result = False
    End If
    
    '一つでもresult=FalseがあればFalseでreturn
    validateInputs = result
End Function

'///////////////////////////////////////////////////////////
'Excelに追記
'///////////////////////////////////////////////////////////
Private Sub writeToExcel(ByVal Data As Variant, ByVal targetrow As Long)
    InputSh.Range(InputSh.Cells(targetrow, 1), InputSh.Cells(targetrow, UBound(Data))).value = Data
    
    'Telが0始まりで崩れる。配列データはCstrしているが、念のためセルの表示形式でも文字列に設定。
    InputSh.Cells(targetrow, COL_TEL).NumberFormat = "@"
    
    ThisWorkbook.Save 'ブック保存しておく
End Sub
Private Sub delToExcel(ByVal targetrow As Long)
    InputSh.Cells(targetrow, COL_A).value = InputSh.Cells(targetrow, COL_A).value & "_del"
    InputSh.Cells(targetrow, COL_DATE).value = InputSh.Cells(targetrow, COL_DATE).value & "_del"
    ThisWorkbook.Save 'ブック保存しておく
End Sub

'///////////////////////////////////////////////////////////
'操作イベント発火（非遷移系操作）
'///////////////////////////////////////////////////////////
'プレースホルダ再表示(入力欄抜け)
Private Sub TextBoxDate_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call switchPlaceholder(Me): End Sub
Private Sub TextBoxTime_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call switchPlaceholder(Me): End Sub
Private Sub TextBoxName_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call switchPlaceholder(Me): End Sub
Private Sub TextBoxTel_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call switchPlaceholder(Me): End Sub
Private Sub TextBoxNG_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call switchPlaceholder(Me): End Sub
Private Sub TextBoxNotes_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call switchPlaceholder(Me): End Sub
Private Sub TextBoxCast_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call switchPlaceholder(Me): End Sub
Private Sub TextBoxCourse_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call switchPlaceholder(Me): End Sub
Private Sub TextBoxService_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call switchPlaceholder(Me): End Sub
Private Sub TextBoxOP_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call switchPlaceholder(Me): End Sub
Private Sub TextBoxDestination_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call switchPlaceholder(Me): End Sub
Private Sub TextBoxExpand_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call switchPlaceholder(Me): End Sub
Private Sub TextBoxSales_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call switchPlaceholder(Me): End Sub
Private Sub TextBoxCost_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call switchPlaceholder(Me): End Sub
Private Sub TextBoxProfit_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call switchPlaceholder(Me): End Sub
Private Sub TextBoxQB_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call switchPlaceholder(Me): End Sub
Private Sub TextBoxSB_Exit(ByVal Cancel As MSForms.ReturnBoolean): Call switchPlaceholder(Me): End Sub
'Profit計算
Private Sub TextBoxSales_Change(): Call calculateProfit: End Sub
Private Sub TextBoxCost_Change(): Call calculateProfit: End Sub
Private Sub calculateProfit()
    Dim sales As Long
    If IsNumeric(TextBoxSales.Text) Then
        sales = TextBoxSales.Text
    Else
        sales = 0
    End If
    Dim cost As Long
    If IsNumeric(TextBoxCost.Text) Then
        cost = TextBoxCost.Text
    Else
        cost = 0
    End If
    
    Dim profit As Long: profit = sales - cost
    TextBoxProfit.Text = profit
End Sub

'///////////////////////////////////////////////////////////
'特定のキー操作
' - 特定項目でのEnterKeyの挙動制御
' * ここで制御していない項目は、Tab送りになる
'    ->確定処理にするか迷ったが、誤操作時に害ありそうなので、無害なTab送りのままとした
'///////////////////////////////////////////////////////////
Private Sub TextBoxName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                                ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call searchNameOrTel("name")
    End If
End Sub
Private Sub TextBoxTel_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                                ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call searchNameOrTel("tel")
    End If
End Sub

'///////////////////////////////////////////////////////////
'遷移系以外のボタン操作
'///////////////////////////////////////////////////////////
'Date欄をクリックしたとき_プレースホルダ
Private Sub PlaceholderDate_Click()
    PlaceholderDate.Visible = False
    TextBoxDate.SetFocus
End Sub
'Date欄をクリック_TextBox
Private Sub TextBoxDate_Enter()
    PlaceholderDate.Visible = False
End Sub

'Time欄をクリックしたとき_プレースホルダ
Private Sub PlaceholderTime_Click()
    PlaceholderTime.Visible = False
    TextBoxTime.Text = PlaceholderTime.Caption
    TextBoxTime.SetFocus
End Sub
'Time欄をクリック_TextBox
Private Sub TextBoxTime_Enter()
    PlaceholderTime.Visible = False
    If TextBoxTime.Text = "" Then TextBoxTime.Text = PlaceholderTime.Caption
End Sub

'Name欄をクリックしたとき_プレースホルダ
Private Sub PlaceholderName_Click()
    PlaceholderName.Visible = False
    TextBoxName.SetFocus
End Sub
'Name欄をクリック_TextBox
Private Sub TextBoxName_Enter()
    PlaceholderName.Visible = False
End Sub

'Tel欄をクリックしたとき_プレースホルダ
Private Sub PlaceholderTel_Click()
    PlaceholderTel.Visible = False
    TextBoxTel.SetFocus
End Sub
'Tel欄をクリック_TextBox
Private Sub TextBoxTel_Enter()
    PlaceholderTel.Visible = False
End Sub

'Cast欄をクリックしたとき_プレースホルダ
Private Sub PlaceholderCast_Click()
    PlaceholderCast.Visible = False
    TextBoxCast.SetFocus
End Sub
'Cast欄をクリック_TextBox
Private Sub TextBoxCast_Enter()
    PlaceholderCast.Visible = False
End Sub

'Course欄をクリックしたとき_プレースホルダ
Private Sub PlaceholderCourse_Click()
    PlaceholderCourse.Visible = False
    TextBoxCourse.SetFocus
End Sub
'Course欄をクリック_TextBox
Private Sub TextBoxCourse_Enter()
    PlaceholderCourse.Visible = False
End Sub

'Service欄をクリックしたとき_プレースホルダ
Private Sub PlaceholderService_Click()
    PlaceholderService.Visible = False
    TextBoxService.SetFocus
End Sub
'Service欄をクリック_TextBox
Private Sub TextBoxService_Enter()
    PlaceholderService.Visible = False
End Sub

'OP欄をクリックしたとき_プレースホルダ
Private Sub PlaceholderOP_Click()
    PlaceholderOP.Visible = False
    TextBoxOP.SetFocus
End Sub
'OP欄をクリック_TextBox
Private Sub TextBoxOP_Enter()
    PlaceholderOP.Visible = False
End Sub

'Destination欄をクリックしたとき_プレースホルダ
Private Sub PlaceholderDestination_Click()
    PlaceholderDestination.Visible = False
    TextBoxDestination.SetFocus
End Sub
'Destination欄をクリック_TextBox
Private Sub TextBoxDestination_Enter()
    PlaceholderDestination.Visible = False
End Sub

'Expand欄をクリックしたとき_プレースホルダ
Private Sub PlaceholderExpand_Click()
    PlaceholderExpand.Visible = False
    TextBoxExpand.SetFocus
End Sub
'Expand欄をクリック_TextBox
Private Sub TextBoxExpand_Enter()
    PlaceholderExpand.Visible = False
End Sub

'Sales欄をクリックしたとき_プレースホルダ
Private Sub PlaceholderSales_Click()
    PlaceholderSales.Visible = False
    TextBoxSales.SetFocus
End Sub
'Sales欄をクリック_TextBox
Private Sub TextBoxSales_Enter()
    PlaceholderSales.Visible = False
End Sub

'Cost欄をクリックしたとき_プレースホルダ
Private Sub PlaceholderCost_Click()
    PlaceholderCost.Visible = False
    TextBoxCost.SetFocus
End Sub
'Cost欄をクリック_TextBox
Private Sub TextBoxCost_Enter()
    PlaceholderCost.Visible = False
End Sub

'QB欄をクリックしたとき_プレースホルダ
Private Sub PlaceholderQB_Click()
    PlaceholderQB.Visible = False
    TextBoxQB.SetFocus
End Sub
'QB欄をクリック_TextBox
Private Sub TextBoxQB_Enter()
    PlaceholderQB.Visible = False
End Sub

'SB欄をクリックしたとき_プレースホルダ
Private Sub PlaceholderSB_Click()
    PlaceholderSB.Visible = False
    TextBoxSB.SetFocus
End Sub
'SB欄をクリック_TextBox
Private Sub TextBoxSB_Enter()
    PlaceholderSB.Visible = False
End Sub

'Labelのマウスオーバー関連(WindowsAPI使用)
Private Sub LabelBack_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                             ByVal X As Single, ByVal Y As Single)
    Dim hCursor As LongPtr
    hCursor = LoadCursor(0, IDC_HAND)
    SetCursor hCursor
    LabelBack.BorderStyle = fmBorderStyleSingle
    LabelBack.BorderColor = &H8000000D
End Sub
Private Sub ImageDelete_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                             ByVal X As Single, ByVal Y As Single)
    Dim hCursor As LongPtr
    hCursor = LoadCursor(0, IDC_HAND)
    SetCursor hCursor
    ImageDelete.BorderStyle = fmBorderStyleSingle
    ImageDelete.BorderColor = &HFF&
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    SetCursor LoadCursor(0, 32512) ' IDC_ARROW（通常の矢印カーソル）
    LabelBack.BorderStyle = fmBorderStyleNone
    ImageDelete.BorderStyle = fmBorderStyleNone
End Sub

'//////////////////////////////////////////
'画面呼び出し時のプレースホルダー切替処理
'//////////////////////////////////////////
Private Sub switchPlaceholder(ByRef InputFormDisp As Object)
    With InputFormDisp
        If .TextBoxDate.Text = "" Then
            .PlaceholderDate.Visible = True
        Else
            .PlaceholderDate.Visible = False
        End If
        
        If .TextBoxTime.Text = "" Then
            .PlaceholderTime.Visible = True
        Else
            .PlaceholderTime.Visible = False
        End If
        
        If .TextBoxName.Text = "" Then
            .PlaceholderName.Visible = True
        Else
            .PlaceholderName.Visible = False
        End If
        
        If .TextBoxTel.Text = "" Then
            .PlaceholderTel.Visible = True
        Else
            .PlaceholderTel.Visible = False
        End If
        
        If .TextBoxCast.Text = "" Then
            .PlaceholderCast.Visible = True
        Else
            .PlaceholderCast.Visible = False
        End If
        
        If .TextBoxCourse.Text = "" Then
            .PlaceholderCourse.Visible = True
        Else
            .PlaceholderCourse.Visible = False
        End If
        
        If .TextBoxService.Text = "" Then
            .PlaceholderService.Visible = True
        Else
            .PlaceholderService.Visible = False
        End If
        
        If .TextBoxOP.Text = "" Then
            .PlaceholderOP.Visible = True
        Else
            .PlaceholderOP.Visible = False
        End If
        
        If .TextBoxDestination.Text = "" Then
            .PlaceholderDestination.Visible = True
        Else
            .PlaceholderDestination.Visible = False
        End If
        
        If .TextBoxExpand.Text = "" Then
            .PlaceholderExpand.Visible = True
        Else
            .PlaceholderExpand.Visible = False
        End If
        
        If .TextBoxSales.Text = "" Then
            .PlaceholderSales.Visible = True
        Else
            .PlaceholderSales.Visible = False
        End If
        
        If .TextBoxCost.Text = "" Then
            .PlaceholderCost.Visible = True
        Else
            .PlaceholderCost.Visible = False
        End If
        
        If .TextBoxQB.Text = "" Then
            .PlaceholderQB.Visible = True
        Else
            .PlaceholderQB.Visible = False
        End If
        
        If .TextBoxSB.Text = "" Then
            .PlaceholderSB.Visible = True
        Else
            .PlaceholderSB.Visible = False
        End If
    End With
End Sub

'///////////////////////////////////////////////////////////
'入力欄クリアリング
'///////////////////////////////////////////////////////////
Private Sub clearAllValue()
    TextBoxDate.Text = ""
    ComboBoxAd.ListIndex = 0
    ComboBoxType.ListIndex = 0
    TextBoxTime.Text = ""
    TextBoxName.Text = ""
    TextBoxTel.Text = ""
    TextBoxNG.Text = ""
    TextBoxNotes.Text = ""
    TextBoxCast.Text = ""
    TextBoxCourse.Text = ""
    TextBoxService.Text = ""
    TextBoxOP.Text = ""
    TextBoxDestination.Text = ""
    TextBoxExpand.Text = ""
    TextBoxSales.Text = ""
    TextBoxCost.Text = ""
    TextBoxProfit.Text = ""
    TextBoxQB.Text = ""
    TextBoxSB.Text = ""
End Sub
