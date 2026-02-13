VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrontDataDisp 
   Caption         =   "受付データ表示"
   ClientHeight    =   3612
   ClientLeft      =   -3276
   ClientTop       =   -13152
   ClientWidth     =   14352
   OleObjectBlob   =   "FrontDataDisp.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FrontDataDisp"
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

'ListViewのクラス宣言
Private drawer As ListViewDrawer

'日付ラベル用の現在表示日付_モジュールレベル
Private currentDate As Date

'日付更新
Private Sub updateDateLabel()
    Me.LabelDate.Caption = Format(currentDate, "yyyy/mm/dd")
End Sub

'///////////////////////////////////////////////////////////
'初期化処理
'///////////////////////////////////////////////////////////
Private Sub UserForm_Initialize()
    '日付初期化(最初は本日表示)
    currentDate = Date
    Call updateDateLabel
    
    'UIの見た目（サイズや位置などの見た目要素）を設定
    Call UiConfig_FrontDataDisp.configUiDesign(Me)
    
    'ListView用クラスのインスタンス生成
    Set drawer = New ListViewDrawer
    Dim widths As Variant: widths = UiConfig_FrontDataDisp.configFrontDataView(widths)
    Dim listHeader As Variant: listHeader = DispMod.getArrHeader()
    Dim listBody As Variant: listBody = DispMod.getArrBody(currentDate)
    Call drawer.init(Me.ListView1, listHeader, widths)
    Call drawer.Draw(listBody)
    
    '日付処理
    Call updateDateLabel
End Sub

'///////////////////////////////////////////////////////////
'受付データ画面のListView更新
'入力画面での登録操作から遷移時に必ず必要
'///////////////////////////////////////////////////////////
Public Sub FrontDataUpdate(ByVal targetDate As Date)
    '日付更新
    currentDate = targetDate
    
    'ListViewクラス更新
    Dim widths As Variant: widths = UiConfig_FrontDataDisp.configFrontDataView(widths)
    Dim arrHeader As Variant: arrHeader = DispMod.getArrHeader()
    Dim arrData As Variant: arrData = DispMod.getArrBody(currentDate)
    Call drawer.init(Me.ListView1, arrHeader, widths)
    Call drawer.Draw(arrData)
    
    '日付処理
    Call updateDateLabel
End Sub

'///////////////////////////////////////////////////////////
'日付操作系遷移
'///////////////////////////////////////////////////////////
Private Sub LabelPrevDate_Click()
    LabelPrevDate.BorderStyle = fmBorderStyleSingle
    LabelPrevDate.BorderColor = &H8000000F
    currentDate = DateAdd("d", -1, currentDate)
    
    Call updateDateLabel
    Call FrontDataUpdate(currentDate)
End Sub
Private Sub LabelNextDate_Click()
    LabelNextDate.BorderStyle = fmBorderStyleSingle
    LabelNextDate.BorderColor = &H8000000F
    currentDate = DateAdd("d", 1, currentDate)
    
    Call updateDateLabel
    Call FrontDataUpdate(currentDate)
End Sub
'日付指定ジャンプ
Private Sub LabelDate_Click()
    LabelDate.BorderStyle = fmBorderStyleSingle
    LabelDate.BorderColor = &H8000000F
    
    Dim strInput As String
    Do: strInput = InputBox(prompt:="Please enter the date you want to view.", _
                            Title:="Jump to date", _
                            Default:=Format(Date, "yymmdd"))
        
        '閉じる、キャンセル時Exit
        If strInput = "" Then Exit Sub
        
        'バリデーションチェック
        Dim errMsg As String: errMsg = apiValidate(strInput, Array("yymmdd"))
        If errMsg <> "" Then
            MsgBox errMsg, vbCritical, "Validation Error"
            strInput = ""
        End If
    Loop While errMsg <> ""
    
    '指定日付を確定
    currentDate = parseYymmdd(strInput)
    
    Call updateDateLabel
    Call FrontDataUpdate(currentDate)
End Sub

'///////////////////////////////////////////////////////////
'遷移系ボタン操作
'///////////////////////////////////////////////////////////
'戻る
Private Sub LabelBack_Click()
    Me.Hide
    navigateTo HomeDisp
End Sub
'新規入力ボタン
Private Sub CommandButtonInput_Click()
    Me.Hide
    InputFormDisp.reloadInputs currentDate, ListView1.ListItems.Count
    navigateTo InputFormDisp
End Sub
'タブ切替
Private Sub LabelTab1_Click()
'    Me.Hide
'    navigateTo FrontDataDisp
End Sub
Private Sub LabelTab2_Click()
    Me.Hide
    navigateTo HistoryDataDisp
End Sub
Private Sub LabelTab3_Click()
    Me.Hide
    navigateTo CheckDisp
End Sub
Private Sub LabelTab4_Click()
    Me.Hide
    navigateTo InspectionDisp
End Sub
Private Sub LabelTab5_Click()
    Me.Hide
    navigateTo ClosingDisp
End Sub

'リストビュー内レコードダブルクリック
Private Sub ListView1_dblClick()
    If Me.ListView1.SelectedItem Is Nothing Then Exit Sub
    
    Dim rowList As Variant
    rowList = drawer.getSelectedListViewRow(Me.ListView1)
    Me.Hide
    Call InputFormDisp.editInputs(rowList, _
                                  ListView1.SelectedItem.Text, _
                                  drawer.getGeneralId(rowList(1)) _
    )
    
    navigateTo InputFormDisp
End Sub

'閉じるボタン_フォームオブジェクトクリア
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Call DispMod.clearingForms
    End If
End Sub

'///////////////////////////////////////////////////////////
'見た目系の反応（非遷移系操作）
'///////////////////////////////////////////////////////////
'Labelのマウスオーバー関連(WindowsAPI使用)
Private Sub LabelBack_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                             ByVal X As Single, ByVal Y As Single)
    Dim hCursor As LongPtr
    hCursor = LoadCursor(0, IDC_HAND)
    SetCursor hCursor
    LabelBack.BorderStyle = fmBorderStyleSingle
    LabelBack.BorderColor = &H8000000D
End Sub
Private Sub LabelTab1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                             ByVal X As Single, ByVal Y As Single)
    Dim hCursor As LongPtr
    hCursor = LoadCursor(0, IDC_HAND)
    SetCursor hCursor
End Sub
Private Sub LabelTab2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                             ByVal X As Single, ByVal Y As Single)
    Dim hCursor As LongPtr
    hCursor = LoadCursor(0, IDC_HAND)
    SetCursor hCursor
End Sub
Private Sub LabelTab3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                             ByVal X As Single, ByVal Y As Single)
    Dim hCursor As LongPtr
    hCursor = LoadCursor(0, IDC_HAND)
    SetCursor hCursor
End Sub
Private Sub LabelTab4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                             ByVal X As Single, ByVal Y As Single)
    Dim hCursor As LongPtr
    hCursor = LoadCursor(0, IDC_HAND)
    SetCursor hCursor
End Sub
Private Sub LabelTab5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                             ByVal X As Single, ByVal Y As Single)
    Dim hCursor As LongPtr
    hCursor = LoadCursor(0, IDC_HAND)
    SetCursor hCursor
End Sub
Private Sub LabelDate_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                             ByVal X As Single, ByVal Y As Single)
    Dim hCursor As LongPtr
    hCursor = LoadCursor(0, IDC_HAND)
    SetCursor hCursor
    LabelDate.BorderStyle = fmBorderStyleSingle
    LabelDate.BorderColor = &H8000000D
End Sub
Private Sub LabelPrevDate_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                             ByVal X As Single, ByVal Y As Single)
    Dim hCursor As LongPtr
    hCursor = LoadCursor(0, IDC_HAND)
    SetCursor hCursor
    LabelPrevDate.BorderStyle = fmBorderStyleSingle
    LabelPrevDate.BorderColor = &H8000000D
End Sub
Private Sub LabelNextDate_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                             ByVal X As Single, ByVal Y As Single)
    Dim hCursor As LongPtr
    hCursor = LoadCursor(0, IDC_HAND)
    SetCursor hCursor
    LabelNextDate.BorderStyle = fmBorderStyleSingle
    LabelNextDate.BorderColor = &H8000000D
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    SetCursor LoadCursor(0, 32512) ' IDC_ARROW（通常の矢印カーソル）
    LabelBack.BorderStyle = fmBorderStyleNone
    LabelDate.BorderStyle = fmBorderStyleNone
    LabelPrevDate.BorderStyle = fmBorderStyleNone
    LabelNextDate.BorderStyle = fmBorderStyleNone
End Sub
