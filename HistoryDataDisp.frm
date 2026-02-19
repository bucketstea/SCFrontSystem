VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HistoryDataDisp 
   Caption         =   "UserForm1"
   ClientHeight    =   2760
   ClientLeft      =   -276
   ClientTop       =   -1368
   ClientWidth     =   12408
   OleObjectBlob   =   "HistoryDataDisp.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "HistoryDataDisp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

'ListViewのクラス宣言
Private drawer As ListViewDrawer

'/////////////////////////////////////////////////////////// Windows API の宣言
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" _
    (ByVal hInstance As LongPtr, ByVal lpCursorName As Long) As LongPtr
Private Declare PtrSafe Function SetCursor Lib "user32" _
    (ByVal hCursor As LongPtr) As LongPtr
Private Const IDC_HAND = 32649
'///////////////////////////////////////////////////////////

Public Sub api_searchHistory(Optional ByVal str As String = "")
    Me.TextBoxFree.Text = str
    Call CommandButtonSearch_Click
End Sub

Private Sub UserForm_Initialize()
    Call UiConfig_HistoryDataDisp.configUiDesign(Me)
End Sub

'///////////////////////////////////////////////////////////
'遷移系ボタン操作
'///////////////////////////////////////////////////////////
'戻る
Private Sub LabelBack_Click()
    Me.Hide
    navigateTo HomeDisp
End Sub
'タブ切替
Private Sub LabelTab1_Click()
    Me.Hide
    navigateTo FrontDataDisp
End Sub
Private Sub LabelTab2_Click()
'    Me.Hide
'    navigateTo HistoryDataDisp
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

'///////////////////////////////////////////////////////////
'検索機能
'///////////////////////////////////////////////////////////
Private Sub CommandButtonSearch_Click()
    Dim targetstr As String: targetstr = TextBoxFree.Text
    If targetstr = "" Then Exit Sub
    Dim targetArr As Variant
    targetArr = quickSort2dByColumn(searchFromAllByFreeword(targetstr), COL_DATE, "DESC")
    
    'ListView用クラスのインスタンス生成
    Set drawer = New ListViewDrawer
    Dim widths As Variant: widths = UiConfig_HistoryDataDisp.configHistoryDataView(widths)
    Dim listHeader As Variant: listHeader = DispMod.getArrHeader()
    Dim listBody As Variant: listBody = targetArr
    Call drawer.init(Me.ListView1, listHeader, widths)
    Call drawer.Draw(listBody)
    
    'Hitカウンタ更新
    If IsEmpty(targetArr) Then
        LabelHits.Caption = 0 & " Hit　"
        MsgBox "該当なし", vbOKOnly, "検索結果"
    ElseIf UBound(targetArr, 1) = 1 Then
        LabelHits.Caption = 1 & " Hit　"
    Else
        LabelHits.Caption = UBound(listBody, 1) & " Hits!"
    End If
End Sub
Private Function searchFromAllByFreeword(ByVal str As String) As Variant
    Dim searchStr As String: searchStr = str
    Dim searchCols As Variant: searchCols = Array(COL_DATE, _
                                                  COL_STAFF, _
                                                  COL_CUSTM, _
                                                  COL_TEL, _
                                                  COL_NG, _
                                                  COL_NOTE)
    
    Dim appendedArr As Variant
    Dim i As Long
    For i = LBound(searchCols) To UBound(searchCols)
        Dim hitsArr As Variant: hitsArr = DispMod.searchAfromB(searchStr, searchCols(i))
        If Not IsEmpty(hitsArr) Then appendedArr = utils.appendArray(appendedArr, hitsArr)
    Next i
    If IsEmpty(appendedArr) Then Exit Function
    
    'このままだと重複あるので、ID(A列)でユニーク化
    Dim uniqueArr As Variant: uniqueArr = utils.extractUnique(appendedArr, COL_A)
    '日付列が文字列だとやっかい、一度数値化する
    Dim dateNormalizedArr As Variant: dateNormalizedArr = DispMod.normalizeDateCol(uniqueArr)
    'ソート処理
    Dim sortedArr As Variant: sortedArr = quickSort2dByColumn(dateNormalizedArr, COL_DATE)
    
    'return
    searchFromAllByFreeword = sortedArr
End Function

'高度な検索機能
Private Sub ImageAdvance_Click()
    AdvancedSearchDisp.Show vbModal
End Sub

'///////////////////////////////////////////////////////////
'キー操作
'///////////////////////////////////////////////////////////
Private Sub TextBoxFree_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
                                ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        CommandButtonSearch_Click
    End If
End Sub

'リストビュー内レコードダブルクリック
Private Sub ListView1_dblClick()
    If Me.ListView1.SelectedItem Is Nothing Then Exit Sub
    
    Dim rowList As Variant
    rowList = drawer.getSelectedListViewRow(Me.ListView1)
    Me.Hide
    Call CustomerDetailDisp.setupDetail(rowList(COL_CUSTM), rowList(COL_TEL))
    
    navigateTo CustomerDetailDisp
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
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    SetCursor LoadCursor(0, 32512) ' IDC_ARROW（通常の矢印カーソル）
    LabelBack.BorderStyle = fmBorderStyleNone
End Sub
Private Sub ImageAdvance_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                                   ByVal X As Single, ByVal Y As Single)
    Dim hCursor As LongPtr
    hCursor = LoadCursor(0, IDC_HAND)
    SetCursor hCursor
    ImageAdvance.BorderStyle = fmBorderStyleSingle
End Sub

