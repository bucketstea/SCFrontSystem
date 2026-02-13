VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ClosingDisp 
   Caption         =   "UserForm1"
   ClientHeight    =   3084
   ClientLeft      =   96
   ClientTop       =   372
   ClientWidth     =   17640
   OleObjectBlob   =   "ClosingDisp.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ClosingDisp"
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
'//////////////////////////////////////////////////////////

Private Sub UserForm_Initialize()
    Call UiConfig_ClosingDisp.configUiDesign(Me)
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
'    Me.Hide
'    navigateTo ClosingDisp
End Sub

Private Sub CommandButton1_Click()
    
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

