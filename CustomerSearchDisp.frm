VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerSearchDisp 
   Caption         =   "選択"
   ClientHeight    =   2928
   ClientLeft      =   -564
   ClientTop       =   -2220
   ClientWidth     =   3372
   OleObjectBlob   =   "CustomerSearchDisp.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CustomerSearchDisp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

'ListViewのクラス宣言
Private drawer As New ListViewDrawer

'///////////////////////////////////////////////////////////
'初期化処理
'///////////////////////////////////////////////////////////
Private Sub UserForm_Initialize()
    'UIの見た目（サイズや位置などの見た目要素）を設定
    Call UiConfig_CustomerSearchDisp.configUiDesign(Me)
End Sub
'Initialize前にデータを受けてセッティングする
Public Sub setupScreen(ByRef customersDic As Object)
    'ListView用クラスのインスタンス生成
    Dim widths As Variant: widths = Array(20, 156, 160)
    Dim listHeader As Variant: ReDim listHeader(1 To 1, 1 To 3)
    listHeader(1, 1) = "Id": listHeader(1, 2) = "Name": listHeader(1, 3) = "Tel"
    Dim listBody As Variant: listBody = convertDicToArr(customersDic)
    Call drawer.init(Me.ListView1, listHeader, widths)
    Call drawer.Draw(listBody)
End Sub
'Dicを配列化
Private Function convertDicToArr(ByVal dic As Object) As Variant
    Dim resultArr As Variant: ReDim resultArr(1 To dic.Count, 1 To 3)
    Dim arrCt As Long: arrCt = 1
    Dim key As Variant
    For Each key In dic.keys
        Dim j As Long
        For j = 1 To 3
            resultArr(arrCt, j) = dic(key)(j)
        Next j
        arrCt = arrCt + 1
    Next key
    convertDicToArr = resultArr
End Function

Private Sub ListView1_dblClick()
    If Me.ListView1.SelectedItem Is Nothing Then Exit Sub
    
    Dim rowList As Variant
    rowList = drawer.getSelectedListViewRow(Me.ListView1)
    Me.Hide
    
    '入力欄に反映
    InputFormDisp.TextBoxName.Text = rowList(2)
    InputFormDisp.TextBoxTel.Text = rowList(3)
End Sub
