VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AdvancedSearchDisp 
   Caption         =   "UserForm1"
   ClientHeight    =   2220
   ClientLeft      =   -264
   ClientTop       =   -984
   ClientWidth     =   3408
   OleObjectBlob   =   "AdvancedSearchDisp.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "AdvancedSearchDisp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub UserForm_Initialize()
    Call UiConfig_AdvancedSearchDisp.configUiDesign(Me)

End Sub

'///////////////////////////////////////////////////////////
'遷移系ボタン操作
'///////////////////////////////////////////////////////////
Private Sub LabelBack_Click()
    Me.Hide
End Sub

