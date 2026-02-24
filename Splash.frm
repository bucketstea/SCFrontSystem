VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Splash 
   Caption         =   "S-Kore front system"
   ClientHeight    =   732
   ClientLeft      =   48
   ClientTop       =   108
   ClientWidth     =   2136
   OleObjectBlob   =   "Splash.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'スプラッシュ表示
'なんかやってそうな見た目
'何か予備処理など必要なことあればこれをcallする前後に
Private Sub UserForm_Activate()
    Dim i As Long
    Dim dots As String
    
    For i = 0 To 3
        dots = String(i Mod 4, ".")
        Me.lblstatus.Caption = "Starting" & dots
        DoEvents
        Application.Wait Now + TimeValue("0:00:01")
    Next i
End Sub
Private Sub UserForm_Initialize()
    With Me
        .Caption = "S-Core Front System"
        .Height = 100
        .Width = 180
        With lblstatus
            .Top = 20
            .Left = 40
            .Font.Size = 18
        End With
    End With
End Sub
