VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HomeDisp 
   Caption         =   "UserForm1"
   ClientHeight    =   2280
   ClientLeft      =   -12
   ClientTop       =   36
   ClientWidth     =   1620
   OleObjectBlob   =   "HomeDisp.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "HomeDisp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub UserForm_Initialize()
    With HomeDisp
        .StartUpPosition = 0
        .Left = Application.Left + 10
        .Top = Application.Top + 10
        .Height = 520
        .Width = 300
        
        .Caption = "ホーム画面"
        .CommandButtonFront.Caption = "Front"
        .CommandButtonHistory.Caption = "History"
        .CommandButtonChecking.Caption = "Checking"
        .CommandButtonInspection.Caption = "Inspection"
        .CommandButtonClosing.Caption = "Closing"
        
        With .CommandButtonFront
            .Top = 40
            .Left = 50
            .Height = 40
            .Width = 200
            .TabStop = False
            .Font.name = "Yu Gothic UI"
            .Font.Size = 20
            .Font.Bold = True
            .BackColor = RGB(255, 255, 255)
            .ForeColor = RGB(0, 0, 0)
        End With
        With .CommandButtonHistory
            .Top = 120
            .Left = 50
            .Height = 40
            .Width = 200
            .TabStop = False
            .Font.name = "Yu Gothic UI"
            .Font.Size = 20
            .Font.Bold = True
            .BackColor = RGB(255, 255, 255)
            .ForeColor = RGB(0, 0, 0)
        End With
        With .CommandButtonChecking
            .Top = 200
            .Left = 50
            .Height = 40
            .Width = 200
            .TabStop = False
            .Font.name = "Yu Gothic UI"
            .Font.Size = 20
            .Font.Bold = True
            .BackColor = RGB(255, 255, 255)
            .ForeColor = RGB(0, 0, 0)
        End With
        With .CommandButtonInspection
            .Top = 280
            .Left = 50
            .Height = 40
            .Width = 200
            .TabStop = False
            .Font.name = "Yu Gothic UI"
            .Font.Size = 20
            .Font.Bold = True
            .BackColor = RGB(255, 255, 255)
            .ForeColor = RGB(0, 0, 0)
        End With
        With .CommandButtonClosing
            .Top = 360
            .Left = 50
            .Height = 40
            .Width = 200
            .TabStop = False
            .Font.name = "Yu Gothic UI"
            .Font.Size = 20
            .Font.Bold = True
            .BackColor = RGB(255, 255, 255)
            .ForeColor = RGB(0, 0, 0)
        End With
    End With
End Sub

Private Sub CommandButtonFront_Click()
    Me.Hide
    navigateTo FrontDataDisp
End Sub
Private Sub CommandButtonHistory_Click()
    Me.Hide
    navigateTo HistoryDataDisp
End Sub
Private Sub CommandButtonChecking_Click()
    Me.Hide
    navigateTo CheckDisp
End Sub
Private Sub CommandButtonInspection_Click()
    Me.Hide
    navigateTo InspectionDisp
End Sub
Private Sub CommandButtonClosing_Click()
    Me.Hide
    navigateTo ClosingDisp
End Sub

'閉じるボタン_フォームオブジェクトクリア
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Call DispMod.clearingForms
    End If
End Sub
