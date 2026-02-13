Attribute VB_Name = "UiConfig_InspectionDisp"
Option Explicit
Option Base 1

'///////////////////////////////////////////////////////////
'UI初期化_定義
'///////////////////////////////////////////////////////////
Public Sub configUiDesign(ByRef Disp As Object)
    '文字列系の設定用Helper
    Call configLabelName(Disp)
    
    'サイズや位置など設定用Helper
    Call configSizePosition(Disp)
End Sub

Private Sub configLabelName(Disp)
    With Disp
        .Caption = "点検画面"
        .LabelTab1.Caption = "一覧データ"
        .LabelTab2.Caption = "History"
        .LabelTab3.Caption = "Checking"
        .LabelTab4.Caption = "Inspection"
        .LabelTab5.Caption = "Closing"
    End With
End Sub

Private Sub configSizePosition(Disp)
    With Disp
        .StartUpPosition = 0
        .Left = Application.Left + 10
        .Top = Application.Top + 10
        .Height = 520
        .Width = 1100
        
        'ラベル類初期化
        With .LabelBack
            .Top = 12
            .Height = 24
            .Left = 12
            .Width = 54
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .Font.Bold = True
            .BorderStyle = fmBorderStyleNone
            .BorderColor = &H8000000D
        End With
        With .LabelTab1
            .Top = 42
            .Height = 24
            .Left = 54
            .Width = 200
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .Font.Bold = True
            .BackColor = RGB(255, 255, 255)
            .ForeColor = RGB(0, 0, 0)
        End With
        With .LabelTab2
            .Top = 42
            .Height = 24
            .Left = 254
            .Width = 200
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .Font.Bold = True
            .BackColor = RGB(255, 255, 255)
            .ForeColor = RGB(0, 0, 0)
        End With
        With .LabelTab3
            .Top = 42
            .Height = 24
            .Left = 454
            .Width = 200
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .Font.Bold = True
            .BackColor = RGB(255, 255, 255)
            .ForeColor = RGB(0, 0, 0)
        End With
        With .LabelTab4
            .Top = 42
            .Height = 24
            .Left = 654
            .Width = 200
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .Font.Bold = True
            .BackColor = RGB(68, 114, 196)
            .ForeColor = RGB(255, 255, 255)
        End With
        With .LabelTab5
            .Top = 42
            .Height = 24
            .Left = 854
            .Width = 200
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .Font.Bold = True
            .BackColor = RGB(255, 255, 255)
            .ForeColor = RGB(0, 0, 0)
        End With
    End With
End Sub
