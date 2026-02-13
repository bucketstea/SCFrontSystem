Attribute VB_Name = "UiConfig_CustomerSearchDisp"
Option Explicit
Option Base 1

'///////////////////////////////////////////////////////////
'UI初期化_定義
'///////////////////////////////////////////////////////////
Public Sub configUiDesign(ByRef CustomerSearchDisp As Object)
    'UIの見た目（サイズや位置などの見た目要素）を設定
    With CustomerSearchDisp
        
        .Height = 300
        .Width = 400
        
        '説明文
        With .LabelAnnounce
            .Caption = "複数Hitしました。" & vbCrLf & _
                       "どれかをダブルクリックしてください。"
            .Top = 20
            .Left = 30
            .Height = 40
            .Width = 340
            .Font.name = "Yu Gothic UI"
            .Font.Size = 10
            .Font.Bold = True
        End With
        
        'ビュー初期化
        With .ListView1
            .Top = 60
            .Left = 30
            .Height = 180
            .Width = 340
            .Font.name = "Yu Gothic UI"
            .Font.Size = 9
            .Font.Bold = True
            .View = lvwReport
            .Gridlines = True
            .FullRowSelect = True
            .ListItems.Clear
            .ColumnHeaders.Clear
        End With
    End With
End Sub
