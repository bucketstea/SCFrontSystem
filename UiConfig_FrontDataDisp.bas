Attribute VB_Name = "UiConfig_FrontDataDisp"
Option Explicit
Option Base 1

'///////////////////////////////////////////////////////////
'UI初期化_定義
'///////////////////////////////////////////////////////////
Public Sub configUiDesign(ByRef FrontDataDisp As Object)
    '文字列系の設定用Helper
    Call configLabelName(FrontDataDisp)
    
    'サイズや位置など設定用Helper
    Call configSizePosition(FrontDataDisp)
End Sub
Private Sub configLabelName(FrontDataDisp)
    With FrontDataDisp
        .Caption = "受付データ画面"
        .LabelTab1.Caption = "一覧データ"
        .LabelTab2.Caption = "History"
        .LabelTab3.Caption = "Checking"
        .LabelTab4.Caption = "Inspection"
        .LabelTab5.Caption = "Closing"
        
        .LabelPrevDate.Caption = "<< Prev"
        .LabelNextDate.Caption = "Next >>"
    End With
End Sub
Private Sub configSizePosition(FrontDataDisp)
    With FrontDataDisp
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
            .BackColor = RGB(68, 114, 196)
            .ForeColor = RGB(255, 255, 255)
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
            .BackColor = RGB(255, 255, 255)
            .ForeColor = RGB(0, 0, 0)
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
        With .LabelDate
            .Top = 115
            .Height = 36
            .Left = 470
            .Width = 180
            .Font.name = "Yu Gothic UI"
            .Font.Size = 28
            .Font.Bold = True
            .BorderStyle = fmBorderStyleNone
            .BorderColor = &H8000000D
        End With
        With .LabelPrevDate
            .Top = 120
            .Height = 30
            .Left = 290
            .Width = 120
            .Font.name = "Yu Gothic UI"
            .Font.Size = 20
            .Font.Bold = True
            .BorderStyle = fmBorderStyleNone
            .BorderColor = &H8000000D
        End With
        With .LabelNextDate
            .Top = 120
            .Height = 30
            .Left = 710
            .Width = 120
            .Font.name = "Yu Gothic UI"
            .Font.Size = 20
            .Font.Bold = True
            .BorderStyle = fmBorderStyleNone
            .BorderColor = &H8000000D
        End With
        'ビュー初期化
        With .ListView1
            .Top = 160
            .Left = 5
            .Height = 270
            .Width = 1080
            .Font.name = "Yu Gothic UI"
            .Font.Size = 9
            .Font.Bold = True
            .View = lvwReport
            .Gridlines = True
            .FullRowSelect = True
            .ListItems.Clear
            .ColumnHeaders.Clear
        End With
        'ボタン初期化
        With .CommandButtonInput
            .Caption = "Create"
            .Top = 450
            .Left = 350
            .Height = 30
            .Width = 400
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .Font.Bold = True
            .BackColor = RGB(68, 114, 196)
        End With
    End With
End Sub
'///////////////////////////////////////////////////////////
'ListViewの列幅定義
'///////////////////////////////////////////////////////////
Public Function configFrontDataView(ByVal widths As Variant) As Variant
    ReDim widths(0 To COL_LAST)
    widths(COL_ID) = 15
    widths(COL_DATE) = 40
    widths(COL_NEW) = 0
    widths(COL_ROOT) = 40
    widths(COL_TYPE) = 40
    widths(COL_STAFF) = 50
    widths(COL_CUSTM) = 80
    widths(COL_TEL) = 70
    widths(COL_NG) = 60
    widths(COL_NOTE) = 150
    widths(COL_DEST) = 45
    widths(COL_SERV) = 135
    widths(COL_COURS) = 30
    widths(COL_EXPAN) = 30
    widths(COL_OP) = 60
    widths(COL_TIME) = 40
    widths(COL_SALES) = 40
    widths(COL_CCOST) = 40
    widths(COL_PROFI) = 40
    widths(COL_QBACK) = 40
    widths(COL_SBACK) = 30
    
    'return
    configFrontDataView = widths
End Function
