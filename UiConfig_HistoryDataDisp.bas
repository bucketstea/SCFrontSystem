Attribute VB_Name = "UiConfig_HistoryDataDisp"
Option Explicit
Option Base 1

Private hitsCt As Long

'///////////////////////////////////////////////////////////
'UI初期化_定義
'///////////////////////////////////////////////////////////
Public Sub configUiDesign(ByRef HistoryDataDisp As Object)
    '文字列系の設定用Helper
    Call configLabelName(HistoryDataDisp)
    
    'サイズや位置など設定用Helper
    Call configSizePosition(HistoryDataDisp)
End Sub
Private Sub configLabelName(HistoryDataDisp)
    With HistoryDataDisp
        .Caption = "履歴検索画面"
        .LabelTab1.Caption = "一覧データ"
        .LabelTab2.Caption = "History"
        .LabelTab3.Caption = "Checking"
        .LabelTab4.Caption = "Inspection"
        .LabelTab5.Caption = "Closing"
        
        hitsCt = 0
        .LabelHits.Caption = hitsCt & " Hit　"
        .LabelFree.Caption = "Freeword:"
        .CommandButtonSearch.Caption = "Search!"
    End With
End Sub
Private Sub configSizePosition(HistoryDataDisp)
    With HistoryDataDisp
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
            .BackColor = RGB(68, 114, 196)
            .ForeColor = RGB(255, 255, 255)
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
        
        With .LabelHits
            .Top = 132
            .Height = 24
            .Left = 10
            .Width = 60
            .TextAlign = fmTextAlignRight
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .Font.Bold = False
        End With
        With .LabelFree
            .Top = 128
            .Height = 24
            .Left = 600
            .Width = 80
            .TextAlign = fmTextAlignRight
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .Font.Bold = True
        End With
        With .TextBoxFree
            .Top = 122
            .Height = 34
            .Left = 684
            .Width = 320
            .SpecialEffect = fmSpecialEffectEtched
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
        End With
        With .CommandButtonSearch
            .Top = 121
            .Height = 34
            .Left = 1004
            .Width = 80
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .Font.Bold = True
            .BackColor = RGB(68, 114, 196)
            .ForeColor = RGB(255, 255, 255)
        End With
        With .ImageAdvance
            .Top = 127
            .Height = 24
            .Left = 970
            .Width = 30
        End With
        
        'ビュー初期化
        With .ListView1
            .Top = 160
            .Left = 5
            .Height = 320
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
    End With
End Sub

Public Function configHistoryDataView(widths) As Variant
    ReDim widths(0 To COL_LAST)
    widths(COL_ID) = 15
    widths(COL_DATE) = 40
    widths(COL_NEW) = 0
    widths(COL_ROOT) = 0
    widths(COL_TYPE) = 40
    widths(COL_STAFF) = 50
    widths(COL_CUSTM) = 80
    widths(COL_TEL) = 70
    widths(COL_NG) = 150
    widths(COL_NOTE) = 550
    widths(COL_DEST) = 0
    widths(COL_SERV) = 0
    widths(COL_COURS) = 40
    widths(COL_EXPAN) = 40
    widths(COL_OP) = 0
    widths(COL_TIME) = 0
    widths(COL_SALES) = 0
    widths(COL_CCOST) = 0
    widths(COL_PROFI) = 0
    widths(COL_QBACK) = 0
    widths(COL_SBACK) = 0
    
    'return
    configHistoryDataView = widths
End Function
