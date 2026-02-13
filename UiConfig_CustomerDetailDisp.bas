Attribute VB_Name = "UiConfig_CustomerDetailDisp"
Option Explicit
Option Base 1

'///////////////////////////////////////////////////////////
'UI初期化_定義
'///////////////////////////////////////////////////////////
Public Sub configUiDesign(ByRef CustomerDetailDisp As Object)
    '文字列系の設定用Helper
    Call configLabelName(CustomerDetailDisp)
    
    'サイズや位置など設定用Helper
    Call configSizePosition(CustomerDetailDisp)
End Sub
Private Sub configLabelName(CustomerDetailDisp)
    With CustomerDetailDisp
        .Caption = "カスタマー詳細画面"
        
        .LabelName.Caption = "Name:"
        .LabelTel.Caption = "Tel:"
        .LabelRoot.Caption = "Root:"
        .LabelCt.Caption = "Ct:"
        .LabelNg.Caption = "Ng:"
        .LabelNotes.Caption = "Notes:"
        
        .CommandButtonAdd.Caption = "Add New"
        .CommandButtonChange.Caption = "Change!"
    End With
End Sub
Private Sub configSizePosition(CustomerDetailDisp)
    With CustomerDetailDisp
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
        With .LabelName
            .Left = 114
            .Top = 30
            .Height = 24
            .Width = 60
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
            .Font.Bold = True
        End With
        With .LabelTel
            .Left = 114
            .Top = 60
            .Height = 24
            .Width = 60
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
            .Font.Bold = True
        End With
        With .LabelRoot
            .Left = 114
            .Top = 90
            .Height = 24
            .Width = 60
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
            .Font.Bold = True
        End With
        With .LabelCt
            .Left = 114
            .Top = 120
            .Height = 24
            .Width = 60
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
            .Font.Bold = True
        End With
        With .LabelNg
            .Left = 402
            .Top = 30
            .Height = 58
            .Width = 60
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
            .Font.Bold = True
        End With
        With .LabelNotes
            .Left = 402
            .Top = 90
            .Height = 58
            .Width = 60
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
            .Font.Bold = True
        End With
        
        With .TextBoxName
            .TabStop = True
            .TabIndex = 1
            .Left = 180
            .Top = 24
            .Height = 28
            .Width = 200
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .Font.Bold = True
            .SpecialEffect = fmSpecialEffectEtched
        End With
        With .TextBoxTel
            .TabStop = True
            .TabIndex = 2
            .Left = 180
            .Top = 54
            .Height = 28
            .Width = 200
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .Font.Bold = True
            .SpecialEffect = fmSpecialEffectEtched
        End With
        With .LabelRootVal
            .Left = 180
            .Top = 84
            .Height = 24
            .Width = 200
            .Font.name = "Yu Gothic UI"
            .Font.Size = 18
            .Font.Bold = False
        End With
        With .LabelCtVal
            .Left = 180
            .Top = 114
            .Height = 24
            .Width = 200
            .Font.name = "Yu Gothic UI"
            .Font.Size = 18
            .Font.Bold = False
        End With
        With .LabelNgVal
            .Left = 468
            .Top = 24
            .Height = 58
            .Width = 450
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
            .Font.Bold = False
        End With
        With .LabelNotesVal
            .Left = 468
            .Top = 84
            .Height = 58
            .Width = 450
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
            .Font.Bold = False
        End With
        
        '新規追加ボタン
        With .CommandButtonAdd
            .TabStop = True
            .TabIndex = 4
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
        '情報変更ボタン
        With .CommandButtonChange
            .TabStop = True
            .TabIndex = 3
            .Top = 82
            .Height = 34
            .Left = 1004
            .Width = 80
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .Font.Bold = True
            .BackColor = RGB(237, 125, 49)
            .ForeColor = RGB(255, 255, 255)
        End With
        
        'ビュー初期化
        With .ListView1
            .TabStop = False
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

Public Function configCustomerDataView(widths) As Variant
    ReDim widths(0 To COL_LAST)
    widths(COL_ID) = 15
    widths(COL_DATE) = 40
    widths(COL_NEW) = 0
    widths(COL_ROOT) = 0
    widths(COL_TYPE) = 40
    widths(COL_STAFF) = 50
    widths(COL_CUSTM) = 0
    widths(COL_TEL) = 0
    widths(COL_NG) = 200
    widths(COL_NOTE) = 300
    widths(COL_DEST) = 50
    widths(COL_SERV) = 150
    widths(COL_COURS) = 40
    widths(COL_EXPAN) = 40
    widths(COL_OP) = 110
    widths(COL_TIME) = 40
    widths(COL_SALES) = 0
    widths(COL_CCOST) = 0
    widths(COL_PROFI) = 0
    widths(COL_QBACK) = 0
    widths(COL_SBACK) = 0
    
    'return
    configCustomerDataView = widths
End Function
