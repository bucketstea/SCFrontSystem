Attribute VB_Name = "UiConfig_AdvancedSearchDisp"
Option Explicit
Option Base 1

'///////////////////////////////////////////////////////////
'UI初期化_定義
'///////////////////////////////////////////////////////////
Public Sub configUiDesign(ByRef AdvancedSearchDisp As Object)
    '文字列系の設定用Helper
    Call configLabelName(AdvancedSearchDisp)
    
    'サイズや位置など設定用Helper
    Call configSizePosition(AdvancedSearchDisp)
End Sub
Private Sub configLabelName(AdvancedSearchDisp)
    With AdvancedSearchDisp
        .Caption = "高度な検索"
    End With
End Sub
Private Sub configSizePosition(AdvancedSearchDisp)
    With AdvancedSearchDisp
        .Height = 400
        .Width = 600
        
        'ラベル類初期化
'        With .LabelBack
'            .Top = 12
'            .Height = 24
'            .Left = 12
'            .Width = 54
'            .BorderStyle = fmBorderStyleNone
'            .BorderColor = &H8000000D
'        End With
    End With
End Sub
