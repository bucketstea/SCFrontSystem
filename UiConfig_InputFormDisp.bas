Attribute VB_Name = "UiConfig_InputFormDisp"
Option Explicit
Option Base 1

'///////////////////////////////////////////////////////////
'UI初期化_定義
'///////////////////////////////////////////////////////////
Public Sub configUiDesign(ByRef InputFormDisp As Object)
    'コントロールUI生成
'    Call addUi(InputFormDisp)
    
    '文字列系の設定用Helper
    Call configLabelName(InputFormDisp)
    
    'サイズや位置など設定用Helper
    Call configSizePosition(InputFormDisp)
End Sub
Private Sub addUi(ByRef InputFormDisp As Object)
    With InputFormDisp
'        .Controls.Add "Forms.Label.1", "LabelBase"
'        .Controls.Add "Forms.Label.1", "LabelCustomer"
'        .Controls.Add "Forms.Label.1", "LabelUse"
'        .Controls.Add "Forms.Label.1", "LabelAccount"
    End With
End Sub
Private Sub configLabelName(ByRef InputFormDisp As Object)
    With InputFormDisp
        .Caption = "データ入力画面"
        
        '大項目名
        .LabelBase.Caption = "label1"
        .LabelCustomer.Caption = "label2"
        .LabelUse.Caption = "label3"
        .LabelAccount.Caption = "label4"
        
        '各小項目名設定
        .LabelDate.Caption = "日付"
        .LabelAd.Caption = "labelad"
        .LabelType.Caption = "labeltype"
        .LabelTime.Caption = "labeltime"
        .LabelName.Caption = "labelname"
        .LabelTel.Caption = "labeltel"
        .LabelNg.Caption = "LabelNG"
        .LabelNotes.Caption = "labelnotes"
        .LabelCast.Caption = "labelcast"
        .LabelCourse.Caption = "labelcourse"
        .LabelService.Caption = "labelservice"
        .LabelOP.Caption = "labelop"
        .LabelDestination.Caption = "labeldestination"
        .LabelExpand.Caption = "labelexpand"
        .LabelSales.Caption = "labelsales"
        .LabelCost.Caption = "labelcost"
        .LabelProfit.Caption = "labelprofit"
        .LabelQB.Caption = "labelqb"
        .LabelSB.Caption = "labelsb"
        
        'プレースホルダ
        .PlaceholderDate.Caption = Format(Date, "yymmdd")
        .PlaceholderTime.Caption = CStr(Format(Now, "hh:mm"))
        .PlaceholderName.Caption = "オガワアツシ"
        .PlaceholderTel.Caption = "09012345678"
        .PlaceholderCast.Caption = "あつこ"
        .PlaceholderCourse.Caption = "60"
        .PlaceholderService.Caption = "service"
        .PlaceholderOP.Caption = "op"
        .PlaceholderDestination.Caption = "destination"
        .PlaceholderExpand.Caption = "expand"
        .PlaceholderSales.Caption = "16000"
        .PlaceholderCost.Caption = "8000"
        .PlaceholderQB.Caption = "1000"
        .PlaceholderSB.Caption = "20"
        
        'エラーメッセージ(デフォルト)
        .LabelErrorDate.Caption = "入力してください。"
        .LabelErrorAd.Caption = "選択してください。"
        .LabelErrorType.Caption = "選択してください。"
        .LabelErrorName.Caption = "入力してください。"
        .LabelErrorTel.Caption = "入力してください。"
        .LabelErrorSales.Caption = "入力してください。"
        .LabelErrorCost.Caption = "入力してください。"
        
        
        'その他
        .CommandButtonSearchName.Caption = "search"
        .CommandButtonSearchTel.Caption = "search"
    End With
End Sub

Private Sub configSizePosition(ByRef InputFormDisp As Object)
    With InputFormDisp
        .StartUpPosition = 0
        .Left = Application.Left + 10
        .Top = Application.Top + 10
        .Height = 520
        .Width = 1100
        
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
        With .LabelEditId
            .Top = 30
            .Height = 24
            .Left = 174
            .Width = 150
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
        End With
        With .ImageDelete
            .Top = 12
            .Left = 1048
            .BorderStyle = fmBorderStyleNone
            .BorderColor = &HFF&
        End With
        '入力欄グループ
        With .LabelBase
            .Top = 78
            .Height = 30
            .Left = 24
            .Width = 126
            .Font.name = "Yu Gothic UI"
            .Font.Size = 20
            .TextAlign = fmTextAlignCenter
        End With
        With .LabelCustomer
            .Top = 168
            .Height = 30
            .Left = 24
            .Width = 126
            .Font.name = "Yu Gothic UI"
            .Font.Size = 20
            .TextAlign = fmTextAlignCenter
        End With
        With .LabelUse
            .Top = 258
            .Height = 30
            .Left = 24
            .Width = 126
            .Font.name = "Yu Gothic UI"
            .Font.Size = 20
            .TextAlign = fmTextAlignCenter
        End With
        With .LabelAccount
            .Top = 348
            .Height = 30
            .Left = 24
            .Width = 126
            .Font.name = "Yu Gothic UI"
            .Font.Size = 20
            .TextAlign = fmTextAlignCenter
        End With
        '//入力欄/////////////////////////////////////////////////////////
        'Date入力
        With .LabelDate
            .Top = 66
            .Left = 174
            .Height = 18
            .Width = 132
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
        End With
        With .TextBoxDate
            .TabStop = True
            .TabIndex = 21
            .Top = 84
            .Left = 174
            .Height = 34
            .Width = 132
            .SpecialEffect = fmSpecialEffectEtched
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
        End With
        With .PlaceholderDate
            .Top = 88
            .Left = 180
            .Height = 21
            .Width = 120
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .BackColor = &H8000000E
            .ForeColor = &H80000010
        End With
        With .LabelErrorDate
            .Top = 120
            .Left = 174
            .Height = 12
            .Width = 132
            .ForeColor = &HFF&
            .Visible = False
            .Font.name = "Yu Gothic UI"
            .Font.Size = 8
        End With
        'AD入力
        With .LabelAd
            .Top = 66
            .Left = 324
            .Height = 18
            .Width = 132
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
        End With
        With .ComboBoxAd
            .TabStop = True
            .TabIndex = 1
            .Top = 84
            .Left = 324
            .Height = 34
            .Width = 132
            .SpecialEffect = fmSpecialEffectEtched
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
        End With
        With .LabelErrorAd
            .Top = 120
            .Left = 324
            .Height = 12
            .Width = 84
            .ForeColor = &HFF&
            .Visible = False
            .Font.name = "Yu Gothic UI"
            .Font.Size = 8
        End With
        'Type入力
        With .LabelType
            .Top = 66
            .Left = 474
            .Height = 18
            .Width = 132
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
        End With
        With .ComboBoxType
            .TabStop = True
            .TabIndex = 2
            .Top = 84
            .Left = 474
            .Height = 34
            .Width = 132
            .SpecialEffect = fmSpecialEffectEtched
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
        End With
        With .LabelErrorType
            .Top = 120
            .Left = 474
            .Height = 12
            .Width = 84
            .ForeColor = &HFF&
            .Visible = False
            .Font.name = "Yu Gothic UI"
            .Font.Size = 8
        End With
        'Time入力
        With .LabelTime
            .Top = 66
            .Left = 624
            .Height = 18
            .Width = 132
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
        End With
        With .TextBoxTime
            .TabStop = True
            .TabIndex = 3
            .Top = 84
            .Left = 624
            .Height = 34
            .Width = 132
            .SpecialEffect = fmSpecialEffectEtched
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
        End With
        With .PlaceholderTime
            .Top = 88
            .Left = 630
            .Height = 21
            .Width = 120
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .BackColor = &H8000000E
            .ForeColor = &H80000010
        End With
        'Name入力
        With .LabelName
            .Top = 156
            .Left = 174
            .Height = 18
            .Width = 132
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
        End With
        With .TextBoxName
            .TabStop = True
            .TabIndex = 4
            .Top = 174
            .Left = 174
            .Height = 34
            .Width = 132
            .SpecialEffect = fmSpecialEffectEtched
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
        End With
        With .PlaceholderName
            .Top = 178
            .Left = 180
            .Height = 21
            .Width = 120
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .BackColor = &H8000000E
            .ForeColor = &H80000010
        End With
        With .LabelErrorName
            .Top = 210
            .Left = 174
            .Height = 12
            .Width = 84
            .ForeColor = &HFF&
            .Visible = False
            .Font.name = "Yu Gothic UI"
            .Font.Size = 8
        End With
        'Tel
        With .LabelTel
            .Top = 156
            .Left = 324
            .Height = 18
            .Width = 84
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
        End With
        With .TextBoxTel
            .TabStop = True
            .TabIndex = 6
            .Top = 174
            .Left = 324
            .Height = 34
            .Width = 132
            .SpecialEffect = fmSpecialEffectEtched
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
        End With
        With .PlaceholderTel
            .Top = 178
            .Left = 330
            .Height = 21
            .Width = 120
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .BackColor = &H8000000E
            .ForeColor = &H80000010
        End With
        With .LabelErrorTel
            .Top = 210
            .Left = 324
            .Height = 12
            .Width = 84
            .ForeColor = &HFF&
            .Visible = False
            .Font.name = "Yu Gothic UI"
            .Font.Size = 8
        End With
        'NG
        With .LabelNg
            .Top = 156
            .Left = 474
            .Height = 18
            .Width = 132
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
        End With
        With .TextBoxNG
            .TabStop = True
            .TabIndex = 8
            .Top = 174
            .Left = 474
            .Height = 54
            .Width = 282
            .SpecialEffect = fmSpecialEffectEtched
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .Multiline = True
            .ScrollBars = fmScrollBarsVertical
        End With
        
        'Notes
        With .LabelNotes
            .Top = 156
            .Left = 774
            .Height = 18
            .Width = 132
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
        End With
        With .TextBoxNotes
            .TabStop = True
            .TabIndex = 9
            .Top = 174
            .Left = 774
            .Height = 54
            .Width = 282
            .SpecialEffect = fmSpecialEffectEtched
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .Multiline = True
            .ScrollBars = fmScrollBarsVertical
        End With
        
        'Cast
        With .LabelCast
            .Top = 246
            .Left = 174
            .Height = 18
            .Width = 132
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
        End With
        With .TextBoxCast
            .TabStop = True
            .TabIndex = 10
            .Top = 264
            .Left = 174
            .Height = 34
            .Width = 132
            .SpecialEffect = fmSpecialEffectEtched
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
        End With
        With .PlaceholderCast
            .Top = 268
            .Left = 180
            .Height = 21
            .Width = 120
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .BackColor = &H8000000E
            .ForeColor = &H80000010
        End With
        'Course
        With .LabelCourse
            .Top = 246
            .Left = 324
            .Height = 18
            .Width = 132
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
        End With
        With .TextBoxCourse
            .TabStop = True
            .TabIndex = 11
            .Top = 264
            .Left = 324
            .Height = 34
            .Width = 132
            .SpecialEffect = fmSpecialEffectEtched
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
        End With
        With .PlaceholderCourse
            .Top = 268
            .Left = 330
            .Height = 21
            .Width = 120
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .BackColor = &H8000000E
            .ForeColor = &H80000010
        End With
        'Service
        With .LabelService
            .Top = 246
            .Left = 474
            .Height = 18
            .Width = 132
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
        End With
        With .TextBoxService
            .TabStop = True
            .TabIndex = 12
            .Top = 264
            .Left = 474
            .Height = 34
            .Width = 132
            .SpecialEffect = fmSpecialEffectEtched
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
        End With
        With .PlaceholderService
            .Top = 268
            .Left = 480
            .Height = 21
            .Width = 120
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .BackColor = &H8000000E
            .ForeColor = &H80000010
        End With
        'OP
        With .LabelOP
            .Top = 246
            .Left = 624
            .Height = 18
            .Width = 132
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
        End With
        With .TextBoxOP
            .TabStop = True
            .TabIndex = 13
            .Top = 264
            .Left = 624
            .Height = 34
            .Width = 132
            .SpecialEffect = fmSpecialEffectEtched
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
        End With
        With .PlaceholderOP
            .Top = 268
            .Left = 630
            .Height = 21
            .Width = 120
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .BackColor = &H8000000E
            .ForeColor = &H80000010
        End With
        'Destination
        With .LabelDestination
            .Top = 246
            .Left = 774
            .Height = 18
            .Width = 132
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
        End With
        With .TextBoxDestination
            .TabStop = True
            .TabIndex = 14
            .Top = 264
            .Left = 774
            .Height = 34
            .Width = 132
            .SpecialEffect = fmSpecialEffectEtched
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
        End With
        With .PlaceholderDestination
            .Top = 268
            .Left = 780
            .Height = 21
            .Width = 120
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .BackColor = &H8000000E
            .ForeColor = &H80000010
        End With
        'Expand
        With .LabelExpand
            .Top = 246
            .Left = 924
            .Height = 18
            .Width = 132
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
        End With
        With .TextBoxExpand
            .TabStop = True
            .TabIndex = 15
            .Top = 264
            .Left = 924
            .Height = 34
            .Width = 132
            .SpecialEffect = fmSpecialEffectEtched
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
        End With
        With .PlaceholderExpand
            .Top = 268
            .Left = 930
            .Height = 21
            .Width = 120
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .BackColor = &H8000000E
            .ForeColor = &H80000010
        End With
        'Sales
        With .LabelSales
            .Top = 336
            .Left = 174
            .Height = 18
            .Width = 84
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
        End With
        With .TextBoxSales
            .TabStop = True
            .TabIndex = 16
            .Top = 354
            .Left = 174
            .Height = 34
            .Width = 132
            .SpecialEffect = fmSpecialEffectEtched
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
        End With
        With .PlaceholderSales
            .Top = 358
            .Left = 180
            .Height = 21
            .Width = 120
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .BackColor = &H8000000E
            .ForeColor = &H80000010
        End With
        With .LabelErrorSales
            .Top = 390
            .Left = 174
            .Height = 12
            .Width = 132
            .ForeColor = &HFF&
            .Visible = False
            .Font.name = "Yu Gothic UI"
            .Font.Size = 8
        End With
        'Cost
        With .LabelCost
            .Top = 336
            .Left = 324
            .Height = 18
            .Width = 84
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
        End With
        With .TextBoxCost
            .TabStop = True
            .TabIndex = 17
            .Top = 354
            .Left = 324
            .Height = 34
            .Width = 132
            .SpecialEffect = fmSpecialEffectEtched
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
        End With
        With .PlaceholderCost
            .Top = 358
            .Left = 330
            .Height = 21
            .Width = 120
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .BackColor = &H8000000E
            .ForeColor = &H80000010
        End With
        With .LabelErrorCost
            .Top = 390
            .Left = 324
            .Height = 12
            .Width = 132
            .ForeColor = &HFF&
            .Visible = False
            .Font.name = "Yu Gothic UI"
            .Font.Size = 8
        End With
        'Profit
        With .LabelProfit
            .Top = 336
            .Left = 474
            .Height = 18
            .Width = 132
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
        End With
        With .TextBoxProfit
            .TabStop = False
            .Enabled = False
            .Top = 354
            .Left = 474
            .Height = 34
            .Width = 132
            .BackColor = RGB(230, 230, 230)
            .SpecialEffect = fmSpecialEffectEtched
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .ForeColor = RGB(255, 255, 255)
        End With
        'QB
        With .LabelQB
            .Top = 336
            .Left = 624
            .Height = 18
            .Width = 132
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
        End With
        With .TextBoxQB
            .TabStop = True
            .TabIndex = 18
            .Top = 354
            .Left = 624
            .Height = 34
            .Width = 132
            .SpecialEffect = fmSpecialEffectEtched
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
        End With
        With .PlaceholderQB
            .Top = 358
            .Left = 630
            .Height = 21
            .Width = 120
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .BackColor = &H8000000E
            .ForeColor = &H80000010
        End With
        'SB
        With .LabelSB
            .Top = 336
            .Left = 774
            .Height = 18
            .Width = 132
            .Font.name = "Yu Gothic UI"
            .Font.Size = 14
        End With
        With .TextBoxSB
            .TabStop = True
            .TabIndex = 19
            .Top = 354
            .Left = 774
            .Height = 34
            .Width = 132
            .SpecialEffect = fmSpecialEffectEtched
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
        End With
        With .PlaceholderSB
            .Top = 358
            .Left = 780
            .Height = 21
            .Width = 120
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .BackColor = &H8000000E
            .ForeColor = &H80000010
        End With
        
        'ボタン初期化
        With .CommandButtonSave
            .TabStop = True
            .TabIndex = 20
            .Top = 450
            .Left = 350
            .Height = 30
            .Width = 400
            .Font.name = "Yu Gothic UI"
            .Font.Size = 16
            .Font.Bold = True
        End With
        With .CommandButtonSearchName
            .TabStop = True
            .TabIndex = 5
            .Top = 204
            .Left = 258
            .Height = 24
            .Width = 48
            .Font.name = "Yu Gothic UI"
            .Font.Size = 11
            .BackColor = RGB(68, 114, 196)
        End With
        With .CommandButtonSearchTel
            .TabStop = True
            .TabIndex = 7
            .Top = 204
            .Left = 408
            .Height = 24
            .Width = 48
            .Font.name = "Yu Gothic UI"
            .Font.Size = 11
            .BackColor = RGB(68, 114, 196)
        End With
    End With
End Sub

Public Sub configComboBox(ByRef InputFormDisp As Object)
    With InputFormDisp
        Dim adsList As Variant
        adsList = Array("経路を選択", _
                        "ad1", _
                        "ad2", _
                        "ad3", _
                        "ad4", _
                        "ad5", _
                        "ad6", _
                        "ad7", _
                        "ad8", _
                        "ad9", _
                        "ad10")
        Dim ad As Variant
        For Each ad In adsList
            .ComboBoxAd.AddItem ad
        Next ad
        
        Dim typList As Variant
        typList = Array("種別を選択", _
                         "type1", _
                         "type2", _
                         "type3", _
                         "type4", _
                         "type5")
        Dim typ As Variant
        For Each typ In typList
            .ComboBoxType.AddItem typ
        Next typ
    End With
End Sub

