VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CustomerDetailDisp 
   Caption         =   "UserForm1"
   ClientHeight    =   48
   ClientLeft      =   -324
   ClientTop       =   -1056
   ClientWidth     =   60
   OleObjectBlob   =   "CustomerDetailDisp.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CustomerDetailDisp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

'ListViewのクラス宣言
Private drawer As ListViewDrawer

Private nameVal As String
Private telVal As String
Private targetArr As Variant

'/////////////////////////////////////////////////////////// Windows API の宣言
Private Declare PtrSafe Function LoadCursor Lib "user32" Alias "LoadCursorA" _
    (ByVal hInstance As LongPtr, ByVal lpCursorName As Long) As LongPtr
Private Declare PtrSafe Function SetCursor Lib "user32" _
    (ByVal hCursor As LongPtr) As LongPtr
Private Const IDC_HAND = 32649
'//////////////////////////////////////////////////////////

Private Sub UserForm_Initialize()
    Call UiConfig_CustomerDetailDisp.configUiDesign(Me)
End Sub

Public Sub setupDetail(ByVal targetName As String, ByVal targetTel As String)
    
    nameVal = targetName
    telVal = targetTel
    If nameVal = "" Or telVal = "" Then Exit Sub
    targetArr = searchByNameAndTel(nameVal, telVal)
    
    Me.TextBoxName.Text = nameVal
    Me.TextBoxTel.Text = telVal
    Me.LabelRootVal.Caption = targetArr(LBound(targetArr, 1), COL_ROOT)
    Me.LabelCtVal.Caption = UBound(targetArr, 1)
    Me.LabelNgVal.Caption = joinStrByCol(targetArr, COL_NG)
    Me.LabelNotesVal.Caption = joinStrByCol(targetArr, COL_NOTE)
    
    'ListViewクラス更新
    Set drawer = New ListViewDrawer
    Dim widths As Variant: widths = UiConfig_CustomerDetailDisp.configCustomerDataView(widths)
    Dim arrHeader As Variant: arrHeader = DispMod.getArrHeader()
    Dim arrData As Variant: arrData = targetArr
    Call drawer.init(Me.ListView1, arrHeader, widths)
    Call drawer.Draw(arrData)
End Sub
Private Function searchByNameAndTel(ByVal targetName As String, _
                                    ByVal targetTel As String) As Variant
    Dim arr As Variant: arr = DispMod.searchCustomerByNameAndTel(targetName, targetTel)
    
    '日付列が文字列だとやっかい、一度数値化する
    Dim dateNormalizedArr As Variant: dateNormalizedArr = DispMod.normalizeDateCol(arr)
    'ソート処理
    Dim sortedArr As Variant
    sortedArr = quickSort2dByColumn(dateNormalizedArr, COL_DATE)
    
    'return
    searchByNameAndTel = sortedArr
End Function
Private Function joinStrByCol(ByVal targetArr As Variant, _
                              ByVal targetCol As Long) As Variant
    Dim resultStr As String
    Dim i As Long
    For i = LBound(targetArr, 1) To UBound(targetArr, 1)
        If i <> LBound(targetArr, 1) _
        And targetArr(i, targetCol) <> "" Then
            resultStr = resultStr & ", "
        End If
        
        resultStr = resultStr & CStr(targetArr(i, targetCol))
    Next i
    
    joinStrByCol = resultStr
End Function
    
'///////////////////////////////////////////////////////////
'遷移系ボタン操作
'///////////////////////////////////////////////////////////
Private Sub LabelBack_Click()
    Me.Hide
    goBack
End Sub
'新規入力画面へ遷移する
Private Sub CommandButtonAdd_Click()
    Me.Hide
    Me.TextBoxName.Text = nameVal
    Me.TextBoxTel.Text = telVal
    
    Call InputFormDisp.reloadInputs(Date, _
                                    UBound(DispMod.getArrBody(Date), 1), _
                                    nameVal, _
                                    telVal)
    navigateTo InputFormDisp
End Sub
'情報変更機能
Private Sub CommandButtonChange_Click()
    If TextBoxName.Text = nameVal And TextBoxTel.Text = telVal Then
        MsgBox prompt:="「名前」と「番号」入力欄に変更がありません。" & vbCrLf & vbCrLf & _
                       "このボタンから、現在表示中のユーザーの全履歴の「名前」と「番号」を一括変更できます。", _
               Buttons:=vbInformation, Title:="機能Help"
        Exit Sub
    End If
    
    If (MsgBox(prompt:="表示中のユーザーの「名前」と「番号」を変更します。" & vbCrLf & _
                       "よろしいですか？", _
               Buttons:=vbYesNo + vbExclamation, Title:="情報変更") = vbNo) Then
        Exit Sub
    End If
    
    Dim targetUsersArr As Variant: targetUsersArr = DispMod.searchCustomerByNameAndTel(TextBoxName.Text, TextBoxTel.Text)
    Dim existsCustomer As Boolean: existsCustomer = (Not IsEmpty(targetUsersArr))
    If existsCustomer Then
        If (MsgBox(prompt:="「名前」と「番号」が同一のユーザーが既に存在します。" & vbCrLf & _
                           "以降、現在表示されている元【" & nameVal & "】さんの履歴と、既存の【" & TextBoxName.Text & "】さんの履歴を統合し、同一人物として扱います。" & vbCrLf & _
                           "本当によろしいですか？" & vbCrLf & vbCrLf & _
                           "変更前: " & nameVal & "|" & telVal & vbCrLf & _
                           "↓" & vbCrLf & _
                           "変更後: " & TextBoxName.Text & "|" & TextBoxTel.Text, _
                   Buttons:=vbYesNo + vbExclamation, Title:="注意") = vbNo) Then
            Exit Sub
        End If
    End If
    
    
    Call changeInformation(targetArr, TextBoxName.Text, TextBoxTel.Text)
    
    Call FrontDataDisp.FrontDataUpdate(Date)
    Call CustomerDetailDisp.setupDetail(TextBoxName.Text, TextBoxTel.Text)
    Call HistoryDataDisp.api_searchHistory(HistoryDataDisp.TextBoxFree.Text)
    
    nameVal = TextBoxName.Text
    telVal = TextBoxTel.Text
    
    MsgBox "「名前」と「番号」を更新しました。", Title:="完了"
End Sub
'リストビュー内レコードダブルクリック
Private Sub ListView1_dblClick()
    If Me.ListView1.SelectedItem Is Nothing Then Exit Sub
    
    Dim rowList As Variant
    rowList = drawer.getSelectedListViewRow(Me.ListView1)
    Me.Hide
    Call InputFormDisp.editInputs(rowList, _
                                  ListView1.SelectedItem.Text, _
                                  drawer.getGeneralId(rowList(1)) _
    )
    
    navigateTo InputFormDisp
End Sub

'データ(実体)修正処理
Private Sub changeInformation(ByVal targetArr As Variant, _
                              ByVal name As String, _
                              ByVal tel As String)
    With InputSh
        If .ProtectContents Then .Unprotect "042595"
        
        Dim lastRow As Long: lastRow = .Cells(.Rows.Count, COL_A).End(xlUp).Row
        
        Dim i As Long
        For i = LBound(targetArr, 1) To UBound(targetArr, 1)
            Dim j As Long
            For j = 1 To lastRow
                If targetArr(i, COL_A) = .Cells(j, COL_A) Then
                    .Cells(j, COL_CUSTM) = name
                    .Cells(j, COL_TEL) = tel
                End If
            Next j
        Next i
        
        If Not .ProtectContents Then .Protect "042595"
        
    End With
End Sub

'閉じるボタン_フォームオブジェクトクリア
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Call DispMod.clearingForms
    End If
End Sub

'Labelのマウスオーバー関連(WindowsAPI使用)
Private Sub LabelBack_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                             ByVal X As Single, ByVal Y As Single)
    Dim hCursor As LongPtr
    hCursor = LoadCursor(0, IDC_HAND)
    SetCursor hCursor
    LabelBack.BorderStyle = fmBorderStyleSingle
    LabelBack.BorderColor = &H8000000D
End Sub
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                               ByVal X As Single, ByVal Y As Single)
    SetCursor LoadCursor(0, 32512) ' IDC_ARROW（通常の矢印カーソル）
    LabelBack.BorderStyle = fmBorderStyleNone
End Sub
