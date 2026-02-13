VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListViewDrawer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ListViewDrawer.cls
Option Explicit
Option Base 1

Private lv As MSComctlLib.ListView

Private listHeaders As Variant
Private widths      As Variant
Private idMapDic    As Object

'初期化
Public Sub init(ByRef targetLv As MSComctlLib.ListView, _
                ByVal listHeaderArray As Variant, _
                ByVal widthArray As Variant)
    Set lv = targetLv
    listHeaders = listHeaderArray
    widths = widthArray
End Sub

'データ描画
Public Sub Draw(Data As Variant)
    Dim itm As ListItem
    
    Set idMapDic = CreateObject("Scripting.Dictionary")
    
    With lv
        .ColumnHeaders.Clear
        .ListItems.Clear
        
        'ヘッダー追加_幅設定
        Dim i As Long, j As Long
        
        .ColumnHeaders.Add , , "id", widths(1)
        For i = 2 To UBound(listHeaders, 2)
            .ColumnHeaders.Add , , listHeaders(1, i), widths(i)
        Next i
        
        'データ追加
        If Not IsEmpty(Data) Then
            For i = LBound(Data, 1) To UBound(Data, 1)
                '当日Id<->全体Idの対応マッピング
                idMapDic.Add i, Data(i, COL_A)
                
                Set itm = .ListItems.Add(, , i)
                For j = 2 To UBound(Data, 2)
                    itm.ListSubItems.Add , , Data(i, j)
                Next j
            Next i
        End If
    End With
End Sub

Public Function getSelectedListViewRow(ByRef lv As ListView) As Variant
    Dim i        As Long
    Dim rowArr() As String
    
    If lv.SelectedItem Is Nothing Then Exit Function

    ReDim rowArr(1 To lv.ColumnHeaders.Count)
    rowArr(1) = lv.SelectedItem.Text
    For i = 1 To lv.ColumnHeaders.Count - 1
        rowArr(i + 1) = lv.SelectedItem.ListSubItems(i).Text
    Next i
    
    getSelectedListViewRow = rowArr
End Function

'当日Id->全体Id変換
Public Property Get getGeneralId(ByVal dailyId As Long) As String
    getGeneralId = idMapDic(dailyId)
End Property
