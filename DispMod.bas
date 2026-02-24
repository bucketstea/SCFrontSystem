Attribute VB_Name = "DispMod"
Option Explicit
Option Base 1

'定数_定義
'listView用ヘッダーIndex/帳票ヘッダーIndex
Public Const COL_A     As Long = 1
Public Const COL_ID    As Long = 1
Public Const COL_COUNT As Long = 2
Public Const COL_DATE  As Long = 3
Public Const COL_NEW   As Long = 4
Public Const COL_ROOT  As Long = 5
Public Const COL_TYPE  As Long = 6
Public Const COL_STAFF As Long = 7
Public Const COL_CUSTM As Long = 8
Public Const COL_TEL   As Long = 9
Public Const COL_NG    As Long = 10
Public Const COL_NOTE  As Long = 11
Public Const COL_DEST  As Long = 12
Public Const COL_SERV  As Long = 13
Public Const COL_COURS As Long = 14
Public Const COL_EXPAN As Long = 15
Public Const COL_OP    As Long = 16
Public Const COL_TIME  As Long = 17
Public Const COL_SALES As Long = 18
Public Const COL_CCOST As Long = 19
Public Const COL_PROFI As Long = 20
Public Const COL_QBACK As Long = 21
Public Const COL_SBACK As Long = 22
Public Const COL_LAST  As Long = 22

'シートオブジェクト
Public InputSh As Worksheet

Public Sub dispFrontData()
    'スプラッシュ画面表示
    Splash.Show vbModeless: DoEvents: Application.Wait Now + TimeValue("0:00:00")
    
    Call init
    
    Unload Splash
    
    'メイン画面表示
    showStart HomeDisp
End Sub

Public Sub init()
    Set InputSh = ThisWorkbook.Sheets("入力シート")
    If InputSh.ProtectContents Then InputSh.Unprotect "042595"
    
    Call replaceAllLineBreaks
    
    If Not InputSh.ProtectContents Then InputSh.Protect "042595"
End Sub
'改行除去
Private Sub replaceAllLineBreaks()
    Dim c As Range
    Dim s As String

    For Each c In InputSh.UsedRange.Cells
        If Not IsError(c.value) Then
            s = CStr(c.value)
            
            If InStr(s, vbLf) > 0 Or InStr(s, vbCr) > 0 Then
                c.NumberFormat = "@"
                
                s = Replace(s, vbCr, "")
                s = Replace(s, vbLf, "")
                
                c.value = s
            End If
        End If
    Next c
End Sub

Public Sub clearingForms()
    Dim uf As Object
    For Each uf In VBA.UserForms
        Unload uf
    Next uf
    Set uf = Nothing
End Sub

'ヘッダー取得
Public Function getArrHeader() As Variant
    Dim arrHeader As Variant
    arrHeader = InputSh.Range(InputSh.Cells(1, COL_A), InputSh.Cells(1, COL_LAST)).value
    
    'return
    getArrHeader = arrHeader
End Function

'データボディ取得
Public Function getArrBody(ByVal dd As Date) As Variant
    Dim lastRow As Long
    lastRow = InputSh.Cells(InputSh.Rows.Count, COL_DATE).End(xlUp).Row
    Dim arrBodyAll As Variant
    arrBodyAll = InputSh.Range(InputSh.Cells(2, COL_A), InputSh.Cells(lastRow, COL_LAST)).value
    Dim arrBodyAtDate As Variant: arrBodyAtDate = getDayRangeArray(arrBodyAll, dd)
    
    If Not IsEmpty(arrBodyAtDate) Then
        'データ形式修正
        Dim i As Long
        For i = LBound(arrBodyAtDate, 1) To UBound(arrBodyAtDate, 1)
            arrBodyAtDate(i, COL_TIME) = Format(arrBodyAtDate(i, COL_TIME), "hh:mm")
        Next i
    End If
    
    'return
    getArrBody = arrBodyAtDate
End Function

'///////////////////////////////////////////////////////////
'日付検索
'各日データ表示に使用する
'下から探索したいため専用関数で実装
'///////////////////////////////////////////////////////////
Private Function getDayRangeArray(ByVal arrBody As Variant, _
                                  ByVal dd As Date) As Variant
    Dim i As Long, j As Long
    Dim strDateYymmdd As String: strDateYymmdd = Format(dd, "yymmdd")
    Dim lowerColumn As Long: lowerColumn = LBound(arrBody, 2)
    Dim upperColumn As Long: upperColumn = UBound(arrBody, 2)
    
    '下(最新)から線形探索して日付ヒットしたらコレクションに入れていく
    Dim resultColl As New Collection
    For i = UBound(arrBody, 1) To LBound(arrBody, 1) Step -1
        If CStr(arrBody(i, COL_DATE)) = CStr(strDateYymmdd) Then
            Dim rowArr As Variant
            ReDim rowArr(1 To upperColumn)
            For j = lowerColumn To upperColumn
                rowArr(j) = arrBody(i, j)
            Next j
            resultColl.Add rowArr
        End If
    Next i
    
    'コレクションからresultの2次元配列に詰めなおしていく
    If resultColl.Count = 0 Then
        getDayRangeArray = Empty 'return
    Else
        Dim resultArr As Variant: ReDim resultArr(1 To resultColl.Count, lowerColumn To upperColumn)
        Dim targetRow As Long: targetRow = 1 '下から取った配列を上から詰めなおすためiとtargetRowを分けている
        For i = resultColl.Count To 1 Step -1
            rowArr = resultColl(i)
            For j = lowerColumn To upperColumn
                resultArr(targetRow, j) = rowArr(j)
            Next j
            targetRow = targetRow + 1
        Next i
        getDayRangeArray = resultArr 'return
    End If
End Function

'///////////////////////////////////////////////////////////
'AfromB検索
'履歴表示に使用する
' "_del"レコード(フォーム操作による削除済み)はここで取り除く
'///////////////////////////////////////////////////////////
Public Function searchAfromB(ByVal Aval As String, _
                             ByVal Bcol As Long) As Variant
    Dim allData As Variant: allData = getAllData()
    Dim record As Variant: ReDim record(1 To UBound(allData, 2))
    Dim records As Collection: Set records = New Collection
    
    Dim i As Long
    For i = UBound(allData, 1) To LBound(allData, 1) Step -1
        If Not Right(allData(i, COL_A), 4) = "_del" Then
            If allData(i, Bcol) Like "*" & Aval & "*" Then
                Dim j As Long
                For j = LBound(record) To UBound(record)
                    record(j) = allData(i, j)
                Next j
                '時刻表示崩れるのを修正
                record(COL_TIME) = Format(record(COL_TIME), "hh:mm")
                
                records.Add record
            End If
        End If
    Next i
    
    If records.Count < 1 Then Exit Function
    
    Dim resultArr As Variant: resultArr = convertCollectionToArray(records)
    
    'return
    searchAfromB = resultArr
End Function

'///////////////////////////////////////////////////////////
'名前&番号検索
'履歴表示に使用する
' "_del"レコード(フォーム操作による削除済み)はここで取り除く
'///////////////////////////////////////////////////////////
Public Function searchCustomerByNameAndTel(ByVal targetName As String, _
                                           ByVal targetTel As String) As Variant
    Dim allData As Variant: allData = getAllData()
    Dim record As Variant: ReDim record(1 To UBound(allData, 2))
    Dim records As Collection: Set records = New Collection
    
    Dim i As Long
    For i = UBound(allData, 1) To LBound(allData, 1) Step -1
        If Not Right(allData(i, COL_A), 4) = "_del" Then
            If allData(i, COL_CUSTM) = targetName _
            And allData(i, COL_TEL) = targetTel Then
                Dim j As Long
                For j = LBound(record) To UBound(record)
                    record(j) = allData(i, j)
                Next j
                '時刻表示崩れるのを修正
                record(COL_TIME) = Format(record(COL_TIME), "hh:mm")
                
                records.Add record
            End If
        End If
    Next i
    
    If records.Count < 1 Then Exit Function
    
    Dim resultArr As Variant: resultArr = convertCollectionToArray(records)
    
    'return
    searchCustomerByNameAndTel = resultArr
End Function

'日付の数値化処理_日付ソートが文字列では不安定そうなので、特にソート前に噛ます推奨
Public Function normalizeDateCol(arr)
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        If IsNumeric(arr(i, COL_DATE)) Then arr(i, COL_DATE) = CLng(arr(i, COL_DATE))
    Next i
    normalizeDateCol = arr
End Function

'///////////////////////////////////////////////////////////
'全件取得
'///////////////////////////////////////////////////////////
Public Function getAllData() As Variant
    Dim lastRow As Long: lastRow = InputSh.Cells(InputSh.Rows.Count, COL_DATE).End(xlUp).Row
    Dim lastCol As Long: lastCol = COL_SBACK
    
    'return
    getAllData = Range(InputSh.Cells(1, 1), InputSh.Cells(lastRow, lastCol))
End Function
'///////////////////////////////////////////////////////////
'NameとTelでDic化
' key = generalId | value = name, tel
'///////////////////////////////////////////////////////////
Public Function summaryByCustom(ByVal arr As Variant) As Object
    Dim customDic As Object: Set customDic = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Dim id     As Long:   id = arr(i, COL_A)
        Dim name As String: name = arr(i, COL_CUSTM)
        Dim tel  As String:  tel = arr(i, COL_TEL)
        
        If Not customDic.exists(name & "|" & tel) Then
            customDic.Add name & "|" & tel, Array(id, name, tel)
        End If
    Next i
    Set summaryByCustom = customDic
End Function

'///////////////////////////////////////////////////////////
'各種ユーザー関数
'///////////////////////////////////////////////////////////
'format関数などで文字列化した日付をDate型に変換する
Public Function parseYymmdd(ByVal str As String) As Date
     Dim yyyy As Long
     Dim mm As Long
     Dim dd As Long
    
     yyyy = 2000 + CInt(Left(str, 2))
     mm = CInt(Mid(str, 3, 2))
     dd = CInt(Right(str, 2))
    
     parseYymmdd = DateSerial(yyyy, mm, dd)
End Function
Public Function parseYyyySmmSdd(ByVal str As String) As Date
     Dim yyyy As Long
     Dim mm As Long
     Dim dd As Long

     yyyy = CInt(Left(str, 4))
     mm = CInt(Mid(str, 6, 2))
     dd = CInt(Right(str, 2))

     parseYyyySmmSdd = DateSerial(yyyy, mm, dd)
End Function
Public Function convDateForFront(ByVal dd As Date) As String
    convDateForFront = Format(dd, "yyyy/mm/dd")
End Function
