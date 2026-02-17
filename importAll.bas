Attribute VB_Name = "importAll"
Option Explicit
Sub importAll()
    Dim strPath As String: strPath = ""
    Call ImportVbaSourcesFromFolder(strPath)
End Sub

' === 公開エントリポイント（PowerShellから叩く）===
' 引数 folderPath: 展開したソース格納フォルダ（末尾\ ありなしOK）
Public Sub ImportVbaSourcesFromFolder(ByVal folderPath As String)
    Dim fso As Object, fld As Object, fil As Object
    Dim ext As String, p As String
    Dim vbProj As VBIDE.VBProject

    folderPath = NormalizeFolderPath(folderPath)

    Set vbProj = ThisWorkbook.VBProject
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        Err.Raise vbObjectError + 1, , "Folder not found: " & folderPath
    End If
    Set fld = fso.GetFolder(folderPath)

    ' 1) 既存モジュールを必要に応じて削除（安全側に「標準/クラスのみ」削除）
'    RemoveAllStdAndClassModules vbProj

    ' 2) ファイルを全インポート
    For Each fil In fld.Files
        ext = LCase$(fso.GetExtensionName(fil.Path))
        Select Case ext
            Case "bas", "cls", "frm"
                vbProj.VBComponents.Import fil.Path
            Case Else
                ' ignore
        End Select
    Next fil

    ' 3) フォーム(.frm)がある場合、同名の.frxが必要（同フォルダに置かれていればOK）
    ' 4) 保存
    ThisWorkbook.Save
End Sub

Private Sub RemoveAllStdAndClassModules(ByVal vbProj As VBIDE.VBProject)
    Dim comp As VBIDE.VBComponent
    Dim toRemove As Collection
    Set toRemove = New Collection

    For Each comp In vbProj.VBComponents
        Select Case comp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule
                ' ※ThisWorkbook/Sheet/フォームは削除しない
                toRemove.Add comp
        End Select
    Next comp

    Dim i As Long
    For i = toRemove.Count To 1 Step -1
        vbProj.VBComponents.Remove toRemove(i)
    Next i
End Sub

Private Function NormalizeFolderPath(ByVal p As String) As String
    p = Trim$(p)
    If Len(p) = 0 Then NormalizeFolderPath = p: Exit Function
    If Right$(p, 1) <> "\" Then p = p & "\"
    NormalizeFolderPath = p
End Function
