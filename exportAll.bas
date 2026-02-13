Attribute VB_Name = "exportAll"
Option Explicit

Public Sub ExportAllModules()
    Dim vbProj As VBIDE.VBProject
    Dim comp As VBIDE.VBComponent
    Dim outDir As String
    
    outDir = ThisWorkbook.Path & "\_vba_export\"
    If Dir(outDir, vbDirectory) = "" Then MkDir outDir
    
    Set vbProj = ThisWorkbook.VBProject
    
    For Each comp In vbProj.VBComponents
        Select Case comp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm, vbext_ct_Document
                comp.Export outDir & comp.name & ModuleExt(comp.Type)
        End Select
    Next
    
    MsgBox "Exported to: " & outDir
End Sub

Private Function ModuleExt(t As VBIDE.vbext_ComponentType) As String
    Select Case t
        Case vbext_ct_StdModule:  ModuleExt = ".bas"
        Case vbext_ct_ClassModule: ModuleExt = ".cls"
        Case vbext_ct_MSForm:     ModuleExt = ".frm" ' .frxも一緒に出ます
        Case vbext_ct_Document:   ModuleExt = ".cls" ' ThisWorkbook/Sheet等
        Case Else:                ModuleExt = ".txt"
    End Select
End Function
