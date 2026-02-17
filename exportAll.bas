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
        If InStr(comp.name, "Sheet") > 0 Then GoTo Continue
        If InStr(comp.name, "ThisWorkbook") > 0 Then GoTo Continue
        
        Select Case comp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm, vbext_ct_Document
                comp.Export outDir & comp.name & ModuleExt(comp.Type)
        End Select
Continue:
    Next comp
    
    MsgBox "Exported to: " & outDir
End Sub

Private Function ModuleExt(t As VBIDE.vbext_ComponentType) As String
    Select Case t
        Case vbext_ct_StdModule:  ModuleExt = ".bas"
        Case vbext_ct_ClassModule: ModuleExt = ".cls"
        Case vbext_ct_MSForm:     ModuleExt = ".frm" ' .frxÅE
        Case vbext_ct_Document:   ModuleExt = ".cls" ' ThisWorkbook/Sheet
        Case Else:                ModuleExt = ".txt"
    End Select
End Function

