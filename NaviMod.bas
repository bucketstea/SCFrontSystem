Attribute VB_Name = "NaviMod"
Option Explicit

Private mStack As Collection
Private mCurrent As Object ' UserForm

Public Sub initNav()
    Set mStack = New Collection
    Set mCurrent = Nothing
End Sub

Public Sub showStart(frm As Object)
    If mStack Is Nothing Then initNav
    Set mCurrent = frm
    frm.Show vbModeless
End Sub

Public Sub navigateTo(nextFrm As Object)
    If mStack Is Nothing Then initNav

    If Not mCurrent Is Nothing Then
        mStack.Add mCurrent
        mCurrent.Hide
    End If

    Set mCurrent = nextFrm
    mCurrent.Show vbModeless
End Sub

Public Sub goBack()
    If mStack Is Nothing Then Exit Sub
    If mStack.Count = 0 Then Exit Sub

    ' 現在画面を閉じる/隠す（要件で選ぶ）
    If Not mCurrent Is Nothing Then
        mCurrent.Hide   ' or Unload mCurrent
    End If

    Set mCurrent = mStack(mStack.Count)
    mStack.Remove mStack.Count

    mCurrent.Show vbModeless
End Sub
