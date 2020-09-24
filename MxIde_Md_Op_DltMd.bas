Attribute VB_Name = "MxIde_Md_Op_DltMd"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Md_Op_DltMd."
Sub DltMdMdn(Mdn$)
Dim C As VBComponent: Set C = CmpMdn(Mdn)
Stop
End Sub

Sub DltMdC()
Dim C As VBComponent: Set C = CCmp: If IsNothing(C) Then Exit Sub
Dim Mdn$: Mdn = C.Name
Dim M$: M = "Input [Yes] to delete:" & vbCrLf & vbCrLf & Mdn & vbCrLf & RepLstsLines(SrclCmp(C))
If InputBox(M, "Delete Md") = "Yes" Then
    CCmp.Collection.Remove CCmp
    Debug.Print Mdn & " Deleted"
End If
End Sub
