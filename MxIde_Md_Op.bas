Attribute VB_Name = "MxIde_Md_Op"
Option Compare Database
Option Explicit

Sub ClrMd(M As CodeModule)
If M.CountOfLines > 0 Then
    M.DeleteLines 1, M.CountOfLines
End If
End Sub

