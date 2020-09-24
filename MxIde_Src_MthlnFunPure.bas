Attribute VB_Name = "MxIde_Src_MthlnFunPure"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_MthlnFunPure."
Dim Inf$()
Private Sub B_IsLnPurFun()
Dim O$()
Dim L: For Each L In SrcPC
    If IsLnPurFun(L) Then PushI O, L
Next
Vc O
End Sub
Function IsLnPurFun(L) As Boolean
If Not IsLnFun(L) Then Exit Function
If L = "Me.RecordSource = FmtStr(""SELECT CdCo,x.*, NmHse"" & _" Then Stop

IsLnPurFun = BetBkt(L) = ""
End Function
