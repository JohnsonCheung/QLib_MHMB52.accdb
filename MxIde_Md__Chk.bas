Attribute VB_Name = "MxIde_Md__Chk"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Md__Chk."
Function ChkMdn(P As VBProject, Mdn) As Boolean
If HasMdnP(P, Mdn) Then Exit Function
MsgBox FmtQQ("Mdn not found: ?|In Pj: ?", Mdn, P.Name), vbCritical
ChkMdn = True
End Function
