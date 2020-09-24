Attribute VB_Name = "MxIde_Pj_HasMd"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Pj_HasMd."
Function HasMdnP(P As VBProject, Mdn) As Boolean: HasMdnP = HasItn(P.VBComponents, Mdn): End Function
Function HasMdn(Mdn) As Boolean:                   HasMdn = HasMdnP(CPj, Mdn):           End Function

Sub ChkMdnExist(P As VBProject, Mdn, Fun$)
If Not HasMdnP(P, Mdn) Then Thw Fun, "Should be a Mod", "Mdn CmpTy", Mdn, ShtCmpTy(Cmp(Mdn).Type)
End Sub

Function HasMod(P As VBProject, Modn) As Boolean
If Not HasMdnP(P, Modn) Then Exit Function
HasMod = P.VBComponents(Modn).Type = vbext_ct_StdModule
End Function
