Attribute VB_Name = "MxIde_Cmp_Op_Cpy"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Cmp_Op_Cpy."
Sub CpyCmpToPj(C As VBComponent, PjTo As VBProject, Optional PfxNw$)
Const CSub$ = CMod & "CpyCmpToPj"
Dim NmNw$: NmNw = PfxNw & C.Name
If HasMdnP(PjTo, NmNw) Then Thw CSub, "CmpnTo exist in PjTo", "CmpnTo PjTo", NmNw, Pjn(PjTo)
AddCmpCdl PjTo, NmNw, C.Type, SrclCmp(C)
End Sub

Sub CpyCmpToMd(C As VBComponent, MdTo As CodeModule)
RplMd MdTo, SrclCmp(C)
End Sub
