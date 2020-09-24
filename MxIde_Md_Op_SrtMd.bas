Attribute VB_Name = "MxIde_Md_Op_SrtMd"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Md_MdOp_SrtMd."
Sub SrtMdMC():                      SrtMd CMd:                      End Sub
Private Sub SrtMd(M As CodeModule): RplMdSrclopt M, SrcloptSrtM(M): End Sub
Sub SrtMdPC()
BkuPjC
Dim C As VBComponent
For Each C In CPj.VBComponents
    SrtMd C.CodeModule
Next
End Sub
