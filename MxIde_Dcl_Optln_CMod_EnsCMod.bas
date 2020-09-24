Attribute VB_Name = "MxIde_Dcl_Optln_CMod_EnsCMod"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Optln_CMod_EnsCMod."

Function CModv$(Dcl$()): CModv = RmvSfxDot(Cnststrv(Dcl, "CMod")): End Function

Sub EnsCModPC()
Dim C As VBComponent: For Each C In CPj.VBComponents
    EnsCModM C.CodeModule
Next
End Sub

Sub EnsCModMC():      EnsCModM CMd:        End Sub
Sub EnsCModMdn(Mdn$): EnsCModM MdMdn(Mdn): End Sub
Private Sub EnsCModM(M As CodeModule)
Dim CModln$: CModln = FmtQQ("Const CMod$ = ""?.""", Mdn(M))
EnsCnstln M, CModln
End Sub
