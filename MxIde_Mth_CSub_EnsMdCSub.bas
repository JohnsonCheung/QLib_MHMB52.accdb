Attribute VB_Name = "MxIde_Mth_CSub_EnsMdCSub"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_EnsCSub."
Private Sub B_EnsMdCSubM()
GoSub T1
Exit Sub
Dim M As CodeModule
T1:  Set M = CMd: GoTo Tst
Tst: EnsMdCSubM M: Return
End Sub
Sub WrtMsrcCSub():                       WrtMsrcy MsrcyCSubPC:            End Sub
Sub EnsMdCSubMC():                       EnsMdCSubM CMd:                  End Sub
Sub EnsMdCSubMdn(Mdn$):                  EnsMdCSubM Md(Mdn):              End Sub
Private Sub EnsMdCSubM(M As CodeModule): RplMdSrcopt M, SrcoptEnsCSub(M): End Sub
