Attribute VB_Name = "MxIde_Mth_Slm_Ali_AliSlm"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_Slmb_SlmAliRpl."
Private Sub B_AliSlmMdn():            AliSlmzMdn "MxVb_Dta_Itr":         End Sub
Sub AliSlmMC():                       AliSlmM CMd:                       End Sub
Sub AliSlmzMdn(Mdn$):                 AliSlmM Md(Mdn):                   End Sub
Private Sub AliSlmM(M As CodeModule): RplMdSrcopt M, SrcoptSlm(SrcM(M)): End Sub
Sub WrtMsrcSlm():                     WrtMsrcy MsrcySlmPC:               End Sub
