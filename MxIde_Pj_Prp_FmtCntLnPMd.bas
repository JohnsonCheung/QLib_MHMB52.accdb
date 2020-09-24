Attribute VB_Name = "MxIde_Pj_Prp_FmtCntLnPMd"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Pj_Len."
Function FmtCntWrdPC$():              FmtCntWrdPC = FmtCntWrdP(CPj):     End Function
Function FmtCntWrdP$(P As VBProject):  FmtCntWrdP = FmtCntWrd(SrclP(P)): End Function
Function FmtCntLnPC$():                FmtCntLnPC = FmtCntLnP(CPj):      End Function
Function FmtCntLnP$(P As VBProject):    FmtCntLnP = FmtCntLn(SrclP(P)):  End Function
Function FmtCntLnM$(M As CodeModule):   FmtCntLnM = FmtCntLn(SrclM(M)):  End Function
Function LenPjf&(P As VBProject):          LenPjf = FileLen(Pjf(P)):     End Function
Function LenPjfC&():                      LenPjfC = LenPjf(CPj):         End Function
