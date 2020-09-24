Attribute VB_Name = "MxIde_Pj_f"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Pj_f."
Public XlsPjf As New Excel.Application

Function TmpFxa$(Optional Fdr$, Optional Fnn$): TmpFxa = FfnTmp(".xlam", Fdr, Fnn): End Function
