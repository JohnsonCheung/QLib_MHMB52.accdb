Attribute VB_Name = "MxAcs_Acs_SrcTxt_BrwSrcTxt"
Option Compare Text
Option Explicit
Const CMod$ = "MxAcs_SrcTxt_BrwSrcTxt."
Sub BrwTxtFrm(Frmn$): BrwFt WFtFrm(Frmn): End Sub
Sub BrwTxtRpt(Rptn$): BrwFt WFtRpt(Rptn): End Sub
Sub VcTxtFrm(Frmn$):  VcFt WFtFrm(Frmn):  End Sub
Sub VcTxtRpt(Rptn$):  VcFt WFtRpt(Rptn):  End Sub
Private Function WFtFrm$(Frmn$)
WFtFrm = FtTmp("TxtFrm", Frmn & "_")
Acs.SaveAsText acForm, Frmn, WFtFrm
End Function
Private Function WFtRpt$(Rptn$)
WFtRpt = FtTmp("TxtRpt", Rptn & "_")
Acs.SaveAsText acReport, Rptn, WFtRpt
End Function
