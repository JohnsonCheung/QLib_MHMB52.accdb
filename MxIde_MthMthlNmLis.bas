Attribute VB_Name = "MxIde_MthMthlNmLis"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_MthMthlNmLis."

Sub DmpMthlsyNm(Mthn, Optional ShtMthTy$)
DmpAy WFmt(Mthn, ShtMthTy)
End Sub
Private Function WFmt(Mthn, ShtMthTy$) As String()
WFmt = FmtLsy(MthlsyNmPC(Mthn, ShtMthTy))
End Function
