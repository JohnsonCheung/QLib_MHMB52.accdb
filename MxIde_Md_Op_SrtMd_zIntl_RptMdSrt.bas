Attribute VB_Name = "MxIde_Md_Op_SrtMd_zIntl_RptMdSrt"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_MdSrc_Rpt."
Private Function WFmt(Src$(), Optional MdnDot$) As String()
Dim X As Dictionary
Dim Y As Dictionary
Set X = DiMthlSrc(Src, MdnDot)
Set Y = DiSrt(X)
WFmt = FmtCprDi(X, Y, "BefSrt AftSrt", IsExlSam:=True)
End Function

Sub RptMdSrtPC(): Brw WFmtP(CPj): End Sub
Sub RptMdSrtMC(): Brw WFmtM(CMd): End Sub

Private Function WFmtP(P As VBProject) As String()
Dim O$(), C As VBComponent
For Each C In P.VBComponents
    PushIAy O, WFmtM(C.CodeModule)
Next
WFmtP = O
End Function

Private Function WFmtM(M As CodeModule) As String(): WFmtM = WFmt(SrcM(M)): End Function
