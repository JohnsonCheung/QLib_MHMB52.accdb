Attribute VB_Name = "MxIde_Dv_Udt_zTool_ToolsDvUdt"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_Udt_CprDvu."
Private Sub B_CprSrcDvUdt()
GoSub T1
Exit Sub
Dim Mdn$
T1:
    Mdn = "MxDao_Db_Lnk"
    GoTo Tst
Tst:
    SrcoptDvUdt SrcMdn(Mdn)
    Return
End Sub
Sub CprSrcDvUdtMdn(Mdn$): CprSrcDvUdtM Md(Mdn): End Sub
Sub CprSrcDvUdtMC():      CprSrcDvUdtM CMd:     End Sub
Private Sub CprSrcDvUdtM(M As CodeModule)
Dim S$(): S = SrcM(M)
CprLyopt S, SrcoptDvUdt(S), "SrcBefDvUdt SrcAftDvUdt ", "Comparing DvUdt "
End Sub

Sub DmpMdnyNeedDvUdt()
Dim O$()
Dim C As VBComponent: For Each C In CPj.VBComponents
    With SrcoptDvUdt(SrcCmp(C))
        If .Som Then PushI O, C.Name
    End With
Next
Dmp O
End Sub

