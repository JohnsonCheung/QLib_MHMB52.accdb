Attribute VB_Name = "MxIde_Src_LnOptEns"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_LnOptEns."
Const LnOptExp$ = "Option Explicit"
Const LnOptCprBin$ = "Option Compare Binary"
Const LnOptCprDb$ = "Option Compare Database"
Const LnOptCprTxt$ = "Option Compare Text"

Sub Ens3OptMC():       Ens3OptM CMd:        End Sub
Sub Ens3OptPC():       Ens3OptP CPj:        End Sub
Sub Ens3OptzMdn(Mdn$): Ens3OptM MdMdn(Mdn): End Sub
Private Sub Ens3OptP(P As VBProject)
Dim C As VBComponent: For Each C In P.VBComponents
    Ens3OptM C.CodeModule
Next
End Sub

Private Sub B_Ens3OptM()
Dim M As CodeModule
Const Mdn$ = "AA"
GoSub Setup
GoSub T0
GoSub Clean
Exit Sub
T0:
    Set M = Md(Mdn)
    GoTo Tst
Tst:
    Ens3OptM M
    Return
Setup:
    AddCls Mdn
    Return
Clean:
    RmvMdn Mdn
    Return
End Sub
Private Sub Ens3OptM(M As CodeModule)
If IsMdEmp(M) Then Exit Sub
'Debug.Print CSub; " Md: "; Mdn(M)
WDltLnOptM M, LnOptCprDb
If Mdn(M) = "MxDta_Da_Op_Srt_RxySrtDy" Then
    WDltLnOptM M, LnOptCprTxt
    WEnsLnOptM M, LnOptCprBin
Else
    WDltLnOptM M, LnOptCprBin
    WEnsLnOptM M, LnOptCprTxt
End If
WEnsLnOptM M, LnOptExp
End Sub
Private Sub WDltLnOptM(M As CodeModule, LnOpt)
Const CSub$ = CMod & "WDltLnOptM"
Dim I%: I = LnoPfxDclln(M, LnOpt)
If I = 0 Then Exit Sub
M.DeleteLines I
Infln CSub, "[" & LnOpt & "] line is deleted", "Md Lno", Mdn(M), I
End Sub
Private Sub WEnsLnOptM(M As CodeModule, LnOpt)
Const CSub$ = CMod & "WEnsLnOptM"
If M.CountOfLines = 0 Then Exit Sub
If LnoPfxDclln(M, LnOpt) > 0 Then Exit Sub
M.InsertLines 1, LnOpt
Infln CSub, "[" & LnOpt & "] is Inserted", "Md", Mdn(M)
End Sub
