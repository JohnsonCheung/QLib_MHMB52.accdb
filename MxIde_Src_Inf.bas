Attribute VB_Name = "MxIde_Src_Inf"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Inf."
Function SrcMdn(Mdn) As String():  SrcMdn = SrcM(Md(Mdn)):  End Function
Function SrclMdn$(Mdn):           SrclMdn = SrclM(Md(Mdn)): End Function

Sub AsgMthDr(MthDr, OMdy$, OTy$, ONm$, OPrm$, ORet$, OLinRmk$, OLines$, OMmk$)
AsgAy MthDr, OMdy, OTy, ONm, OPrm, ORet, OLinRmk, OLines, OMmk
End Sub

Private Sub B_VbExmLy(): Vc VbExmLy(SrcP(CPj)): End Sub
Function VbExmLy(Src$()) As String()
Dim L: For Each L In Itr(Src)
    If IsLnVmkExm(L) Then PushI VbExmLy, L
Next
End Function

Function IsLnVmkExm(Ln) As Boolean ' #Vb-Exclaimation-Rmk-Line# ! It is a rmk Ln fst-non-spc-chr is ['] and nxt is [!]
Dim L$: L = LTrim(Ln)
If Not IsShfPfx(L, "'") Then Exit Function
L = LTrim(L)
If ChrFst(L) <> "!" Then Exit Function
IsLnVmkExm = True
End Function
Function SrcLcnt(M As CodeModule, A As Lcnt) As String():  SrcLcnt = SplitCrLf(SrclLcnt(M, A)): End Function
Function SrclLcnt$(M As CodeModule, A As Lcnt):           SrclLcnt = M.Lines(A.Lno, A.Cnt):     End Function
Function HasQuoSngDbl(S) As Boolean
If HasQuoSng(S) Then
    If HasQuoDbl(S) Then
        HasQuoSngDbl = True
    End If
End If
End Function

Private Sub B_SrcHasQuoSngDbl()
VcAy SrcHasQuoSngDbl(SrcPC)
End Sub
Function SrcHasQuoSngDbl(Src$()) As String()
Dim L: For Each L In Itr(Src)
    If HasQuoSngDbl(L) Then
        PushI SrcHasQuoSngDbl, L
    End If
Next
End Function


Function HasQuoSngDblSrc(Src$()) As String()
Dim L: For Each L In Itr(Src)
    If HasQuoSngDbl(L) Then PushI HasQuoSngDblSrc, L
Next
End Function

Function NLnSrcPC&(): NLnSrcPC = NLnSrcP(CPj): End Function
Function NLnSrcP&(P As VBProject)
If P.Protection = vbext_pp_locked Then Exit Function
Dim C As VBComponent: For Each C In P.VBComponents
    NLnSrcP = NLnSrcP + C.CodeModule.CountOfLines
Next
End Function
