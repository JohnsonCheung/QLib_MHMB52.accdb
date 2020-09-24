Attribute VB_Name = "MxIde_Src_Pth"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Pth."
Function PjfPth$(PthSrc$)
ChkIsPthSrc PthSrc, "SrcPjf"
PjfPth = RmvFst(Fdr(PthPar(PthPar(PthSrc))))
End Function

Function PthSrcPjDist$(PjDist As VBProject)
Dim P$: P = PthP(PjDist)
PthSrcPjDist = PthAddFdrAp(PthUp(P, 1), ".Src", Fdr(P))
End Function
Function PthSrcCmp$(C As VBComponent):      PthSrcCmp = PthSrcP(PjCmp(C)):     End Function
Function PthSrc$(Pjf$):                        PthSrc = PthAss(Pjf) & ".src\": End Function
Function PthSrcPC$():                        PthSrcPC = PthSrcP(CPj):          End Function
Function PthSrcA$(A As Access.Application):   PthSrcA = PthSrcP(PjMainAcs(A)): End Function
Function PthSrcP$(P As VBProject):            PthSrcP = PthSrc(Pjf(P)):        End Function

Sub EnsPthSrcP(P As VBProject): PthEnsAll PthSrcP(P): End Sub
Sub BrwPthSrcPC():              BrwPth PthSrcPC:      End Sub

Function IsPthSrc(Pth) As Boolean
'IsPthSrc = FdrPar(Pth) <> ".src"
Dim F$: F = Fdr(Pth)
If Not HasSsExt(F, ".xlam .accdb") Then Exit Function
IsPthSrc = Fdr(PthPar(Pth)) = ".Src"
End Function
