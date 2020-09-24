Attribute VB_Name = "MxVb_Dta_Lyy"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Lyy."
Private Sub B_VcLyy(): VcLyy WSrcLyyP: End Sub
Private Function WSrcLyyP() As Variant()
Dim C As VBComponent: For Each C In CPj.VBComponents
    PushI WSrcLyyP, SrcM(C.CodeModule)
Next
End Function
Sub VcLyy(Lyy()):  VcAy FmtLyy(Lyy):  End Sub
Sub BrwLyy(Lyy()): BrwAy FmtLyy(Lyy): End Sub
Sub DmpLyy(Lyy()): DmpAy FmtLyy(Lyy): End Sub

Function FmtLyy(Lyy()) As String(): FmtLyy = FmtLsy(LsyLyy(Lyy)): End Function
Function AmQuoAli(Ay, W%, Optional QuoStr$, Optional IsAliR As Boolean) As String()
Dim Q As S12: Q = BrkQuo(QuoStr)
Dim I: For Each I In Itr(Ay)
    PushI AmQuoAli, Q.S1 & Ali(I, W, IsAliR) & Q.S2
Next
End Function
Function WdtLyy%(Lyy())
Dim O%
Dim Sy: For Each Sy In Itr(Lyy)
    O = Max(O, AyWdt(Sy))
Next
WdtLyy = O
End Function

Function LyLyy(Lyy()) As String()
Dim Ly: For Each Ly In Itr(Lyy)
    PushIAy LyLyy, Ly
Next
End Function
Function LinesLyy$(Lyy()): LinesLyy = JnCrLf(LyLyy(Lyy)): End Function
Function LstssLyy(Lyy()) As Lstss
With LstssLyy
    .Len = LenLyy(Lyy)
    .N = Si(Lyy)
    .NLn = NLnLyy(Lyy)
End With
End Function
Function NLnLyy&(Lyy())
Dim J&: For J = 0 To Si(Lyy)
    NLnLyy = NLnLyy + Si(Lyy(J))
Next
End Function
Function LenLyy&(Lyy())
Dim J&: For J = 0 To Si(Lyy)
    LenLyy = LenLyy + Len(Lyy(J))
Next
End Function
