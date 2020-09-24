Attribute VB_Name = "MxDta_Aet_Brw"
Option Compare Text
Const CMod$ = "MxDta_Aet_Brw."
Option Explicit
Sub VcAet(Aet As Dictionary, Optional PfxFn$ = "VcAet_"):   Vc Aet.Keys, PfxFn:  End Sub
Sub BrwAet(Aet As Dictionary, Optional PfxFn$ = "BrwAet_"): Brw Aet.Keys, PfxFn: End Sub
Sub DmpAet(Aet As Dictionary):                              D Aet.Keys:          End Sub

Function FmtAet(A As Dictionary) As String()
Dim N%: N = NDig(A.Count)
Dim O$()
Dim K: For Each K In A.Keys
    Dim J%: J = J + 1
    PushI O, AliR(J, N) & " " & K
Next
End Function
