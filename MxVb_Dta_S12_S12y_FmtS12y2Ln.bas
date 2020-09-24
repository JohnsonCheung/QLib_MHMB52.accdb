Attribute VB_Name = "MxVb_Dta_S12_S12y_FmtS12y2Ln"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_S12_S12y_Fmt_As2Ln."

Private Sub B_FmtS12y2Ln()
End Sub
Function FmtS12y2Ln(A() As S12, Optional W% = 100) As String()
Dim Ix&
Dim WdtIx As Byte: WdtIx = NDig(SiS12(A))
Dim PfxSndLn$: PfxSndLn = "    "
Dim J%: For J = 0 To UbS12(A)
    Ix = Ix + 1
    With A(J)
        PushIAy FmtS12y2Ln, WLyS1(.S1, WdtIx, W, Ix)
        PushIAy FmtS12y2Ln, WLyS2(.S2, WdtIx, W, PfxSndLn)
    End With
Next
End Function
Private Function WLyS1(S1$, WdtIx As Byte, W%, Ix&) As String()
Stop
End Function
Private Function WLyS2(S2$, WdtIx As Byte, W%, PfxSndLn) As String()
Stop
'WLyS2 = LyWrpWrdLines(S1, LnWdt, SndLnIndt)
End Function
