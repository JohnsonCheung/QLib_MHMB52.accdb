Attribute VB_Name = "MxDta_Da_Fmt_InVertFmt"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_Fmt_InVertFmt."

Function FmtDrsVert(D As Drs, Optional Nm$) As String()
If NoRecDrs(D) Then PushI FmtDrsVert, NoRecMsg(D, Nm): Exit Function
Dim FldLbl$(): FldLbl = AmStrSfx(AmAli(AmAddIxPfx(D.Fny)), ": ")
Dim N&: N = NRecDrs(D)
Dim O$()
    PushIAy O, Box(Nm)
    Dim J&: For J = 0 To UB(D.Dy)
        Dim Dr(): Dr = D.Dy(J)
        PushIAy O, WFmtDrInVertFmt(Dr, FldLbl, J, N)
    Next
FmtDrsVert = O
End Function
Private Function WFmtDrInVertFmt(Dr, FldLbl$(), Ix&, NRec&) As String()
Dim O$()
    PushI O, "Rec#" & Ix & " of " & NRec
    Dim J%: For J = 0 To UB(FldLbl)
        PushI O, FldLbl(J) & Dr(J)
    Next
WFmtDrInVertFmt = O
End Function
