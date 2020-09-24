Attribute VB_Name = "MxIde_Src_TCprSrc"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Fc."
Type TCprSrc: Mdn As String: Befl As String: Aftl As String: End Type 'Deriving(Ctor Ay)
Function UbTCprSrc&(A() As TCprSrc): UbTCprSrc = SiTCprSrc(A) - 1: End Function
Function SiTCprSrc&(A() As TCprSrc): On Error Resume Next: SiTCprSrc = UBound(A): End Function
Sub PushTCprSrc(O() As TCprSrc, M As TCprSrc): Dim N&: N = SiTCprSrc(O): ReDim Preserve O(N): O(N) = M: End Sub
Sub BrwTCprSrc(S() As TCprSrc): BrwAy FmtTCprSrcyByFilCpr(S), "BrwTCprSrc": End Sub
Function FmtTCprSrcyByFilCpr(S() As TCprSrc) As String()
Dim P$: P = PthTmpInst("TCprSrcy")
WWrtAB S, P
ShellMax FcmdCrt("FileCompare", WCdl(S, P)) ' The Script will write a list of *.Fc.Txt to *P and at end write FcSrcy.end
ChkWaitFfn P & "FcSrcy.end"
FmtTCprSrcyByFilCpr = FmtS12y(S12yoFfnCxtzPth(P, "*.Fc.txt"))
End Function
Private Function WWrtAB(S() As TCprSrc, Pth$) ' Write *.A.Txt & *.B.Txt where * is S().Mdn
Dim J%: For J = 0 To UbTCprSrc(S)
    With S(J)
        Dim S1$: S1 = S(J).Befl
        Dim S2$: S1 = S(J).Aftl
        Dim Ft1$: Ft1 = Pth & .Mdn & ".A.txt"
        Dim Ft2$: Ft2 = Pth & .Mdn & ".B.txt"
    End With
    WrtStr S1, Ft1
    WrtStr S2, Ft2
Next
End Function
Private Function WCdl$(S() As TCprSrc, Pth$)
Dim Fc$
    Dim O$(), J%: For J = 0 To UbTCprSrc(S)
        PushI O, FmtQQ("Fc ?.a.txt ?.b.txt >?.fc.txt", S(J).Mdn)
    Next
    Fc = JnCrLf(O)
WCdl = LinesSap( _
    FmtQQ("Cd ""?""", Pth), _
    Fc, _
    "echo ""End"" >FcSrcy.end")
End Function
