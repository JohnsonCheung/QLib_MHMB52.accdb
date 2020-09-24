Attribute VB_Name = "MxVb_Str_Cpr"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Cpr."
Private Sub B_FmtCprLines()
Dim A$, B$
A = LinesVbl("AAAAAAA|bbbbbbbb|cc|dd")
B = LinesVbl("AAAAAAA|bbbbbbbb |cc")
GoSub Tst
Exit Sub
Tst:
    Act = FmtCprLines(A, B)
    Brw Act
    Return

End Sub
Private Sub B_CprLines()
GoSub T1
Exit Sub
Dim A$, B$, P12$, Hdr$
T1:
    A = LinesVbl("AAAAAAA|bbbbbbbb|cc|dd")
    B = LinesVbl("AAAAAAA|bbbbbbbb |cc")
    P12 = "Lines-1 Lines-2"
    Hdr = "XXXX"
    GoTo Tst
Tst:
    CprLines A, B, P12, Hdr
    Return
End Sub
Sub ChkIsEqLines(A$, B$, Optional P12$ = "Lines1 Lines2", Optional Hdr$)
If A = B Then Exit Sub
CprLines A, B, P12, Hdr
End Sub
Sub CprLinesS12(S As S12, Optional P12$ = "Lines1 Lines2", Optional Hdr$): CprLines S.S1, S.S2, P12, Hdr:           End Sub
Sub CprLy(A$(), B$(), Optional P12$ = "Lines1 Lines2", Optional Hdr$):     CprLines JnCrLf(A), JnCrLf(B), P12, Hdr: End Sub
Sub CprLyopt(A$(), B As Lyopt, Optional P12$ = "Lines1 Lines2", Optional Hdr$)
If B.Som Then MsgBox P12 & vbCrLf & "Lyopt.Som = False", vbInformation, "CprLyopt": Exit Sub
CprLy A, B.Ly, P12, Hdr
End Sub
Sub CprLines(A$, B$, Optional N12$ = "Lines1 Lines2", Optional Hdr$)
Const Fdr$ = "CprLines"
If A = B Then
    Dim M$:
        M = N12 & " are equals" & vbCrLf & _
            LstsLines(A)
    MsgBox M, vbInformation, StrDft(Hdr, Fdr)
    Exit Sub
End If
Dim FtA$, FtB$
    FtA = FtTmp(Fdr, Tm1(N12) & "-"): WrtStr A, FtA
    FtB = FtTmp(Fdr, Tm2(N12) & "-"): WrtStr B, FtB
ShellHid FmtQQ("""?"" --diff ""?"" ""?"" --add ""?""", vc_cmdFfn, FtA, FtB, PthTmpFdr(Fdr))
End Sub

Function FmtCprLines(A$, B$, Optional P12$ = "A B", Optional Hdr$) As String()
If A = B Then PushI FmtCprLines, "Two lines are equal.  P12=[" & P12 & "]"
Dim AA$(): AA = SplitCrLf(A)
Dim BB$(): BB = SplitCrLf(B)
Dim N1$, N2$: AsgT1r P12, N1, N2
Dim NIxDig%: NIxDig = Len(Max(Si(AA), Si(BB)))
Dim H$(): H = WhDr(AA, BB, N1, N2, Hdr)
Dim L$(): L = WCpr(AA, BB, N1, N2, NIxDig)
Dim R$(): R = WRst(AA, BB, N1, N2)
FmtCprLines = SyAddAp(H, L, R)
End Function
Private Function WhDr(A$(), B$(), N1$, N2$, Hdr$) As String()  ' The [Hdr] part
Dim O$()
PushI O, FmtQQ("LinesCnt=? (?)", Si(A), N1)
PushI O, FmtQQ("LinesCnt=? (?)", Si(B), N2)
WhDr = O
End Function
Private Function WCpr(A$(), B$(), N1$, N2$, NIxDig%) As String() ' The [Cpr] part
Dim J&: For J = 0 To Min(UB(A), UB(B))
    PushIAy WCpr, WOneLn(A(J), B(J), J, NIxDig)
Next
End Function
Private Function WOneLn(A$, B$, Ix&, NIxDig%) As String() ' The [Lin] which will be 1 line if same or 2 lines if dif
PushI WOneLn, WFmtLn(Ix, NIxDig, A)
If A = B Then Exit Function
PushI WOneLn, WFmtLnSpc(NIxDig, B)
End Function
Private Function WRst(A$(), B$(), N1$, N2$) As String() ' The [Rst] part
Dim NA&, NB&, N&: NA = Si(A): NB = Si(B): N = Max(NA, NB)
If NA = NB Then Exit Function
Dim MoreNm$, LessNm$, NLn&
    If NA > NB Then
        MoreNm = N1
        LessNm = N2
    Else
        MoreNm = N2
        LessNm = N1
    End If
    NLn = Abs(NA - NB)

Dim O$()
    PushI O, FmtQQ("-- ? has more ? lines then ? -------", MoreNm, NLn, LessNm) '<===
    Dim Large$()
        If NA > NB Then
            Large = A
        Else
            Large = B
        End If
    Dim J&
    Dim NIxDig%: NIxDig = Len(N)
    For J = Min(NA, NB) To N - 1
        PushI WRst, WFmtLn(J, NIxDig, Large(J))
    Next
WRst = O
End Function
Private Function WFmtLn$(Ix&, NIxDig%, Ln$):    WFmtLn = FmtQQ("? ?", AliR(Ix + 1, NIxDig), Ln): End Function
Private Function WFmtLnSpc$(NIxDig%, Ln):    WFmtLnSpc = Space(NIxDig + 1) & Ln:                 End Function

Sub ChkIsEqDi(A As Dictionary, B As Dictionary, Optional P12$, Optional Hdr$)
Const CSub$ = CMod & "ChkIsEqDi"
If IsEqDi(A, B) Then Exit Sub
Thw CSub, "Two given di are diff"
End Sub
Sub ChkIsEqStr(A$, B$, Optional P12$ = "A B", Optional Hdr$)
If A = B Then Exit Sub
Brw FmtCprLines(A, B, P12, Hdr)
End Sub

Function FmtCprStr(A$, B$, Optional N12$ = "A B", Optional Hdr$) As String()
Dim N1$, N2$
AsgT1r N12, N1, N2
If A = B Then
    PushI FmtCprStr, FmtQQ("Str(?) = Str(?).  Len(?)", N1, N2, Len(A))
    Exit Function
End If
Select Case True
Case IsLines(A), IsLines(B): FmtCprStr = FmtCprLines(A, B, N12)
Case Else: FmtCprStr = W2FmtCprStr(A, B, N1, N2, Hdr)
End Select
End Function
Private Function W2FmtCprStr(A$, B$, N1$, N2$, Hdr$) As String() '
Dim P&: P = PosAtDif(A, B)
    Dim L1&, L2&: L1 = Len(A): L2 = Len(B)
Dim O$()
    PushIAy O, Box(Hdr)
    PushI O, FmtQQ("Str1 (Len / Nm): ? / ?", AliR(L1, 6), N1)
    PushI O, FmtQQ("Left2 (Len / Nm): ? / ?", AliR(L2, 6), N2)
    PushI O, FmtQQ("Dif at position: ?", P)
    PushI O, LinesLbl(Max(L1, L2))
    PushI O, A
    PushI O, B
    PushI O, Space(P - 1) & "^"
W2FmtCprStr = O
End Function
