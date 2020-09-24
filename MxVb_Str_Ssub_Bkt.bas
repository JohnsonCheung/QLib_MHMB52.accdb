Attribute VB_Name = "MxVb_Str_Ssub_Bkt"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Ssub_Bkt."

Private Sub B_S123Bkt()
Dim S$, BktOpn$, Act As S123, B As S123
S = "aaaa((a),(b))xxx":    BktOpn = "(":          Ept = Sy("aaaa", "(a),(b)", "xxx"): GoSub Tst
Exit Sub
Tst:
    Act = S123Bkt(S, BktOpn)
    C
    Return
End Sub
Function S123Bkt(S, Optional BktOpn$ = vbBktOpn) As S123 ' Ret 3 string as Sy which is (Bef Bet Aft)-Bkt
Dim P As P12: P = WP12Bkt(S, BktOpn): If IsEmpP12(P) Then Exit Function
S123Bkt = S123( _
    Left(S, P.P1 - 1), _
    BetP12(S, P), _
    Mid(S, P.P2 + 1))
End Function
Function BetBktMust$(S, Fun$, Optional BktOpn$ = vbBktOpn): BetBktMust = BetP12(S, WP12BktMust(S, Fun, BktOpn)): End Function
Private Function WP12BktMust(S, Fun$, Optional BktOpn$ = vbBktOpn) As P12
Const CSub$ = CMod & "WP12BktMust"
Dim PosO%: PosO = InStr(S, BktOpn): If PosO = 0 Then Thw CSub, "No SBktOpn in @S", "@S @BktOpn", S, BktOpn
Dim PosC%: PosC = PosBktCls(S, PosO, BktOpn): If PosC = 0 Then Thw CSub, "No BktCls in @S", "@S @BktOpn", S, BktOpn
WP12BktMust = P12(PosO, PosC)
End Function

Function BetBkt$(S, Optional BktOpn$ = vbBktOpn): BetBkt = BetP12(S, WP12Bkt(S, BktOpn)): End Function

Function AftBkt$(S, Optional BktOpn$ = vbBktOpn)
Dim P As P12: P = WP12Bkt(S, BktOpn): If IsEmpP12(P) Then Exit Function
AftBkt = Mid(S, P.P2 + 1)
End Function

Function BefBkt$(S, Optional BktOpn$ = vbBktOpn): BefBkt = Left(S, WP12Bkt(S, BktOpn).P1 - 1): End Function

Private Sub B_WP12Bkt()
Dim A$, Act As P12, Ept As P12
'
A = "(A(B)A)A"
Ept = P12(1, 7)
GoSub Tst
'
A = " (A(B)A)A"
Ept = P12(2, 8)
GoSub Tst
'
A = " (A(B)A )A"
Ept = P12(2, 9)
GoSub Tst
'
Exit Sub
Tst:
    Act = WP12Bkt(A)
    Debug.Assert IsEqP12(Act, Ept)
    Return
End Sub
Private Function WP12Bkt(S, Optional BktOpn$ = vbBktOpn) As P12
Dim PosO%: PosO = InStr(S, BktOpn): If PosO = 0 Then Exit Function
Dim PosC%: PosC = PosBktCls(S, PosO, BktOpn)
WP12Bkt = P12(PosO, PosC)
End Function
Private Sub B_RmvBetBktAll()
GoSub T1
Exit Sub
Dim S, BktOpn$
T1:
    S = "$(1) AA(2)"
    BktOpn = "("
    Ept = "$() AA()"
    GoTo Tst
Tst:
    Act = RmvBetBktAll(S, BktOpn)
    C
    Return
End Sub
Function P12Ssub(L, Ssub, Optional C As eCas = eCasSen) As P12
Dim P%: P = PosSsub(L, Ssub, C): If P = 0 Then Exit Function
P12Ssub = P12(P, P + Len(Ssub) - 1)
End Function
Function P12Ssuby(L, Ssuby$(), Optional C As eCas) As P12
Dim Ssub: For Each Ssub In Ssuby
    Dim P As P12: P = P12Ssub(L, Ssub, C)
    If P.P1 > 0 Then P12Ssuby = P: Exit Function
Next
End Function

Function P12yBkt(S, Optional BktOpn$) As P12()
Dim PosBeg%: PosBeg = 1
Dim J%
Again:
    ThwLoopTooMuch CSub, J, 200
    Dim P1%: P1 = InStr(PosBeg, S, BktOpn): If P1 = 0 Then Exit Function
    Dim P2%: P2 = PosBktCls(S, P1, BktOpn)
    PushP12 P12yBkt, P12(P1, P2)
    PosBeg = P2 + 1
    GoTo Again
End Function
Function P12Bkt(S, Optional BktOpn$ = vbBktOpn, Optional PosBeg% = 1) As P12
Dim P%: P = InStr(PosBeg, S, BktOpn): If P = 0 Then Exit Function
P12Bkt = P12(P, PosBktCls(S, P, BktOpn))
End Function
Function RmvBetBktAll$(S, Optional BktOpn$ = vbBktOpn): RmvBetBktAll = RmvP12y(S, P12yBkt(S, BktOpn)): End Function
Function RmvBetBkt$(S, Optional BktOpn$ = vbBktOpn):       RmvBetBkt = RmvP12(S, P12Bkt(S, BktOpn)):   End Function
Private Sub B_RmvBetBkt()
GoSub T1
Exit Sub
Dim S, BktOpn$
T1:
    S = "$(1)"
    BktOpn = "("
    Ept = "$()"
    GoTo Tst
Tst:
    Act = RmvBetBkt(S, BktOpn)
    C
    Return
End Sub
