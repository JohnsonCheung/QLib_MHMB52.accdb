Attribute VB_Name = "MxVb_Str_Bkt_NmBkt"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Bkt_NmBkt."

Private Sub B_RmvNmBkt()
Dim S$, Nm, BktOpn$
GoSub T1
Exit Sub
T1:
    S = "AA1 B(A())x"
    Ept = "AA1 x"
    BktOpn = "("
    GoTo Tst
T2:
    S = "aaa B(lsdfj)1aa"
    Nm = "B"
    Ept = "aaa 1aa"
    BktOpn = "("
    GoTo Tst
Tst:
    Stop
    Act = RmvNmBkt(S, Nm, BktOpn)
    If Act <> Ept Then Stop
    C
    Return
End Sub
Private Function WP12yNmbkt(Nmbktl, BktOpn$) As P12()
Stop '
End Function
Function Nmbktny(Nmbktl, Optional BktOpn$ = vbBktOpn) As String()
Dim P12y() As P12: P12y = P12yNmbktn(Nmbktl, BktOpn)
Nmbktny = TakP12y(Nmbktl, P12y)
End Function
Function Nmbkt$(Nm, Bktv$, Optional BktOpn$ = vbBktOpn): Nmbkt = Nm & BktOpn & Bktv & BktCls(BktOpn): End Function
Function P12yNmbktn(Nmbktl, BktOpn$) As P12()
Stop '
End Function
Function Nmbktv$(S, Nm, Optional BktOpn$ = vbBktOpn, Optional C As eCas)
Dim P As P12: P = P12Nmbktv(S, Nm, BktOpn, C)
Nmbktv = BetP12(S, P)
End Function
Function RmvNmBktv$(S, Nm, Optional BktOpn$ = vbBktOpn, Optional C As eCas): RmvNmBktv = RmvP12(S, P12Nmbktv(S, Nm, BktOpn, C), Inl:=True): End Function
Function RmvNmBkt$(S, Nm, Optional BktOpn$ = vbBktOpn, Optional C As eCas):   RmvNmBkt = RmvP12(S, P12Nmbkt(S, Nm, BktOpn, C), Inl:=True):  End Function
Private Function P12Nmbktv(S, Nm, BktOpn$, C As eCas) As P12
Dim P As P12: P = P12Nmbkt(S, Nm, BktOpn, C)
P.P1 = P.P1 + Len(Nm)
P12Nmbktv = P
End Function

Private Sub B_P12Nmbkt()
Dim Act As P12, Ept As P12, S, Nm, BktOpn$, C As eCas
GoSub T1
Exit Sub
T1:
    BktOpn = "("
    S = "A BB(1) xx"
    Ept.P1 = 2
    Ept.P2 = 7
    GoTo Tst
Tst:
    Act = P12Nmbkt(S, Nm, BktOpn, C)
    If IsEqP12(Act, Ept) Then Stop
    Return
End Sub
Private Function P12Nmbkt(S, Nm, BktOpn$, C As eCas) As P12 ' ret P12 pointing the bkt pair
Dim P1%, P2%
    P1 = InStr(1, S, Nm & BktOpn, C): If P1 = 0 Then Exit Function
    P2 = PosBktCls(S, P1 + Len(Nm), BktOpn)
P12Nmbkt = P12(P1, P2)
End Function
Private Function WP12Bktv(S, Nm, Optional BktOpn$ = vbBktOpn) As P12 ' ret P12 pointing the bkt pair
Dim C1%
    C1 = InStr(S, Nm & BktOpn): If C1 = 0 Then Exit Function
    C1 = C1 + Len(Nm)
Dim C2%
    C2 = PosBktCls(S, C1, BktOpn)
WP12Bktv = P12(C1, C2)
End Function
