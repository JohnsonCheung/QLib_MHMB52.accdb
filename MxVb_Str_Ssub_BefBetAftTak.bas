Attribute VB_Name = "MxVb_Str_Ssub_BefBetAftTak"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Ssub_BefBetAftTak."

Function BefDotRev$(S):                                BefDotRev = BefRev(S, "."):             End Function
Function BefDot$(S):                                      BefDot = Bef(S, "."):                End Function
Function BefDotOrAll$(S):                            BefDotOrAll = BefOrAll(S, "."):           End Function
Function BefCma(S, Optional NoTrim As Boolean):           BefCma = Aft(S, vbCma, NoTrim):      End Function
Function BefCmaOrAll(S, Optional NoTrim As Boolean): BefCmaOrAll = BefOrAll(S, vbCma, NoTrim): End Function

Function AftCmaOrAll(S, Optional NoTrim As Boolean):     AftCmaOrAll = AftOrAll(S, vbCma, NoTrim): End Function
Function AftCma(S, Optional NoTrim As Boolean):               AftCma = Aft(S, vbCma, NoTrim):      End Function
Function AftSngQ$(S, Optional NoTrim As Boolean):            AftSngQ = Aft(S, vbQuoSng, NoTrim):   End Function
Function Aft$(S, Ssub$, Optional NoTrim As Boolean):             Aft = Brk1(S, Ssub, NoTrim).S2:   End Function
Function AftMust$(S, Ssub$, Optional NoTrim As Boolean):     AftMust = Brk(S, Ssub, NoTrim).S2:    End Function
Function AftColonOrAll$(S):                              AftColonOrAll = AftOrAll(S, ":", NoTrim:=True): End Function

Function BefColonOrAll$(S): BefColonOrAll = BefOrAll(S, ":", NoTrim:=True): End Function
Function AftAt$(S, At&, Ssub$):
If At = 0 Then Exit Function
AftAt = Mid(S, At + Len(Ssub))
End Function

Function AftDotOrAll$(S):       AftDotOrAll = AftOrAll(S, "."):    End Function
Function AftDotOrAllRev$(S): AftDotOrAllRev = AftOrAllRev(S, "."): End Function
Function AftDot$(S):                 AftDot = Aft(S, "."):         End Function

Function AftOrAll$(S, Ssub$, Optional NoTrim As Boolean):    AftOrAll = Brk2(S, Ssub, NoTrim).S2:    End Function
Function AftOrAllRev$(S, Ssub$):                          AftOrAllRev = StrDft(AftRev(S, Ssub), S):  End Function
Function AftRev$(S, Ssub$, Optional NoTrim As Boolean):        AftRev = Brk1Rev(S, Ssub, NoTrim).S2: End Function
Function BefRev$(S, Ssub$, Optional NoTrim As Boolean):        BefRev = Brk1Rev(S, Ssub, NoTrim).S1: End Function
Function BefSpc$(S):                                           BefSpc = Bef(S, " "):                 End Function
Function AftSpc$(S, Optional NoTrim As Boolean):               AftSpc = Aft(S, " ", NoTrim):         End Function
Function BefSpcOrAll$(S):                                 BefSpcOrAll = BefOrAll(S, " "):            End Function
Function BefSyAny(Sy$(), Ssub$, Optional NoTrim As Boolean) As String()
Dim I: For Each I In Itr(Sy)
    PushI BefSyAny, Bef(I, Ssub, NoTrim)
Next
End Function
Function BetP1P2$(S, P1, P2, Optional Inl As Boolean): BetP1P2 = BetP12(S, P12(P1, P2), Inl): End Function
Function BetSPr$(S, S1$, S2$, Optional PosBeg = 1, Optional Inl As Boolean, Optional C As eCas)
Dim M$
    If S1 = "" Then M = "S1 cannot be blank": GoTo M
    If S2 = "" Then M = "S2 cannot be blank": GoTo M
    If PosBeg <= 0 Then M = "PosBeg cannot be <= 0": GoTo M
    If S = "" Then Exit Function
Dim WP1&: WP1 = PosSsub(S, S1, C, PosBeg): If WP1 = 0 Then Exit Function
Dim WN1&: WN1 = WP1 + Len(S1)
Dim WP2&: WP2 = PosSsub(S, S2, C, WN1): If WP2 = 0 Then Exit Function
Dim WN2&: WN2 = WP2 + Len(S2)
Dim P%, L%: Stop
BetSPr = Mid(S, P, L)
Exit Function
M: ThwPm CSub, M, "S S1 S2 PosBeg Inl Cas", S, S1, S2, PosBeg, Inl, EnmsCas(C)
End Function
Function BetS1Opt2$(S, S1$, OptS2$, Optional BegPos = 1, Optional Inl As Boolean, Optional C As eCas)
Dim M$
    If S1 = "" Then M = "S1 cannot be blank": GoTo M
    If OptS2 = "" Then M = "OptS2 cannot be blank": GoTo M
    If S = "" Then Exit Function
    If BegPos <= 0 Then M = "BegPos cannot be <= 0": GoTo M
Dim WP1&: WP1 = PosSsub(S, S1, C, BegPos): If WP1 = 0 Then Exit Function
Dim WN1&: WN1 = WP1 + Len(S1)
Dim WP2&
Dim WN2&
    Dim P2&: P2 = PosSsub(S, OptS2, C, WN1)
    If P2 = 0 Then
        WP2 = Len(S) + 1
        WN2 = WP2
    Else
        WP2 = PosSsub(S, OptS2, WN1, C)
        WN2 = WP2 + Len(OptS2)
    End If
Dim P%, L%: Stop
BetS1Opt2 = Mid(S, P, L)
Exit Function
M: ThwPm CSub, M, "S S1 Opt2 BegPos Inl Cas", S, S1, OptS2, BegPos, Inl, EnmsCas(C)
End Function
Private Sub BetP12__Tst()
GoSub T1
Exit Sub
Dim S$, P As P12, Inl As Boolean
'        123456789
T1: S = "df(as bf)sdf"
    P = P12(3, 9)
    Inl = False
    Ept = "as bf"
    GoTo Tst
T2: S = "df(as bf)sdf"
    P = P12(3, 9)
    Inl = True
    Ept = "(as bf)"
    GoTo Tst
Tst:
    Act = BetP12(S, P)
    C
    Return
End Sub
Function BetP12$(S, P As P12, Optional Inl As Boolean)
Dim L%: L = LenP12(P): If L = 0 Then Exit Function
Dim Pos%: Pos = IIf(Inl, P.P1, P.P1 + 1)
BetP12 = Mid(S, Pos, L)
End Function
Function Bef$(S, Ssub$, Optional NoTrim As Boolean, Optional C As eCas)
Dim P%: P = InStr(1, S, Ssub, VbCprMth(C)): If P = 0 Then Exit Function
Bef = Left(S, P - 1)
If Not NoTrim Then Bef = Trim(Bef)
End Function
Function RmvBef$(S, Ssub$, Optional NoTrim As Boolean): RmvBef = Brk2(S, Ssub, NoTrim).S2: End Function

Function BefAt(S, At&)
If At = 0 Then Exit Function
BefAt = Left(S, At - 1)
End Function

Function BefMust$(S, Ssub$, Optional NoTrim As Boolean):      BefMust = Brk(S, Ssub, NoTrim).S1:        End Function
Function BefOrAll$(S, Ssub$, Optional NoTrim As Boolean):    BefOrAll = Brk1(S, Ssub, NoTrim).S1:       End Function
Function BefOrAllRev$(S, Ssub$):                          BefOrAllRev = StrDft(BefRev(S, Ssub), Ssub$): End Function

Private Sub B_BefFstLas()
Dim S, Fst$, Las$
S = " A_1$ = ""Function ZChunk$(ConstLy$(), IChunk%)"" & _"
Fst = vbQuoDbl
Las = vbQuoDbl
Ept = "Function ZChunk$(ConstLy$(), IChunk%)"
GoSub Tst
Exit Sub
Tst:
    Act = BetFstLas(S, Fst, Las)
    C
    Return
End Sub

Function BetFstLas$(S, Fst$, Las$):     BetFstLas = BefRev(Aft(S, Fst), Las): End Function
Function BetLng(L&, A&, B&) As Boolean:    BetLng = A <= L And L <= B:        End Function

Function BetStr$(S, S1$, S2$, Optional NoTrim As Boolean, Optional InlMarker As Boolean)
With Brk1(S, S1, NoTrim)
   If .S2 = "" Then Exit Function
   Dim O$: O = Brk1(.S2, S2, NoTrim).S1
   If InlMarker Then O = S1 & O & S2
   BetStr = O
End With
End Function

Private Sub B_AftBkt()
Dim A$
A = "(lsk(aa)df lsdkfj) A"
Ept = " A"
GoSub Tst
Exit Sub
Tst:
    Act = AftBkt(A)
    C
    Return
End Sub

Private Sub B_Bet()
Dim Ln
Ln = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??       | DATABASE= | ; | ??":            GoSub Tst
Ln = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=??;AA=XX | DATABASE= | ; | ??":            GoSub Tst
Ln = "lkjsdf;dkfjl;Data Source=Johnson;lsdfjldf  | Data Source= | ; | Johnson":    GoSub Tst
Exit Sub
Tst:
    Dim FmStr$, ToStr$
    AsgAy AmTrim(SplitVBar(Ln)), Ln, FmStr, ToStr, Ept
    Act = IsBet(Ln, FmStr, ToStr)
    C
    Return
End Sub

Private Sub B_BetBkt()
Dim A$
Ept = "1234()567": A = "sdklfjdsf(1234()567)aaa(": GoSub Tst
Ept = "AA":        A = "XXX(AA)XX":                GoSub Tst
Ept = "A$()A":     A = "(A$()A)XX":                GoSub Tst
Ept = "O$()":      A = "(O$()) As X":              GoSub Tst
Ept = "1234()567": A = "sdklfjdsf(1234()567)aaa(": GoSub Tst
Exit Sub

Tst:
    Act = BetBkt(A)
    C
    Return
End Sub

Function BefRevOrAll$(S, Ssub$)
Dim P%: P = InStrRev(S, Ssub)
If P = 0 Then BefRevOrAll = S: Exit Function
BefRevOrAll = Left(S, P - Len(Ssub))
End Function
