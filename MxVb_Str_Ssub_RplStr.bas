Attribute VB_Name = "MxVb_Str_Ssub_RplStr"
':Q: :S #Str-With-QuestionMark#
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Ssub_Rpl."
Const cInl As Boolean = True
Const cExl As Boolean = False
Private Sub B_RplBet()
Dim S$, Exp$, By$, S1$, S2$
S1 = "Data Source="
S2 = ";"
S = "aa;Data Source=???;klsdf"
By = "xx"
Exp = "aa;Data Source=xx;klsdf"
GoSub Tst
Exit Sub
Tst:
Dim Act$
Act = RplBet(S, By, S1, S2)
Debug.Assert Exp = Act
Return
End Sub

Private Sub B_RplPfx(): Ass RplPfx("aaBB", "aa", "xx") = "xxBB": End Sub

Function RplCr$(S):     RplCr = Replace(S, vbCr, " "):   End Function
Function RplCrLf$(S): RplCrLf = RplLf(RplCr(S)):         End Function
Function RplLf$(S):     RplLf = Replace(S, vbLf, " "):   End Function
Function RplVbl$(S):   RplVbl = RplVBar(S):              End Function
Function RplVBar$(S): RplVBar = Replace(S, "|", vbCrLf): End Function
Function RplBet$(S, By$, S1$, S2$)
Dim P1%, P2%, B$, C$
P1 = InStr(S, S1)
If P1 = 0 Then Stop
P2 = InStr(P1 + Len(S1), S, S2)
If P2 = 0 Then Stop
B = Left(S, P1 + Len(S1) - 1)
C = Mid(S, P2 + Len(S2) - 1)
RplBet = B & By & C
End Function
Function Rpl2DblQ$(S):   Rpl2DblQ = Replace(S, vbQuoDbl2, ""): End Function
Function RplDblSpc$(S): RplDblSpc = Trim(RplDblChr(S, " ")):   End Function
Function RplDblChr$(S, Chr$)
Dim D$: D = Chr & Chr
Dim O$: O = S
Dim J&
While HasSsub(O, D, eCasSen)
    J = J + 1: If J > 10000 Then Stop
    O = Replace(O, D, Chr)
Wend
RplDblChr = O
End Function

Function RplFstChr$(S, By$): RplFstChr = By & RmvFst(S): End Function
Function DisChry(S) As String() ' ret IsDis-Chry from S
Dim O$(), J&: For J = 1 To Len(S): PushI O, Mid(S, J, 1): Next
DisChry = AwDis(O)
End Function
Function RplPfx$(S, PfxFm$, PfxTo$, Optional C As eCas)
If HasPfx(S, PfxFm, C) Then
    RplPfx = PfxTo & RmvPfx(S, PfxFm, C)
Else
    RplPfx = S
End If
End Function

Function RplAlpNum$(S):         RplAlpNum = RxChrAlpNum.Replace(S, " "): End Function
Function RplQ$(Tpq$, By):            RplQ = Replace(Tpq, "?", By):       End Function
Function RplQVbar$(QVbar$, By):  RplQVbar = RplVBar(RplQ(QVbar, By)):    End Function
Private Sub B_SyPosy()
GoSub T1
Exit Sub
Dim Posy%(), S$
T1:
    '12345678
    'ab,cde,f
    Posy = Inty(3, 7)
    S = "ab,cde,f"
    Ept = Sy("ab", "cde", "f")
    GoTo Tst
Tst:
    Act = SyPosy(S, Posy)
    C
    Return
End Sub

Private Sub B_RplBet3()
Dim S$, Exp$, By$, S1$, S2$
S1 = "Data Source="
S2 = ";"
S = "aa;Data Source=???;klsdf"
By = "xx"
Exp = "aa;Data Source=xx;klsdf"
GoSub Tst
Exit Sub
Tst:
Dim Act$
Act = RplBet(S, By, S1, S2)
Debug.Assert Exp = Act
Return
End Sub
Function RplP12Spc$(S, P As P12, Optional Inl As Boolean)
Dim L$, R$
    L = P.P1 - IIf(Inl, 0, 1)
    R = P.P2 + IIf(Inl, 0, 1)
RplP12Spc = L & Space(LenP12(P, Inl)) & R
End Function
Function RplP12$(S, P As P12, By$, Optional Inl As Boolean)
If P.P1 = 0 Then RplP12 = S: Exit Function
Dim L%, R%
    L = P.P1 - IIf(Inl, 1, 0)
    R = P.P2 + IIf(Inl, 1, 0)
RplP12 = Left(S, L) & By & Mid(S, R)
End Function
Function Rpl$(S, Ssub$, By$, Optional C As eCas, Optional Ith% = 1)
Dim P&: P = PosSsub(S, Ssub, C)
If P = 0 Then Rpl = S: Exit Function
Rpl = Replace(S, Ssub, By, P, 1)
End Function
