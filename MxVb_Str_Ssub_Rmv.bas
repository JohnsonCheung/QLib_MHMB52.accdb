Attribute VB_Name = "MxVb_Str_Ssub_Rmv"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Ssub_Rmv."
Function RmvAt$(S, At)
If At <= 0 Then Exit Function
RmvAt = Left(S, At - 1)
End Function
Function RmvA2Hyp$(S): RmvA2Hyp = RTrim(RmvAt(S, InStr(S, "--"))):  End Function
Function RmvA3Hyp$(S): RmvA3Hyp = RTrim(RmvAt(S, InStr(S, "---"))): End Function
Function RmvA3T$(S):     RmvA3T = RmvA2T(RmvA1T(S)):                End Function
Function RmvA2T$(S):     RmvA2T = RmvA1T(RmvA1T(S)):                End Function
Function RmvA1T$(S)
Const CSub$ = CMod & "RmvA1T"
Dim P%
    Dim L$: L = LTrim(S)
    If ChrFst(L) = "[" Then
        P = InStr(L, "]"): If P = 0 Then Thw CSub, "Given S is invalid TmlAy: There is chr-[, but not chr-]", "@S", S
        P = P + 1
    Else
        P = InStr(L, " ")
        If P = 0 Then
            P = Len(L) + 1
        End If
    End If
RmvA1T = LTrim(Mid(S, P))
End Function
Function RmvP12$(S, P As P12, Optional Inl As Boolean) ' Rmv the str point by @Bet inclusive
Dim L%: L = LenP12(P, Inl): If L = 0 Then Exit Function
RmvP12 = Mid(S, IIf(Inl, P.P1 + 1, P.P1), L)
End Function

Function RmvAft$(S, AftSsub$, Optional C As eCas)
If AftSsub = "" Then ThwPm CSub, "AftSsub cannot be blank", "S", S
Dim P%: P = PosSsub(S, AftSsub, C): If P = 0 Then RmvAft = S: Exit Function
RmvAft = Left(S, P + Len(AftSsub) - 1)
End Function
Function RmvFm$(S, FmSsub$, Optional C As eCas)
If FmSsub = "" Then ThwPm CSub, "FmSsub cannot be blank", "S", S
Dim P%: P = PosSsub(S, FmSsub, , C): If P = 0 Then RmvFm = S: Exit Function
RmvFm = Left(S, P - 1)
End Function

Function RmvDblSpc$(S) ' Rpl more than one spc to one.
Dim O$: O = S
While HasSsub(O, "  ")
    O = Replace(O, "  ", " ")
Wend
RmvDblSpc = O
End Function

Function RmvFst$(S):       RmvFst = Mid(S, 2):         End Function
Function RmvFst2$(S):     RmvFst2 = Mid(S, 3):         End Function
Function RmvFstLas$(S): RmvFstLas = RmvFst(RmvLas(S)): End Function
Function RmvFstNonLetter$(S)
If IsAscLetter(Asc(S)) Then
    RmvFstNonLetter = S
Else
    RmvFstNonLetter = RmvFst(S)
End If
End Function
Function RmvFstN$(S, N)
If N <= 0 Then RmvFstN = S: Exit Function
RmvFstN = Mid(S, N + 1)
End Function
Function RmvLas2$(S): RmvLas2 = RmvLasN(S, 2): End Function
Function RmvLas$(S):   RmvLas = RmvLasN(S, 1): End Function
Function RmvLasN$(S, N)
Dim L&: L = Len(S) - N: If L <= 0 Then Exit Function
RmvLasN = Left(S, L)
End Function

Function RmvNm$(S)
If Not IsAscFstNmChr(Asc(ChrFst(S))) Then GoTo X
Dim O%: For O = 1 To Len(S)
    If Not IsAscNmChr(Asc(Mid(S, O, 1))) Then GoTo X
Next
X:
    If O > 0 Then RmvNm = Mid(S, O): Exit Function
    RmvNm = S
End Function

Function AmRmvBktSq(Sy$()) As String()
Dim I
For Each I In Itr(Sy)
    PushI AmRmvBktSq, RmkBktSq(I)
Next
End Function
Function RmkBktSq$(S)
If Not HasBktSq(S) Then RmkBktSq = S: Exit Function
RmkBktSq = RmvFstLas(S)
End Function

Function RmvPfxAll$(S, Pfx$, Optional C As eCas)
Const CSub$ = CMod & "RmvPfxAll"
Dim O$: O = S
Dim J%
While HasPfx(O, Pfx, C)
    ThwLoopTooMuch CSub, J
    O = RmvPfx(O, Pfx, C)
Wend
RmvPfxAll = O
End Function
Function HasPSfx(S, Pfx$, Sfx$, Optional C As eCas) As Boolean
HasPSfx = HasPfx(S, Pfx, C) And HasSfx(S, Sfx, C)
End Function
Function RmvPSfx$(S, Pfx$, Sfx$, Optional C As eCas)
If HasPSfx(S, Pfx, Sfx, C) Then
    RmvPSfx = RmvSfx(RmvPfx(S, Pfx, C), Sfx, C)
Else
    RmvPSfx = S
End If
End Function
Function RmvPfx$(S, Pfx, Optional C As eCas) ' Always Case Sensitive
If HasPfx(S, Pfx, C) Then RmvPfx = Mid(S, Len(Pfx) + 1) Else RmvPfx = S
End Function

Function RmvPfxy$(S, Pfxy$(), Optional C As eCas) ' Rmv one of the pfx in @Pfxy from @S
Dim P: For Each P In Pfxy
    If HasPfx(S, P, C) Then RmvPfxy = RmvPfx(S, P, C): Exit Function
Next
RmvPfxy = S
End Function
Function RmvPfxSpc$(S, Pfx, Optional C As eCas)
If Not HasPfxSpc(S, Pfx, C) Then RmvPfxSpc = S: Exit Function
RmvPfxSpc = Mid(S, Len(Pfx) + 2)
End Function
Function RmvPfxySpc$(S, Pfxy$(), Optional C As eCas)
Dim Pfx: For Each Pfx In Pfxy
    If HasPfxSpc(S, Pfx, C) Then
        RmvPfxySpc = RmvPfxSpc(S, Pfx, C)
        Exit Function
    End If
Next
RmvPfxySpc = S
End Function

Function RmvBktSq$(S):                             RmvBktSq = RmvBkt(S, vbBktOpnSq):              End Function
Function RmvBktBig$(S):                           RmvBktBig = RmvBkt(S, vbBktOpnBig):             End Function
Function RmvBkt$(S, Optional BktOpn$ = vbBktOpn):    RmvBkt = RmvPSfx(S, BktOpn, BktCls(BktOpn)): End Function
Function RmvSfxAp$(S, ParamArray SfxAp()): Dim Av(): Av = SfxAp: RmvSfxAp = RmvSfxAv(S, Av): End Function
Function RmvSfxAv$(S, SfxAv())
Dim O$: O = S
Dim Sfx: For Each Sfx In SfxAv
    O = RmvSfx(O, Sfx)
Next
RmvSfxAv = O
End Function

Function RmvSfxDot$(S): RmvSfxDot = RmvSfx(S, "."): End Function
Function RmvSfx$(S, Sfx, Optional C As eCas)
If HasSfx(S, Sfx, C) Then RmvSfx = Left(S, Len(S) - Len(Sfx)) Else RmvSfx = S
End Function

Function RmvQuoSng$(S)
If Not IsQuoSng(S) Then RmvQuoSng = S: Exit Function
RmvQuoSng = RmvFstLas(S)
End Function

Function TmlRmvX$(ATml$, TmX$): TmlRmvX = Tml(SyeEle(Tmy(ATml), TmX)): End Function

Private Sub B_RmvA1T()
Ass RmvA1T("  df dfdf  ") = "dfdf"
End Sub

Private Sub B_RmvNm()
Dim Nm$
Nm = "lksdjfsd f"
Ept = " f"
GoSub Tst
Exit Sub
Tst:
    Act = RmvNm(Nm)
    C
    Return
End Sub

Private Sub B_RmvPfx()
Ass RmvPfx("aaBB", "aa") = "BB"
End Sub

Private Sub B_RmvPfxy()
GoSub T1
Exit Sub
Dim S, Pfxy$(), Cas As eCas
T1:
    Pfxy = SySs("z_ B__"): Ept = "ABC"
    S = "B_ABC"
    Ept = "B_ABC"
    Cas = eCasSen
    GoTo Tst
T2:
    Pfxy = SySs("B__ B_"): Ept = "ABC"
    S = "B__ABC"
    Ept = "ABC"
    GoTo Tst
Tst:
    Act = RmvPfxy(S, Pfxy)
    C
    Return
End Sub

Function RmvCr$(S)
RmvCr = Replace(S, vbCr, "")
End Function
Function RmvEndDig$(S)
Dim J&: For J = Len(S) To 1 Step -1
    If Not IsAscDig(Asc(Mid(S, J, 1))) Then
        RmvEndDig = Left(S, J)
        Exit Function
    End If
Next
RmvEndDig = Left(S, J)
End Function
