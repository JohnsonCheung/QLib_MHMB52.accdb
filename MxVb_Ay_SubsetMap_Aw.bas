Attribute VB_Name = "MxVb_Ay_SubsetMap_Aw"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_Subset_AwSyw."
Enum eThwOutRgeEr: eNoThwOutRge: eThwOutRge: End Enum
Function AwBE(Ay, Bix, Eix)
AwBE = AyNw(Ay)
Dim J&: For J = Bix To Eix
    Push AwBE, Ay(J)
Next
End Function
Function SywBei(Ly$(), B As Bei) As String(): SywBei = AwBei(Ly, B): End Function
Function AwBef(Ay, IxBef)
If IxBef = 0 Then AwBef = AyNw(Ay): Exit Function
If IxBef > UB(Ay) Then AwBef = Ay: Exit Function
AwBef = AyReDim(Ay, IxBef - 1)
End Function
Function AwBei(Ay, B As Bei):                   AwBei = AwBE(Ay, B.Bix, B.Eix):     End Function
Function AwBN(Ay, Bix, N):                       AwBN = AwBE(Ay, Bix, Bix + N - 1): End Function
Function LinesAwBN$(Ly$(), Bix, N):         LinesAwBN = JnCrLf(SywBN(Ly, Bix, N)):  End Function
Function LinesAwBei$(Ly$(), B As Bei):     LinesAwBei = JnCrLf(SywBei(Ly, B)):      End Function
Function SywBN(Ly$(), Bix, N) As String():      SywBN = AwBN(Ly, Bix, N):           End Function
Function AyyBeiy(Ay, OrdBeiy() As Bei) As Variant()
Dim J%: For J = 0 To UbBei(OrdBeiy)
    PushI AyyBeiy, AwBei(Ay, OrdBeiy(J))
Next
End Function

Private Sub B_AwBefEle()
Dim Ay, Ele, InlEle As Boolean
GoSub T1
GoSub T2
Exit Sub
T1:
    Ay = Array(1, 2, 4, "a", 1, 2)
    Ele = "a"
    InlEle = True
    Ept = Array(1, 2, 4, "a")
    GoTo Tst
T2:
    Ay = Array(1, 2, 4, "a", 1, 2)
    Ele = "a"
    InlEle = False
    Ept = Array(1, 2, 4)
    GoTo Tst
Tst:
    Act = AwBefEle(Ay, Ele, InlEle)
    C
    Return
End Sub
Function AwBefEle(Ay, Ele, Optional InlEle As Boolean)
Const CSub$ = CMod & "AwBefEle"
Dim I&: I = IxEle(Ay, Ele): If I = -1 Then Thw CSub, "No ele is found in Ay", "Ele Ay", Ele, Ay
Dim N&: N = IIf(InlEle, I + 1, I)
AwBefEle = AwFstN(Ay, N)
End Function
Function AwFstNSrt(Ay, Optional N = 0): AwFstNSrt = AwFstN(AySrtQ(Ay), N): End Function

Function AwFstN(Ay, N)
Const CSub$ = CMod & "AwFstN"
If N = 0 Then AwFstN = Ay: Exit Function
Dim NEle&: NEle = Si(Ay)
If N >= NEle Then
    AwFstN = Ay
Else
    Dim UNw&: UNw = Min(NEle, N) - 1
    Dim O: O = Ay: ReDim Preserve O(UNw)
    AwFstN = O
End If
End Function

Function AwBet(Ay, FmEle, ToEle)
Dim O: O = AyNw(Ay)
Dim I: For Each I In Itr(Ay)
    If IsBet(I, FmEle, ToEle) Then
        Push O, I
    End If
Next
AwBet = O
End Function

Function SywDis(Sy$()) As String(): SywDis = AwDis(Sy): End Function
Function AwDis(Ay)
AwDis = AyNw(Ay)
Dim I: For Each I In Itr(Ay)
    PushNoDup AwDis, I
Next
End Function

Function AwDisAsI(Ay) As Integer():  AwDisAsI = CvInty(AwDis(Ay)): End Function
Function AwDisAsSy(Ay) As String(): AwDisAsSy = CvSy(AwDis(Ay)):   End Function
Function AwDisT1(Ay) As String():     AwDisT1 = AwDis(Tm1yAy(Ay)): End Function
Function AetAwDup(Ay, Optional C As eCas) As Dictionary
Dim O As New Dictionary
Dim E As New Dictionary
O.CompareMode = VbCprMth(C)
E.CompareMode = VbCprMth(C)
Dim I: For Each I In Itr(Ay)
    If E.Exists(I) Then
        If Not O.Exists(I) Then O.Add I, Empty
    Else
        E.Add I, Empty
    End If
Next
Set AetAwDup = O
End Function
Function AwDup(Ay, Optional C As eCas)
Dim Into: Into = AyNw(Ay)
AwDup = IntoItr(Into, AetAwDup(Ay, C).Keys)
End Function

Function SywDup(Sy$(), Optional C As eCas) As String(): SywDup = AwDup(Sy, C): End Function

Function AwEQ(Ay, V)
Dim O: O = Ay: Erase O
Dim I: For Each I In Itr(Ay)
    If I = V Then PushI O, I
Next
AwEQ = O
End Function

Function AwFm(Ay, Fmix)
AwFm = AyNw(Ay): Dim J&: For J = Fmix To UB(Ay)
    Push AwFm, Ay(J)
Next
End Function

Function AwNB(Ay) As String()
Dim I: For Each I In Itr(Ay)
    If Trim(I) <> "" Then PushI AwNB, I
Next
End Function

Function AwGT(Ay, V)
If Si(Ay) <= 1 Then AwGT = Ay: Exit Function
AwGT = AyNw(Ay)
Dim I: For Each I In Ay
    If I > V Then PushI AwGT, I
Next
End Function

Function AwInAet(Ay, Aet As Dictionary)
AwInAet = AyNw(Ay)
Dim I: For Each I In Itr(Ay)
    If Aet.Exists(I) Then Push AwInAet, I
Next
End Function

Function AwAft(Ay, IxAft)
If IxAft > UB(Ay) Then AwAft = AyNw(Ay): Exit Function
If IxAft < 0 Then AwAft = Ay: Exit Function
AwAft = AyNw(Ay)
Dim J&: For J = IxAft + 1 To UB(Ay)
    Push AwAft, Ay(J)
Next
End Function
Function AwAftEle(Ay, Ele, Optional InlEle As Boolean)  ' return all ele aft first found @Ele, may inl the @ele
Dim Bix&: Bix = IxEle(Ay, Ele)
If Bix = -1 Then AwAftEle = AyNw(Ay): Exit Function
If Not InlEle Then Bix = Bix + 1
AwAftEle = AwFm(Ay, Bix)
End Function

Function AwIsNm(Ay) As String()
Dim I: For Each I In Itr(Ay)
    If IsNm(I) Then PushI AwIsNm, I
Next
End Function

Function AwIxy(Ay, Ixy)
AwIxy = AyNw(Ay)
Dim Ix: For Each Ix In Itr(Ixy)
    PushI AwIxy, Ay(Ix)
Next
End Function

Function AwIxyMay(Ay, IxyMay) ' Array where elements at pointed by @Ixy allow empty if the Ix is outside range of @Ay
'#IxMay:Lng:Index-May# an index may outside the range of an array
'#IxyMay:Lngy:Index-May-Array# an index array of element which is :IxMay
If Si(IxyMay) = 0 Then
    AwIxyMay = AyNw(Ay)
    Exit Function
End If
Dim O: O = AyNw(Ay)
    Dim U&: U = UB(IxyMay)
    Dim UAy&: UAy = UB(Ay)
    ReDim Preserve O(U)
    Dim IxMay
    For Each IxMay In Itr(IxyMay)
        If IsBet(IxMay, 0, UAy) Then
            O(IxMay) = Ay(IxMay)
        End If
    Next
AwIxyMay = O
End Function

Function SywKssy(Sy$(), Kssy$()) As String()
Dim Liky$(): Liky = LikyKssy(Kssy)
Dim S: For Each S In Itr(Sy)
    If HitLiky(S, Liky) Then PushI SywKssy, S
Next
End Function

Function AwLasN(Ay, N)
Dim O, J&, I&, U&, Fm&, NewU&
U = UB(Ay)
If U < N Then AwLasN = Ay: Exit Function
O = Ay: Erase O
Fm = U - N + 1
For J = Fm To U
    Push O, Ay(J)
Next
AwLasN = O
End Function

Function AwLE(Ay, V)
If Si(Ay) <= 1 Then AwLE = Ay: Exit Function
AwLE = AyNw(Ay)
Dim I: For Each I In Ay
    If I <= V Then PushI AwLE, I
Next
End Function

Function AwLik(Ay, Lik) As String()
Dim I: For Each I In Itr(Ay)
    If I Like Lik Then PushI AwLik, I
Next
End Function

Function AwLiky(Ay, Liky$()) As String()
Dim I: For Each I In Itr(Ay)
    If HitLiky(I, Liky) Then PushI AwLiky, I
Next
End Function

Function AwLikss(Ay, Likss$) As String(): AwLikss = AwLiky(Ay, SySs(Likss)): End Function

Function AwLT(Ay, V)
If Si(Ay) = 1 Then AwLT = Ay: Exit Function
Dim O: O = Ay: Erase O
Dim I
For Each I In Ay
    If I < V Then PushI O, I
Next
AwLT = O
End Function

Function AwMid(Ay, Fm, Optional L = 0)
AwMid = AyNw(Ay)
Dim J&
Dim E&
    Select Case True
    Case L = 0: E = UB(Ay)
    Case Else:  E = Min(UB(Ay), L + Fm - 1)
    End Select
For J = Fm To E
    Push AwMid, Ay(J)
Next
End Function


Function AwNm(Ay) As String()
Dim Nm: For Each Nm In Itr(Ay)
    If IsNm(Nm) Then PushI AwNm, Nm
Next
End Function

Function AwRx(Ay, Rx As RegExp) As String()
Dim I: For Each I In Itr(Ay)
    If Rx.Test(I) Then PushI AwRx, I
Next
End Function
Function AwPatn(Ay, Patn$) As String()
If Patn = "" Then AwPatn = Ay: Exit Function
AwPatn = AwRx(Ay, Rx(Patn))
End Function

Function AwPfx(Ay, Pfx) As String()
Dim S: For Each S In Itr(Ay)
    If HasPfx(S, Pfx) Then PushI AwPfx, S
Next
End Function
Function AwRxAyAnd(Ay, RxayAnd() As RegExp) As String()
If Si(RxayAnd) = 0 Then AwRxAyAnd = SyAy(Ay): Exit Function
Dim S: For Each S In Itr(Ay)
    Dim R: For Each R In RxayAnd
        If Not CvRx(R).Test(S) Then GoTo Nxt
    Next
    PushI AwRxAyAnd, S
Nxt:
Next
End Function
Function AwRxAyOr(Ay, RxAyOr() As RegExp) As String()
If Si(RxAyOr) = 0 Then AwRxAyOr = SyAy(Ay): Exit Function
Dim S: For Each S In Itr(Ay)
    Dim R: For Each R In RxAyOr
        If Not CvRx(R).Test(S) Then
            PushI AwRxAyOr, S
            GoTo Nxt
        End If
    Next
Nxt:
Next
End Function

Function AwRmvEle(Ay, Ele)
AwRmvEle = AyNw(Ay)
Dim I: For Each I In Itr(Ay)
    If I <> Ele Then PushI AwRmvEle, I
Next
End Function

Function AwSfx(Ay, Sfx$, Optional C As eCas) As String()
Dim I: For Each I In Itr(Ay)
    If HasSfx(I, Sfx, C) Then PushI AwSfx, I
Next
End Function

Function AwSkpN(Ay, Optional SkpN& = 1)
Const CSub$ = CMod & "AwSkpN"
If SkpN <= 0 Then AwSkpN = Ay: Exit Function
Dim U&: U = UB(Ay) - SkpN: If SkpN < -1 Then Thw CSub, "Ay is not enough to skip", "Si-Ay SkipN", "Si(Ay),SKipN"
Dim O: O = Ay: Erase O
Dim J&: For J = SkpN To U
    Push O, Ay(J)
Next
AwSkpN = O
End Function

Function AwSng(Ay): AwSng = AyMinus(Ay, AwDup(Ay)): End Function

Function AwPatnssOr(Ay, Patnss$) As String():                       AwPatnssOr = AwRxAyOr(Ay, RxAyPatnss(Patnss)):  End Function
Function AwPatnssAnd(Ay, Patnss$) As String():                     AwPatnssAnd = AwRxAyAnd(Ay, RxAyPatnss(Patnss)): End Function
Function AwSsubDash(Ay) As String():                                AwSsubDash = AwSsub(Ay, "_"):                   End Function
Function AwSsubssOr(Ay, Ssubss$, Optional C As eCas) As String():   AwSsubssOr = AwSsubyOr(Ay, SySs(Ssubss), C):    End Function
Function AwSsubssAnd(Ay, Ssubss$, Optional C As eCas) As String(): AwSsubssAnd = AwSsubyAnd(Ay, SySs(Ssubss), C):   End Function
Function AwSsubyAnd(Ay, Ssuby$(), Optional C As eCas) As String()
If Si(Ssuby) = 0 Then AwSsubyAnd = SyAy(Ay): Exit Function
Dim I: For Each I In Itr(Ay)
    If HasSsubyAnd(I, Ssuby, C) Then
        PushNB AwSsubyAnd, I
    End If
Next
End Function
Function AwSsubyOr(Ay, Ssuby$(), Optional C As eCas) As String()
If Si(Ssuby) = 0 Then AwSsubyOr = SyAy(Ay): Exit Function
Dim I: For Each I In Itr(Ay)
    If I = "MxIde_Mth_Slm_Ali_AliSlm" Then Stop
    If HasSsubyOr(I, Ssuby, C) Then
        PushNB AwSsubyOr, I
    End If
Next
End Function
Function AwSsub(Ay, Ssub, Optional C As eCas) As String()
Dim I: For Each I In Itr(Ay)
    If HasSsub(I, Ssub, C) Then
        PushI AwSsub, I
    End If
Next
End Function
Function AwT1(Ay, T1) As String()
Dim I: For Each I In Itr(Ay)
    If HasTmo1(I, T1) Then
        PushI AwT1, I
    End If
Next
End Function

Function AwT1y(Ay, T1y$()) As String()
Dim O$(), L: For Each L In Itr(Ay)
    If HasEle(T1y, Tm1(L)) Then Push O, L
Next
AwT1y = O
End Function

Function AwT12(Ay, T1, T2) As String() ' ret subset of @Ay where each ele's fst term = @T1 and snd term = @T2
Dim I: For Each I In Itr(Ay)
    If WAwT12_Has(I, T1, T2) Then PushI AwT12, I
Next
End Function
Private Function WAwT12_Has(Ln, T1, T2) As Boolean ' ret True is fst term of @Ln = @T2 and snd term of @Ln = @T2
Dim L$: L = Ln
If ShfTm(L) <> T1 Then Exit Function
If ShfTm(L) <> T2 Then Exit Function
WAwT12_Has = True
End Function

Function AwPfxx(Ay, Pfxx$) As String()
Dim Pfxy$(): Pfxy = SySs(Pfxx)
Dim I: For Each I In Itr(Ay)
    If HasPfxy(I, Pfxy) Then PushI AwPfxx, I
Next
End Function

Function AwEndTrim(Ly$()) As String()
Dim O$(): O = Ly$()
Dim UNw&
    For UNw = UB(Ly) To 0 Step -1
        If LTrim(Ly(UNw)) <> "" Then Exit For
    Next
    If UNw = -1 Then Exit Function
ReDim Preserve O(UNw)
AwEndTrim = O
End Function
Function AwPred(Ay, Pred$)
Dim O: O = AyNw(Ay)
Dim I: For Each I In Itr(Ay)
    If Run(Pred, I) Then PushI O, I
Next
AwPred = O
End Function
