Attribute VB_Name = "MxVb_Ay_SubsetMap_Ae"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_Op_Sel_AeSye."

Private Sub B_AeAtCnt()
Dim Ay(), At&, Cnt&
GoSub YY
Exit Sub
YY:
    Ay = Array(1, 2, 3, 4, 5)
    At = 1
    Cnt = 2
    Ept = Array(1, 4, 5)
    GoTo Tst
Tst:
    Act = AeAtCnt(Ay, 1, 2)
    C
    Return
End Sub

Function AeAtCnt(Ay, Optional At = 0, Optional Cnt = 1)
Stop
Const CSub$ = CMod & "AeAtCnt"
If Cnt <= 0 Then Thw CSub, "Cnt cannot <=0", "At Cnt Ay", At, Cnt, Ay
If Si(Ay) = 0 Then AeAtCnt = Ay: Exit Function
Dim U&: U = UB(Ay)
ChkBet CSub, At, 0, U
Dim NewU&: NewU = U - Cnt
    If At = U - Cnt + 1 Then
        AeAtCnt = AyReDim(Ay, NewU)
        Exit Function
    End If
Dim O: O = Ay
Dim J&: For J = At To U - Cnt
    Asg O(J + Cnt), O(J)
Next
AeAtCnt = AyReDim(Ay, NewU)
End Function

Function AeBlnkAtEnd(A$()) As String()
If Si(A) = 0 Then Exit Function
If EleLas(A) <> "" Then AeBlnkAtEnd = A: Exit Function
Dim J%
For J = UB(A) To 0 Step -1
    If Trim(A(J)) <> "" Then
        Dim O$()
        O = A
        ReDim Preserve O(J)
        AeBlnkAtEnd = O
        Exit Function
    End If
Next
End Function

Function SyeEle(Sy$(), Ele$, Optional C As eCas) As String()
Dim S: For Each S In Itr(Sy)
    If Not IsEqStr(S, Ele, C) Then PushI SyeEle, S
Next
End Function

Function AeBei(Ay, B As Bei)
If IsEmpBei(B) Then AeBei = Ay: Exit Function
AeBei = AyAdd(AwBef(Ay, B.Bix), AwAft(Ay, B.Eix))
End Function
Function AeBeiy(Ay, B() As Bei)
If SiBei(B) = 0 Then AeBeiy = Ay: Exit Function
AeBeiy = AyAdd(AwBef(Ay, B(0).Bix), AeBeiyPart2(Ay, B))
End Function
Function AeBeiyPart2(Ay, B() As Bei)
Dim N&: N = SiBei(B)
If N = 0 Then ThwPm CSub, "@Beiy must have at least 1 element"
Dim EixLas&: EixLas = B(0).Eix
Dim J&: For J = 1 To N - 1
    Dim BixCur&, EixCur&
        With B(J): BixCur = .Bix: BixCur = .Eix: End With
    Dim BixTak&, EixTak&
        BixTak = B(J - 1).Eix + 1
        EixTak = BixCur - 1
    PushI AeBeiyPart2, AwBE(Ay, BixTak, EixTak)
Next
End Function
Function AeEle(Ay, Ele) 'Rmv Fst-Ele eq to Ele from Ay
AeEle = Ay
Erase AeEle
Dim I: For Each I In Itr(Ay)
    If I <> Ele Then PushI AeEle, I
Next
End Function

Function AeAt(Ay, Optional At& = 0, Optional Cnt& = 1)
AeAt = AeAtCnt(Ay, At, Cnt)
End Function

Function AeEmpEle(Ay)
Dim O: O = AyNw(Ay)
If Si(Ay) > 0 Then
    Dim X
    For Each X In Itr(Ay)
        PushNonEmp O, X
    Next
End If
AeEmpEle = O
End Function

Function AeEmpEleAtEnd(Ay)
Dim LasU&, U&
Dim O: O = Ay
For LasU = UB(Ay) To 0 Step -1
    If Not IsEmp(O(LasU)) Then
        Exit For
    End If
Next
If LasU = -1 Then
    Erase O
Else
    ReDim Preserve O(LasU)
End If
AeEmpEleAtEnd = O
End Function
Function RevBeiy(OrdBeiy() As Bei, U&) As Bei()
Dim BeiU&: BeiU = UbBei(OrdBeiy)
If BeiU = -1 Then
    PushBei RevBeiy, Bei(0, U)
    Exit Function
End If
Dim O() As Bei
Dim J%: For J = 0 To BeiU
    Dim WBei As Bei: WBei = OrdBeiy(J)
    Select Case True
    Case J = 0
        If WBei.Bix > 0 Then
            PushBei O, Bei(0, WBei.Bix - 1)
        End If
    Case Else
        If OrdBeiy(BeiU).Eix < U Then
            PushBei O, Bei(WBei.Eix + 1, U)
        End If
    End Select
Next
Dim E&: E = WBei.Eix
If E < U Then
    PushBei O, Bei(E + 1, U)
End If
RevBeiy = O
End Function
Function AeFst(Ay)
Const CSub$ = CMod & "AeFst"
Dim N&: N = Si(Ay): If N = 0 Then Thw CSub, "Given Ay is empty"
Dim O: O = Ay: If N = 1 Then Erase O: AeFst = O: Exit Function
Dim J&: For J = 0 To N - 2
    O(J) = O(J + 1)
Next
ReDim Preserve O(N - 2)
AeFst = O
End Function

Function AeFstLas(Ay): AeFstLas = AeFst(AeLas(Ay)): End Function
Function AeFstN(Ay, Optional N = 1)
Dim O: O = AyNw(Ay)
Dim J&
For J = N To UB(Ay)
    Push O, Ay(J)
Next
AeFstN = O
End Function

Function AeIxSet(Ay, IxSet As Dictionary)
Dim O: O = Ay: Erase O
Dim J&: For J = 0 To UBound(Ay)
    If Not IxSet.Exists(J) Then PushI O, Ay(J)
Next
AeIxSet = O
End Function

Function AeIx(Ay, Ix)
Dim U&: U = UB(Ay)
Dim O
Select Case True
Case Not IsBet(Ix, 0, U)
    AeIx = Ay
Case Ix = U
    O = Ay: ReDim Preserve O(U - 1)
    AeIx = O
Case Else
    O = Ay
    O = Ay: ReDim Preserve O(U - 1)
    Dim J&: For J = Ix To U - 1
        O(J) = Ay(J + 1)
    Next
    AeIx = O
End Select
End Function
Function AeIxy(Ay, Ixy): AeIxy = AwIxy(Ay, IxyInl(Ixy, UB(Ay))): End Function
Function IxyInl(IxyExl, U&) As Long()
Dim O&()
Dim Ix&: For Ix = 0 To U
    If Not HasEle(IxyExl, Ix) Then
        PushI IxyInl, Ix
    End If
Next
End Function

Private Sub B_AeLikk()
Dim Sy$(), Kss$
GoSub Z
GoSub T0
Exit Sub
T0:
    Sy = SySs("A B C CD E E1 E3")
    Kss = "C* E*"
    Ept = SySs("A B")
    GoTo Tst
Z:
    D AeLikk(SySs("A B C CD E E1 E3"), "C* E*")
    Return
Tst:
    Act = AeLikk(Sy, Kss)
    C
    Return
End Sub
Function AeLikk(Ay, Likk) As String()
Dim O: O = Ay
Dim Lik: For Each Lik In SySs(Likk)
    O = AeLik(O, Lik)
Next
AeLikk = O
End Function

Function AeLas(Ay)
Dim U&: U = UB(Ay)
If U = -1 Then Thw CSub, "No Las Ele in @Ay", "Ay-TypeName", TypeName(Ay)
Dim O: O = Ay: ReDim Preserve O(U - 1)
AeLas = O
End Function

Function AeLasN(Ay, Optional NEle% = 1)
If NEle = 0 Then AeLasN = Ay: Exit Function
Dim O: O = Ay
Select Case Si(Ay)
Case Is > NEle:    ReDim Preserve O(UB(Ay) - NEle)
Case NEle: Erase O
Case Else: Stop
End Select
AeLasN = O
End Function

Function AeLik(Ay, Lik) As String()
Dim I: For Each I In Itr(Ay)
    If Not I Like Lik Then PushI AeLik, I
Next
End Function

Function AeNegative(AyNum)
AeNegative = AyNw(AyNum)
Dim I: For Each I In Itr(AyNum)
    If I >= 0 Then
        PushI AeNegative, I
    End If
Next
End Function

Function AePfxx(Ay, Pfxx$) As String()
Dim Pfxy$(): Pfxy = SySs(Pfxx)
Dim I: For Each I In Itr(Ay)
    If Not HasPfxy(I, Pfxy) Then PushI AePfxx, I
Next
End Function
Function AePfx(Ay, Pfx$) As String()
Dim I: For Each I In Itr(Ay)
    If Not HasPfx(I, Pfx) Then PushI AePfx, I
Next
End Function
Function AeSfx(Ay, Sfx$) As String()
Dim I: For Each I In Itr(Ay)
    If Not HasSfx(I, Sfx) Then PushI AeSfx, I
Next
End Function

Function AePatnAp(Ay, ParamArray Patn()) As String()
Dim O$(): O = SyAy(Ay)
Dim Av(): Av = Patn
Dim P: For Each P In Av
    O = AePatn(O, CStr(P))
Next
End Function
Function AePatn(Ay, Patn$) As String()
Dim R As RegExp: Set R = Rx(Patn)
Dim S: For Each S In Itr(Ay)
    If Not HasRx(S, R) Then PushI AePatn, S
Next
End Function
Function AeSsubDash(Ay) As String(): AeSsubDash = AeSsub(Ay, "_"): End Function

Function AeSsub(Ay, Ssub, Optional C As eCas) As String()
Dim S: For Each S In Itr(Ay)
    If Not HasSsub(S, Ssub) Then PushI AeSsub, S
Next
End Function
Function AeSsubssOr(Ay, SsubssOr$, Optional C As eCas) As String()
Dim Ssuby$(): Ssuby = SySs(SsubssOr)
Dim S: For Each S In Itr(Ay)
    If Not HasSsubyOr(S, Ssuby) Then PushI AeSsubssOr, S
Next
End Function

Function AeEmp(Ay)
AeEmp = AyNw(Ay)
Dim I: For Each I In Ay
    If Not IsEmpty(I) Then PushI AeEmp, I
Next
End Function

Function AeKssy(Ay, Kssy$()) As String()
Dim Kss: For Each Kss In Kssy
    AeKssy = AeLikk(AeKssy, Kss)
Next
End Function

Function AeLiky(Ay, Liky$()) As String()
AeLiky = Ay
Dim Lik: For Each Lik In Liky
    AeLiky = AeLik(AeLiky, Lik)
Next
End Function

Function LyExlT1y(Ly$(), ExlT1y$()) As String(): LyExlT1y = AeT1y(Ly, ExlT1y): End Function 'Exclude those Ln in Array-Ay its T1 in ExlAmT10
Function AeT1y(Ay, ExlT1y$()) As String()
Dim L: For Each L In Itr(Ay)
    If Not HitT1y(Tm1(L), ExlT1y) Then PushI AeT1y, L
Next
End Function

Private Sub B_AeEmpEleAtEnd()
Dim Ay: Ay = Array(Empty, Empty, Empty, 1, Empty, Empty)
Dim Act: Act = AeEmpEleAtEnd(Ay)
Ass Si(Act) = 4
Ass Act(3) = 1
End Sub

Private Sub B_AeEmpEleAtEnd1()
Dim Ay: Ay = Array(Empty, Empty, Empty, 1, Empty, Empty)
Dim Act: Act = AeEmpEleAtEnd(Ay)
Ass Si(Act) = 4
Ass Act(3) = 1
End Sub

Private Sub B_AeBei()
Dim Ay
Dim Bei1 As Bei
Dim Act
Ay = SplitSpc("a b c d e")
Bei1 = Bei(1, 2)
Act = AeBei(Ay, Bei1)
Ass Si(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub

Private Sub B_AeBei1()
Dim Ay
Dim Act
Ay = SplitSpc("a b c d e")
Act = AeBei(Ay, Bei(1, 2))
Ass Si(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub

Private Sub B_AeIxy()
Dim Ay(), Ixy
Ay = Array("a", "b", "c", "d", "e", "f")
Ixy = Array(1, 3)
Ept = Array("a", "c", "e", "f")
GoSub Tst
Exit Sub
Tst:
    Act = AeIxy(Ay, Ixy)
    C
    Return
End Sub


