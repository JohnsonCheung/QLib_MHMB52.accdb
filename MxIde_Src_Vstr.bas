Attribute VB_Name = "MxIde_Src_Vstr"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Vstr."

Function EndPos%(Fm%, S, Lvl%)
Const CSub$ = CMod & "EndPos"
ThwLoopTooMuch CSub, Lvl
Dim P%: P = InStr(Fm, S, """"): If P = 0 Then Exit Function
If Mid(S, P + 1, 1) <> """" Then EndPos = P: Exit Function
EndPos = EndPos(P + 2, S, Lvl + 1)
End Function

Private Sub B_RmvVstr()
Dim Ln$
GoSub Z
Exit Sub
Z:
    Ln = "aa""""aa""""": Ept = "aa""bb"
    Ept = "aa""""bb"
    GoTo Tst
Tst:
    Act = RplVstr(Ln)
    C
    Return
End Sub

Private Sub B_SrcRplVstr()
TimFun "WTst1"
TimFun "WTst2" ' Quicker
End Sub
Function SrcRplVstr(Src$()) As String()
Dim L: For Each L In Itr(Src)
    PushI SrcRplVstr, RplVstr(L)
Next
End Function

Private Sub B_RplVstr()
Dim LyBef$(): LyBef = W1LyBef
Dim LyAft$(): LyAft = W1LyAft(LyBef)
CprLy LyBef, LyAft
End Sub
Private Function W1LyBef() As String()
Dim L: For Each L In SrcPC
    If HasQuoDbl(L) Then PushI W1LyBef, L
Next
End Function
Private Function W1LyAft(LyBef$()) As String()
Dim L: For Each L In Itr(LyBef)
    PushI W1LyAft, RplVstr(L)
Next
End Function
Function RplVstr$(Ln) '#Rmv-Vstr-by-same-len-spc#
Const CSub$ = CMod & "RmvVstr"
If Not HasQuoDbl(Ln) Then RplVstr = Ln: Exit Function
Dim L As S12: L = BrkVmk(Ln)
Dim P() As P12: P = WP12yVstr(L.S1)
Dim O$: O = L.S1
Dim J%: For J = 0 To UbP12(P)
    O = RplP12Spc(O, P(J))
Next
RplVstr = O & L.S2
If RplVstr = "" Then Stop
End Function
Private Function WP12yVstr(LnNoVmk) As P12()
Dim IsInside As Boolean
Dim I%: I = 1
Dim M$
Dim L%: L = Len(LnNoVmk)
Dim S$: S = LnNoVmk
Dim O() As P12
Dim J%
Again:
    ThwLoopTooMuch CSub, J, 100
    If I > L Then GoTo Ext
    Dim P1%: P1 = PosQuoDbl(S, I): If P1 = 0 Then GoTo Ext
    Dim P2%: P2 = WPosVquoNxt(S, P1 + 1): If P2 = 0 Then M = "With open-Vquo, but no close-Vquo": GoTo M
    PushP12 O, P12(P1, P2)
    I = P2 + 1
    GoTo Again
M: Thw CSub, M, "LnNoVmk", LnNoVmk
Ext: WP12yVstr = O: Exit Function
End Function
Function RmvVstr$(L): RmvVstr = RmvP12y(L, P12yVstr(L)): End Function
Function P12yVstr(LnNoVmk) As P12()
Dim Pi%: Pi = 1 'PosI
Dim P1%, P2%    'PosVqmk1/2
Again:
    If Pi > LnNoVmk Then Exit Function
    Dim PQ%: PQ = InStr(Pi, LnNoVmk, vbQuoDbl) 'PosQmkRunning
    Select Case True
    Case P1 = 0 And PQ = 0: Exit Function
    Case P1 = 0: P1 = PQ: Pi = PQ + 1
    Case P1 <> 0 And PQ = 0: Thw CSub, "No PosVqmk2", "PosVqmk1 LnNoVmk", P1, LnNoVmk
    Case Else:
        Dim CntLoop%
        P2 = PosQuoDbl2AdjRc(LnNoVmk, P2, CntLoop): Thw CSub, "Inbalance Vb-Quo-Rmk", "Ln-Without-VbRmk", LnNoVmk
        PushP12 P12yVstr, P12(P1, P2)
        Pi = P2 + 1: P1 = 0: P2 = 0
    End Select
    GoTo Again
End Function
Function PosQuoDbl2AdjRc%(LnNoVmk, PosQuoDbl2%, OCntLoop%)
If OCntLoop > 100 Then Stop: Exit Function
Dim ChrNxt$: ChrNxt = Mid(LnNoVmk, PosQuoDbl2 + 1, 1)
If ChrNxt <> vbQuoDbl Then PosQuoDbl2AdjRc = PosQuoDbl2: Exit Function
Dim PosNxtQuoDbl2%: PosNxtQuoDbl2 = InStr(PosQuoDbl2 + 2, LnNoVmk, vbQuoDbl)
If PosNxtQuoDbl2 = 0 Then Thw CSub, "LnNoVmk is imbalance QuoDbl", "LnNoVmk", LnNoVmk
PosQuoDbl2AdjRc = PosQuoDbl2AdjRc(LnNoVmk, PosNxtQuoDbl2, OCntLoop + 1)
End Function
Function RmvP12y$(S, OrdP12() As P12)
Dim O$: O = S
Dim J%: For J = UbP12(OrdP12) To 0 Step -1
    O = RmvP12(O, OrdP12(J))
Next
RmvP12y = O
End Function
Function ShfVstr$(OLn$)
Dim M$
If ChrFst(OLn) <> vbQuoDbl Then M = "ChrFst <>'""'": GoTo M
Dim E%: E = WPosVquoNxt(OLn, 2): If E = 0 Then M = "Given @Ln [Assume ChrFst is vbQuoDbl] has no VquoNxt"
OLn = Mid(OLn, E + 1)
ShfVstr = Mid(OLn, 2, E - 2)
Exit Function
M: Thw CSub, M, "@OLn", OLn
End Function

Private Function WPosVquoNxt%(S, PosVquo%)
Dim B%: B = PosVquo
Dim P%, J%
Do
    P = InStr(B, S, vbQuoDbl): If P = 0 Then Exit Function
    If Mid(S, P + 1, 1) <> vbQuoDbl Then WPosVquoNxt = P: Exit Function
    B = P + 2
    ThwLoopTooMuch "PosVmkNxt", J
Loop
End Function
Function TakPfx$(S, Pfx, Optional C As eCas)
If HasPfx(S, Pfx, C) Then TakPfx = Left(S, Len(Pfx))
End Function
