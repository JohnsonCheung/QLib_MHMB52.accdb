Attribute VB_Name = "MxVb_Ay_SubsetMap_Am"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_A_Am."

Function AmAddIxPfx(Ay, Optional Bix% = 1) As String()
Dim U&: U = UB(Ay)
Dim W%: W = NDig(U + Bix)
Dim J&: For J = 0 To U
    PushI AmAddIxPfx, AliR(J + Bix, W) & ": " & Ay(J)
Next
End Function

Function AmRun(Ay, Fun$) As Variant()
Dim I: For Each I In Itr(Ay)
    PushI AmRun, Run(Fun, I)
Next
End Function

Function AmRunAsSy(Ay, Fun$) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRunAsSy, Run(Fun, I)
Next
End Function
Function AmAddPfx(Ay, Pfx) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmAddPfx, Pfx & I
Next
End Function
Function AmAddSfx(Ay, Sfx) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmAddSfx, I & Sfx
Next
End Function

Function AmAddPfxTab(Ay) As String(): AmAddPfxTab = AmAddPfx(Ay, vbTab): End Function

Function AmAddPfxSfx(Ay, Pfx, Sfx) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmAddPfxSfx, Pfx & I & Sfx
Next
End Function

Function AmStrSfx(Ay, Sfx) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmStrSfx, I & Sfx
Next
End Function

Function AmRev(Ay)
Dim O: O = Ay
Dim U&: U = UB(Ay)
Dim J&: For J = 0 To U
    O(J - U) = Ay(J)
Next
End Function

Function AmAft(Sy$(), Sep$) As String()
Dim S: For Each S In Itr(Sy)
    Push AmAft, Aft(S, Sep)
Next
End Function

Function AmAftRev(Sy$(), Sep$) As String()
Dim S: For Each S In Itr(Sy)
    Push AmAftRev, AftRev(S, Sep)
Next
End Function
Function AmAli(Ay, Optional W0%) As String()
Dim W%: If W0 <= 0 Then W = AyWdt(Ay) Else W = W0
Dim S: For Each S In Itr(Ay)
    PushI AmAli, AliL(S, W)
Next
End Function

Function AmAliR(Ay, Optional W0%) As String()
Dim W%: If W0 <= 0 Then W = AyWdt(Ay) Else W = W0
Dim I: For Each I In Itr(Ay)
    PushI AmAliR, AliR(I, W)
Next
End Function

Function AmBef(Sy$(), Sep$) As String()
Dim S: For Each S In Itr(Sy)
    Push AmBef, Bef(S, Sep)
Next
End Function
Function AmBefOrAll(Sy$(), Sep$) As String()
Dim S: For Each S In Itr(Sy)
    Push AmBefOrAll, BefOrAll(S, Sep)
Next
End Function
Function AmBefOrAllDash(Sy$()) As String(): AmBefOrAllDash = AmBefOrAll(Sy, "_"): End Function

Function AmInc(Nbry, Optional By = 1)
Dim O: O = Nbry
Dim J&: For J = 0 To UB(O)
    O(J) = O(J) + By
Next
AmInc = O
End Function

Function AmRmv2Dash(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRmv2Dash, RmvA2Hyp(I)
Next
End Function

Function AmRmvLasChr(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRmvLasChr, RmvLas(I)
Next
End Function

Function AmRmvPfx(Ay, Pfx$) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRmvPfx, RmvPfx(I, Pfx)
Next
End Function

Function AmRmvSfx(Ay, Sfx$) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRmvSfx, RmvSfx(I, Sfx)
Next
End Function

Function AmRplPfx(Ay, PfxFm$, PfxTo$, Optional C As eCas) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRplPfx, RplPfx(I, PfxFm, PfxTo, C)
Next
End Function

Function AmRpl(Ay, Fm$, By$, Optional Cnt& = 1, Optional C As eCas) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRpl, Replace(I, Fm, By, Count:=Cnt, Compare:=VbCprMth(C))
Next
End Function

Function AmRTrim(Sy$()) As String()
Dim S: For Each S In Itr(Sy)
    Push AmRTrim, RTrim(S)
Next
End Function

Function AmTab(Ay, Optional NTab% = 1) As String(): AmTab = AmAddPfx(Ay, TabN(NTab)):         End Function
Private Sub B_AmTrim():                                     DmpAy AmTrim(Sy(1, 2, 3, "  a")): End Sub
Function AmTrim(Ay) As String()
Dim S: For Each S In Itr(Ay)
    Push AmTrim, Trim(S)
Next
End Function
Function AmLTrim(Ay) As String()
Dim S: For Each S In Itr(Ay)
    Push AmLTrim, LTrim(S)
Next
End Function

Function AmRmvT1(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRmvT1, RmvA1T(I)
Next
End Function
Function AmRmv2T(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRmv2T, RmvA2T(I)
Next
End Function

Function AmRmvSngQuo(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRmvSngQuo, RmvQuoSng(I)
Next
End Function

Function AmRplStar(Ay, By) As String()
Dim I
For Each I In Itr(Ay)
    PushI AmRplStar, Replace(I, By, "*")
Next
End Function

Function AmRplT1(Ay, NewT1) As String()
AmRplT1 = AmAddPfx(AmRmvT1(Ay), NewT1 & " ")
End Function

Function AmTakBefDD(Sy$()) As String()
Dim I: For Each I In Itr(Sy)
    PushI AmTakBefDD, BefcHH(I)
Next
End Function

Function AmTakAftDot(Sy$()) As String(): AmTakAftDot = AmTakAft(Sy, "."): End Function

Function AmTakAft(Sy$(), Sep$) As String()
Dim I: For Each I In Itr(Sy)
    PushI AmTakAft, Aft(I, Sep)
Next
End Function

Function AmTakAftOrAll(Sy$(), Sep$) As String()
Dim I: For Each I In Itr(Sy)
    PushI AmTakAftOrAll, AftOrAll(I, Sep)
Next
End Function

Function AmTakBef(Sy$(), Sep$) As String() 'Return a Sy which is taking Bef-Sep from Given Sy
Dim I: For Each I In Itr(Sy)
    PushI AmTakBef, Bef(CStr(I), Sep)
Next
End Function

Function AmTakBefDot(Sy$()) As String(): AmTakBefDot = AmTakBef(Sy, "."):   End Function
Function AmAliQusq(Fny$()) As String():    AmAliQusq = AmAli(AmQuoSq(Fny)): End Function

Function AmTakBefOrAll(Sy$(), Sep$) As String()
Dim I: For Each I In Itr(Sy)
    Push AmTakBefOrAll, BefOrAll(CStr(I), Sep)
Next
End Function

Function AmBetBkt(Sy$()) As String()
Dim I: For Each I In Itr(Sy)
    PushI AmBetBkt, BetBkt(CStr(I))
Next
End Function

Function AmQuo(Ay, QuoStr$) As String()
If Si(Ay) = 0 Then Exit Function
Dim U&: U = UB(Ay)
Dim Q1$, Q2$
    With BrkQuo(QuoStr)
        Q1 = .S1
        Q2 = .S2
    End With

Dim O$()
    ReDim O(U)
    Dim J&
    For J = 0 To U
        O(J) = Q1 & Ay(J) & Q2
    Next
AmQuo = O
End Function

Function AmQuoDbl(Ay) As String(): AmQuoDbl = AmQuo(Ay, vbQuoDbl): End Function
Function AmQus(Sy$()) As String():    AmQus = AmQuo(Sy, vbQuoSng): End Function
Function AmQuoSq(Ay) As String():   AmQuoSq = AmQuo(Ay, "[]"):     End Function

Function AmRmvFstChr(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushI AmRmvFstChr, RmvFst(I)
Next
End Function

Function AmRmvFstNonLetter(Ay) As String() 'Gen:AyXXX
Dim I: For Each I In Itr(Ay)
    PushI AmRmvFstNonLetter, RmvFstNonLetter(I)
Next
End Function

Private Sub B_AmRplQ()
Dim Ay$(): Ay = SySs("Stm Bus L1 L2 L3 L4 Sku")
D AmRplQ(Ay, "#StkDay?")
End Sub

Function AmRplQ(Ay, QtpTbn$) As String()
Dim I: For Each I In Itr(Ay)
    PushS AmRplQ, RplQ(QtpTbn, I)
Next
End Function

Function AmRplMid(Ay, B As Bei, ByAy)
With Ay3FmAyBei(Ay, B)
AmRplMid = AyAddAp(.A, ByAy, .C)
End With
End Function
