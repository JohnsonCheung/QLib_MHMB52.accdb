Attribute VB_Name = "MxVb_Ay_Op_AyRpl"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_Op_AyRpl."
Function SyRpl(Sy$(), SyBy$(), Bix) As String(): SyRpl = AyRpl(Sy, SyBy, Bix): End Function
Function AyRpl(Ay, AyBy, Bix)
Dim O: O = Ay
Dim J&: For J = Bix To UB(AyBy) + Bix
    O(J) = AyBy(J - Bix)
Next
AyRpl = O
End Function

Function SyRplBE(Sy$(), SyBy$(), Bix, Eix): SyRplBE = AyRplBE(Sy, SyBy, Bix, Eix): End Function
Private Sub B_AyRplBE()
GoSub T1
Exit Sub
Dim Ay, AyBy, Bix, Eix
T1:
    Ay = Array(0, 1, 2, 3, 4)
    AyBy = Array("a", "b", "c")
    Bix = 1
    Eix = 2
    Ept = Array(0, "a", "b", "c", 3, 4)
    GoTo Tst
Tst:
    Act = AyRplBE(Ay, AyBy, Bix, Eix)
    C
    Return
End Sub
Function AyRplBE(Ay, AyBy, Bix, Eix)
If Bix = 0 And Eix = -1 Then AyRplBE = Ay: Exit Function
Dim E: E = AwFm(Ay, Eix + 1) ' Sav the Las part
Dim O: O = Ay
ReDim Preserve O(Bix - 1)    ' Keep the fst part
PushIAy O, AyBy         ' add by snd part - @AyBy
PushIAy O, E            ' add hte las part *E
AyRplBE = O
End Function

Function SyRplBei(Sy$(), SyBy$(), M As Bei): SyRplBei = SyRplBE(Sy, SyBy, M.Bix, M.Eix): End Function
Function AyRplBei(Ay, AyBy, M As Bei):       AyRplBei = AyRplBE(Ay, AyBy, M.Bix, M.Eix): End Function
Function AyRplBeiy(Ay, AyBy, M() As Bei)
If SiBei(M) = 0 Then AyRplBeiy = AyAdd(Ay, AyBy): Exit Function
Dim Ay1: Ay1 = AwBef(Ay, M(0).Bix)
Dim AyRst
    Dim Ixy&(): Stop
    AyRst = AwIxy(Ay, Ixy)
AyRplBeiy = AyAddAp(Ay1, AyBy, AyRst)
End Function
Function SySetStr(Sy$(), Ix, Str$) As String(): SySetStr = AySetEle(Sy, Ix, Str): End Function
Function AySetEle(Ay, Ix, Ele)
Dim O: O = Ay
If IsBet(Ix, 0, UB(Ay)) Then
    O(Ix) = Ele
End If
AySetEle = O
End Function
