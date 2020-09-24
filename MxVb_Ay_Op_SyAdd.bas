Attribute VB_Name = "MxVb_Ay_Op_SyAdd"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_SyAdd."
Function LinesSapNB$(ParamArray SapNB()): Dim Av(): Av = SapNB: LinesSapNB = JnCrLf(AwNB(Av)): End Function
Function LinesSap$(ParamArray Sap()): Dim Av(): Av = Sap: LinesSap = JnCrLf(Av): End Function
Function SySapNB(ParamArray Sap()) As String(): Dim Av(): Av = Sap: SySapNB = AwNB(SyAy(Av)): End Function
Function SySap(S$, ParamArray Sap()) As String()
PushI SySap, S
Dim Av(): Av = Sap
Dim I: For Each I In Av
    PushI SySap, I
Next
End Function
Function SyAddAp(Sy$(), ParamArray ApSy()) As String()
SyAddAp = Sy
Dim Av(): Av = ApSy
Dim I: For Each I In ApSy
    If Not IsSy(I) Then Raise "Some of ele in @SyAp is not Sy, but [" & TypeName(I) & "]"
    PushIAy SyAddAp, I
Next
End Function

Function SySyEle(Sy$(), Str$) As String():   SySyEle = AyAyEle(Sy, Str):        End Function
Function SyStrSy(Str, Sy$()) As String():    SyStrSy = AyEleAy(Str, Sy):        End Function
Function LinesLnLy$(Ln$, Ly$()):           LinesLnLy = JnCrLf(SyStrSy(Ln, Ly)): End Function

Private Sub B_AyyGp()
GoSub T1
GoSub T2
Exit Sub
Dim Ay(), N%
T1:
    Ay = Array(1, 2, 3, 4, 5, 6)
    Ept = Array(Array(1, 2, 3, 4, 5), Array(6))
    N = 5
    GoTo Tst
T2:
    Ay = Array(1, 2, 3, 4, 5, 6)
    Ept = Array(Array(1, 2, 3, 4), Array(5, 6))
    N = 4
    GoTo Tst
Tst:
    Act = AyyGp(Ay, N%)
    C
    Return
End Sub
Function AyyGp(Ay, N%) As Variant()
Dim NEle&: NEle = Si(Ay): If NEle = 0 Then Exit Function
Dim Emp: Emp = Ay: Erase Emp
Dim M: M = Emp
Dim V, GpI%, Ix%: For Each V In Itr(Ay)
    PushI M, V
    GpI = GpI + 1
    If GpI = N Then
        GpI = 0
        PushI AyyGp, M
        M = Emp
    End If
Next
If Si(M) > 0 Then PushI AyyGp, M
End Function

Function AyAddAp(Ay, ParamArray Itm_or_AyAp())
Const CSub$ = CMod & "AyAddAp"
Dim Av(): Av = Itm_or_AyAp
If Not IsArray(Ay) Then Thw CSub, "Ay must be array", "Ay-TypeName", TypeName(Ay)
AyAddAp = Ay
Dim I: For Each I In Av
    If IsArray(I) Then
        PushIAy AyAddAp, I
    Else
        PushI AyAddAp, I
    End If
Next
End Function

Function AyMap(Ay, MapFun$) As Variant()
Dim I: For Each I In Itr(Ay)
    Push AyMap, Run(MapFun, I)
Next
End Function

Function AyAyEle(Ay, Ele)
AyAyEle = Ay
Push AyAyEle, Ele
End Function

Function AyEleAy(Ele, Ay)
Dim O: O = Ay: Erase O
Push O, Ele
PushAy O, Ay
AyEleAy = O
End Function

Function AvAyEle(Ay, Ele) As Variant(): AvAyEle = AyAyEle(Ay, Ele): End Function


Function AyInc(Ay, Optional N& = 1)
AyInc = AyNw(Ay)
Dim X: For Each X In Itr(Ay)
    PushI AyInc, X + N
Next
End Function

Private Sub B_AyAdd()
Dim Ay1(), Ay2()
GoSub T1
Exit Sub
T1:
    Ay1 = Array(1, 2, 2, 2, 4, 5)
    Ay2 = Array(2, 2)
    Ept = Array(1, 2, 2, 2, 4, 5, 2, 2)
    GoTo Tst
Tst:
    Act = AyAdd(Ay1, Ay2)
    C
    Return
End Sub

Private Sub B_AmAddPfx()
Dim Sy$(), Pfx$
GoSub T1
Exit Sub
T1:
    Sy = SySs("1 2 3 4")
    Pfx = "* "
    Ept = SyAp("* 1", "* 2", "* 3", "* 4")
    GoTo Tst
Tst:
    Act = AmAddPfx(Sy, Pfx)
    C
    Return
End Sub

Private Sub B_AmAddPfxSfx()
Dim Sy$(), Act$(), Sfx$, Pfx$, Exp$()
Sy = SyAp(1, 2, 3, 4)
Pfx = "* "
Sfx = "#"
Exp = SyAp("* 1#", "* 2#", "* 3#", "* 4#")
GoSub Tst
Exit Sub
Tst:
Act = AmAddPfxSfx(Sy, Pfx, Sfx)
Debug.Assert IsEqAy(Act, Exp)
Return
End Sub

Private Sub B_AmStrSfx()
Dim Sy$(), Sfx$
Sy = SySs("1 2 3 4")
Sfx = "#"
Ept = SySs("1# 2# 3# 4#")
GoSub Tst
Exit Sub
Tst:
    Act = AmStrSfx(Sy, Sfx)
    C
    Return
End Sub
