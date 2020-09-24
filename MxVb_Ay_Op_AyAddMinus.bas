Attribute VB_Name = "MxVb_Ay_Op_AyAddMinus"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_Op_AddMinus."

Private Sub B_AyMinusAp()
GoSub T1
GoSub T2
Exit Sub
Dim Ay
T1:
    Ay = Array(1, 2)
    Ept = Ay
    Act = AyMinusAp(Ay)
    C 1
    Return
T2:
    Ay = Array(1, 2, 3, 4)
    Ept = Array(1, 2)
    Act = AyMinusAp(Ay, Array(3, 4))
    C 2
    Return
End Sub
Function AyMinusAp(Ay, ParamArray AyAp())
Dim O: O = Ay
Dim IAy: For Each IAy In AyAp
    O = AyMinus(O, IAy)
    If Si(O) = 0 Then AyMinusAp = O: Exit Function
Next
AyMinusAp = O
End Function

Function AyItmAy(Itm, Ay)
Dim O: O = Ay: Erase O
PushI O, Itm
PushIAy O, Ay
AyItmAy = O
End Function
Function AyAyItm(Ay, Itm)
AyAyItm = Ay
PushI AyAyItm, Itm
End Function

Function AyRseq(Ay, AySub)
Dim AySam: AySam = AyIntersect(Ay, AySub)
Dim AyRst: AyRst = AyMinus(Ay, AySub)
AyRseq = AyAdd(AySam, AyRst)
End Function

Function AyAdd(AyA, AyB)
AyAdd = AyA
PushAy AyAdd, AyB
End Function

Function AvAdd(A(), B()) As Variant():      AvAdd = AyAdd(A, B):     End Function
Function SyAdd(SyA$(), SyB$()) As String(): SyAdd = AyAdd(SyA, SyB): End Function
