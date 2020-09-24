Attribute VB_Name = "MxVb_Ay_Op_HasAy"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_Has."
Function HasObj(ObjAy, Obj) As Boolean
Dim OPtr&: OPtr = ObjPtr(Obj)
Dim I: For Each I In ObjAy
    If ObjPtr(I) = OPtr Then HasObj = True: Exit Function
Next
End Function

Function HasDup(Ay) As Boolean
Dim D: D = AyNw(Ay)
Dim I: For Each I In Ay
    If HasEle(D, I) Then HasDup = True: Exit Function
    PushI D, I
Next
End Function

Function HasEleStr(Ay, StrEle, Optional C As eCas) As Boolean
Dim I: For Each I In Itr(Ay)
    If IsEqStr(I, StrEle, C) Then HasEleStr = True: Exit Function
Next
End Function

Function HasEleRe(Ay, Re As RegExp) As Boolean
Dim Ele: For Each Ele In Itr(Ay)
    If Re.Test(Ele) Then HasEleRe = True: Exit Function
Next
End Function
Function NoEle(Ay, Ele) As Boolean: NoEle = Not HasEle(Ay, Ele): End Function

Function HasSubAy(Ay, AySub) As Boolean
Dim S: For Each S In Itr(AySub)
    If NoEle(Ay, S) Then Exit Function
Next
HasSubAy = True
End Function

Function IsInAy(V, Ay): IsInAy = HasEle(Ay, V): End Function
Function HasEle(Ay, Ele) As Boolean
Dim I: For Each I In Itr(Ay)
    If I = Ele Then HasEle = True: Exit Function
Next
End Function

Function HasEleAy(Ay, EleAy) As Boolean
Dim I
For Each I In Itr(EleAy)
    If Not HasEle(Ay, I) Then Exit Function
Next
HasEleAy = True
End Function

Function HasEleInSomAyOfAp(ParamArray AyAp()) As Boolean
Dim Ayav(): Ayav = AyAp
Dim Ay: For Each Ay In Itr(Ayav)
    If Si(Ay) > 0 Then HasEleInSomAyOfAp = True: Exit Function
Next
End Function

Function IsSubAy(AySub, SupAy) As Boolean
Dim I: For Each I In Itr(AySub)
    If Not HasEle(SupAy, I) Then Exit Function
Next
IsSubAy = True
End Function

Function IsSupAy(SupAy, AySub) As Boolean: IsSupAy = IsSubAy(AySub, SupAy): End Function

Private Sub B_ChkIsSupAy()
Dim SupAy, AySub
GoSub T1
Exit Sub
T1:
    SupAy = Array(1)
    AySub = Array(1, 2)
    GoTo Tst
Tst:
    ChkIsSupAy "ChkIsSupAy__Tst", SupAy, AySub
    Return
End Sub
Sub ChkIsSupAy(Fun$, SupAy, AySub)
If IsSupAy(SupAy, AySub) Then Exit Sub
Thw Fun, "SupAy error", "[ErEle in AySub] AySub SupAy", AyMinus(AySub, SupAy), AySub, SupAy
End Sub

Function HasEleAyInSeq(A, B) As Boolean
Dim BItm, Ix&
If Si(B) = 0 Then Stop
For Each BItm In B
    Ix = IxEle(A, BItm, Ix)
    If Ix = -1 Then Exit Function
    Ix = Ix + 1
Next
HasEleAyInSeq = True
End Function

Function HasDupEle(A) As Boolean
If Si(A) = 0 Then Exit Function
Dim Pool: Pool = A: Erase Pool
Dim I
For Each I In A
    If HasEle(Pool, I) Then HasDupEle = True: Exit Function
    Push Pool, I
Next
End Function

Function HasNegEle(A) As Boolean
Dim V
If Si(A) = 0 Then Exit Function
For Each V In A
    If V = -1 Then HasNegEle = True: Exit Function
Next
End Function

Private Sub B_HasEleAyInSeq()
Dim A, B
A = Array(1, 2, 3, 4, 5, 6, 7, 8)
B = Array(2, 4, 6)
Debug.Assert HasEleAyInSeq(A, B) = True
End Sub

Sub ChkHasEle(Ay, Ele, Fun$)
If NoEle(Ay, Ele) Then Thw Fun, "No @Ele in @Ay", "Ele Ay", Ele, Ay
End Sub
