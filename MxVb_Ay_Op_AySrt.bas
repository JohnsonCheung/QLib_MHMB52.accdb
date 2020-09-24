Attribute VB_Name = "MxVb_Ay_Op_AySrt"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_Srt."
Function IxySrtAy(Ay, Optional Ord As eSrt) As Long() ' ret Ixy so thatf @Ay(@@Ixy) is sorted
If Si(Ay) = 0 Then Exit Function
Dim Ix&, V, J&
Dim O&():
Push O, 0
For J = 1 To UB(Ay)
    O = AyIns(O, J, W2Ix(O, Ay, Ay(J)))
Next
If Ord = eSrtDes Then O = AyRev(O)
IxySrtAy = O
End Function
Private Function W2Ix&(Ix&(), Ay, V) ' ret the ix
Dim I, O&: For Each I In Ix
    If V < Ay(I) Then GoTo X
    O = O + 1
Next
X:
W2Ix = O
End Function

Function LinesSrt$(Lines$): LinesSrt = JnCrLf(AySrt(SplitCrLf(Lines))): End Function
Function IsAyNbr(Ay) As Boolean
If IsEmpAy(Ay) Then Exit Function
Dim I: For Each I In Itr(Ay)
    If IsStr(I) Then Exit Function
    If Not IsNumeric(I) Then Exit Function
Next
IsAyNbr = True
End Function

Function IsAySrt(Ay) As Boolean
Dim J&: For J = 0 To UB(Ay) - 1
   If Ay(J) > Ay(J + 1) Then Exit Function
Next
IsAySrt = True
End Function

Private Sub B_AySrtByAy()
Dim Ay, ByAy
Ay = Array(1, 2, 3, 4)
ByAy = Array(3, 4)
Ept = Array(3, 4, 1, 2)
GoSub Tst
Exit Sub
Tst:
    Act = AySrtByAy(Ay, ByAy)
    C
    Return
End Sub

Function AySrtByAy(Ay, ByAy)
Dim O: O = AyNw(Ay)
Dim I
For Each I In ByAy
    If HasEle(Ay, I) Then PushI O, I
Next
PushIAy O, AyMinus(Ay, O)
AySrtByAy = O
End Function

Function AyTab(Ay) As String()
Dim I: For Each I In Itr(Ay)
    PushS AyTab, vbTab & I
Next
End Function

Function SySrtLen(Sy$(), Optional Ord As eSrt) As String()
If IsEmpAy(Sy) Then Exit Function
Dim AyLen&(): ReDim AyLen(UB(Sy))
Dim J&: For J = 0 To UB(Sy)
    AyLen(J) = Len(Sy(J))
Next
Dim Ixy&(): Ixy = IxySrtAy(AyLen, Ord)
SySrtLen = AwIxy(Sy, Ixy)
End Function


Private Sub B_AySrt()
GoSub T1
'GoSub T2
'GoSub T3
Exit Sub
Dim A()
T1:
    A = Array(1, 2, 3, 4, 5): Ept = A:
    GoTo Tst
T2:
    A = Array(":", "~", "P"): Ept = Array(":", "P", "~"):
    GoTo Tst
T3:
    Erase A
    Push A, ":PjUpdTm:Sub"
    Push A, ":MthBrk:Function"
    Push A, "~~:Tst:Sub"
    Push A, ":PjTmNy_WithEr:Function"
    Push A, "~Private:JnContinueLin:Sub"
    Push A, "Private:HasPfx:Function"
    Push A, "Private:MdMthDRsFunBdyLy:Function"
    Push A, "Private:SrcMthLx_ToLx:Function"
    Erase Ept
    Push Ept, ":PjTmNy_WithEr:Function"
    Push Ept, ":PjUpdTm:Sub"
    Push Ept, ":MthBrk:Function"
    Push Ept, "Private:HasPfx:Function"
    Push Ept, "Private:MdMthDRsFunBdyLy:Function"
    Push Ept, "Private:SrcMthLx_ToLx:Function"
    Push Ept, "~Private:JnContinueLin:Sub"
    Push Ept, "~~:Tst:Sub"
    GoTo Tst
Tst:
    Act = AySrtQ(A)
    C
    Act = AySrt(A)
    C
    Return
End Sub
Function AySrt(Ay, Optional By As eSrt)
If IsEmpAy(Ay) Then AySrt = Ay: Exit Function
Dim Ix&
Dim O: O = Ay: Erase O
PushI O, Ay(0)
Dim J&: For J = 1 To UB(Ay)
    WSetSrcIns O, Ay(J)
Next
If By = eSrtDes Then O = AyRev(O)
AySrt = O
End Function
Private Sub WSetSrcIns(OAy, Ele)
Dim IxAt&: IxAt = WIxAt(OAy, Ele)
Dim N&: N = Si(OAy)
ReDim Preserve OAy(N)
WMov OAy, N, IxAt
OAy(IxAt) = Ele
End Sub
Private Function WIxAt&(Ay, Ele) ' ret an ix of @SrtdAy, so that @V should insert be that ix
Dim O&: For O = 0 To UB(Ay)
    If Ay(O) >= Ele Then WIxAt = O: Exit Function
Next
WIxAt = O
End Function
Private Sub WMov(OAy, U&, IxAt&)
Dim IxTo&: For IxTo = U To IxAt + 1 Step -1
    OAy(IxTo) = OAy(IxTo - 1)
Next
End Sub


Private Sub B_IxySrtAy()
GoSub T1
Exit Sub
Dim A, Ord As eSrt
T1:
    A = Array("A", "B", "C", "D", "E")
    Ept = Array(0, 1, 2, 3, 4)
    Ord = eSrtAsc
    GoTo Tst
T2:
    A = Array("A", "B", "C", "D", "E")
    Ept = Array(4, 3, 2, 1, 0)
    Ord = eSrtDes
    GoTo Tst
T4:
    '-----------------
    Erase A
    Push A, ":PjUpdTm:Sub"
    Push A, ":MthBrk:Function"
    Push A, "~~:Tst:Sub"
    Push A, ":PjTmNy_WithEr:Function"
    Push A, "~Private:JnContinueLin:Sub"
    Push A, "Private:HasPfx:Function"
    Push A, "Private:MdMthDRsFunBdyLy:Function"
    Push A, "Private:SrcMthLx_ToLx:Function"
    Ept = SyEmp
    Push Ept, ":PjTmNy_WithEr:Function"
    Push Ept, ":PjUpdTm:Sub"
    Push Ept, ":MthBrk:Function"
    Push Ept, "Private:HasPfx:Function"
    Push Ept, "Private:MdMthDRsFunBdyLy:Function"
    Push Ept, "Private:SrcMthLx_ToLx:Function"
    Push Ept, "~Private:JnContinueLin:Sub"
    Push Ept, "~~:Tst:Sub"
    Act = AySrt(A)
    ChkEq Ept, Act
Tst:
    Act = IxySrtAy(A, Ord)
    C
    Return
End Sub

Function SrtAyInEIxIxy&(Ix&(), A, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In Ix
        If V > A(I) Then SrtAyInEIxIxy& = O: Exit Function
        O = O + 1
    Next
    SrtAyInEIxIxy& = O
    Exit Function
End If
For Each I In Ix
    If V < A(I) Then SrtAyInEIxIxy& = O: Exit Function
    O = O + 1
Next
SrtAyInEIxIxy& = O
End Function
