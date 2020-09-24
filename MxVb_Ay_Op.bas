Attribute VB_Name = "MxVb_Ay_Op"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_Op."
Public Const StmtBrkPatn$ = "(\.  |\r\n|\r)"
#If Doc Then
'CCC:Cml :S #(C)oln-(S)eparated-by-(S)ngle-(S)pace#.  Assume no space in front and at end each column name is separated by one space.
'  Coumn name array is obtained by !SplitSPc.
'Am:Cml :Fun-Pfx #(A)y-(m)ap ! Given @Ay will return same number of ele after doing some mapping
'Ay:Cml :Array #Array# ele can be object
'IAy:Cml :NonObj-Array) #Item-Ay# its ele cannot be object
'Sy:Cml :String-Array #String-Array#
'SS:Cml :Ln #(S)pc-(S)eparated# it is a list of no-space-term can have spaces in front or at end each item is separated by one or more space
'TmlAy:Cml :Ln #Tm(Tm)-Line(l)# It is a list of term can have spaces in front or at end.
                              ' Each item is separated by one or more space.
                              ' If the term has space, the term is quoted by []
'Ss:Cml :Ln #Single-Space-(S)eparated-Str.  It will be splitted by !SplitSSS.  Any 3 same letters will its owner letter meanings as in Ss format.  Example FfF where F means fldn and FfF means Fldn-Ss.
'FfF:Cml :Ss #Ss-Fldn#
'SnoInt: :Inty #Int-Sequence# ! Each next element is always 1 more than previous one
':LngSeg: :Lngy #Int-Sequence# ! Each next element is always 1 more than previous one
#End If

Sub AsgAy(Ay, ParamArray OAp())
Dim OAv(): OAv = OAp
Dim J%: For J = 0 To Min(UB(Ay), UB(OAv))
    OAp(J) = Ay(J)
Next
End Sub

Private Sub B_MsglDupEle()
Dim Ay
Ay = Array("1", "1", "2")
Ept = Sy("This item[1] is duplicated")
GoSub Tst
Exit Sub
Tst:
    Act = MsglDupEle(Ay, "This item[?] is duplicated")
    C
    Return
End Sub
Function MsglDupEle$(Ay, QMsg$)
Dim Dup: Dup = AwDup(Ay)
If Si(Dup) = 0 Then Exit Function
MsglDupEle = FmtQQ(QMsg, JnSpc(Dup))
End Function

Sub ChkNEmpAy(Ay): ThwTrue IsEmpAy(Ay), "ChkNEmpAy", "Given array is empty": End Sub

Function EleCnt%(Ay, M)
If Si(Ay) = 0 Then Exit Function
Dim O%, X
For Each X In Itr(Ay)
    If X = M Then O = O + 1
Next
EleCnt = O
End Function

Sub ResiHigh(OAy1, OAy2)
Dim U1&, U2&: U1 = UB(OAy1): U2 = UB(OAy2)
Select Case True
Case U1 > U2: ReDim Preserve OAy2(U1)
Case U2 > U1: ReDim Preserve OAy1(U2)
End Select
End Sub
Sub ResiLow(OAy1, OAy2)
Dim U1&, U2&: U1 = UB(OAy1): U2 = UB(OAy2)
Select Case True
Case U2 > U1: ReDim Preserve OAy2(U1)
Case U1 > U2: ReDim Preserve OAy1(U2)
End Select
End Sub
Function AyNw(Ay): AyNw = Ay: Erase AyNw: End Function

Function AyRev(IAy) ' rev an @IAy
Dim O: O = IAy
Dim U&: U = UB(O)
Dim J&: For J = 0 To U
    O(J) = IAy(U - J)
Next
AyRev = O
End Function

Function AyWdt%(Ay)
Dim O%, V: For Each V In Itr(Ay)
    O = Max(O, Len(V))
Next
AyWdt = O
End Function

Private Sub B_AyAsgAy()
Dim O%, Ay$
AsgAy Array(234, "abc"), O, Ay
Ass O = 234
Ass Ay = "abc"
End Sub


Private Sub B_EleMax()
Dim Ay()
Dim Act
Act = EleMax(Ay)
Stop
End Sub

Private Sub B_HasDupEle()
Ass HasDupEle(Array(1, 2, 3, 4)) = False
Ass HasDupEle(Array(1, 2, 3, 4, 4)) = True
End Sub

Private Sub B_AyInsAy()
Dim Act, Exp, Ay(), B(), At&
Ay = Array(1, 2, 3, 4)
B = Array("X", "Z")
At = 1
Exp = Array(1, "X", "Z", 2, 3, 4)

Act = AyInsAy(Ay, B, At)
Ass IsEqAy(Act, Exp)
End Sub


Private Sub B_KKCMiDy()
Dim Dy(), Act As KKCntMulItmColDy, KKColIx%(), IxEle%
PushI Dy, Array()
PushI Dy, Array()
PushI Dy, Array()
PushI Dy, Array()
PushI Dy, Array()
PushI Dy, Array()
'Ass Si(Act) = 4
'Ass IsEqAy(Act(0), Array("Ay", 3, 1, 2, 3))
'Ass IsEqAy(Act(1), Array("B", 3, 2, 3, 4))
'Ass IsEqAy(Act(2), Array("C", 0))
'Ass IsEqAy(Act(3), Array("D", 1, "X"))
End Sub

Function ItrSS(Ss): Asg Itr(SySs(Ss)), ItrSS: End Function
Function SsSrt$(Ss$): SsSrt = JnSpc(AySrt(SySs(Ss))): End Function

Function Cnoy(Cny$(), CC$) As Integer() ' Column no array of @CC according @Cny.  Cno starts from 1
Const CSub$ = CMod & "Cnoy"
Dim C: For Each C In ItrTml(CC)
    Dim Ix&: Ix = IxEle(Cny, C): If Ix = -1 Then Thw CSub, "A Coln in @CC not in @Cny", "C @CC @Cny", C, CC, Cny
    PushI Cnoy, Ix + 1
Next
End Function

Function IsEqSy(A$(), B$()) As Boolean
If Not IsEqSi(A, B) Then Exit Function
IsEqSy = StrComp(JnCrLf(A), JnCrLf(B), vbBinaryCompare) = 0
End Function

Function IsEqDr(A, B) As Boolean
Dim X, J&
For Each X In Itr(A)
    If X <> B(J) Then Exit Function
    J = J + 1
Next
IsEqDr = True
End Function

Sub ChkAyOrdered(Ay, Optional Fun$ = "ChkAyOrdered")
Dim J&: For J = 0 To UB(Ay) - 1
    If Ay(J) > Ay(J + 1) Then
        Dim Msg$: Msg = FmtQQ("Ele [?] and [?] are not in ascending ordere", J, J + 1)
        Dim NN$:  NN = FmtQQ("?-th-Ele ?-th-Ele Ay", J, J + 1)
        Thw Fun, Msg, NN, J, J + 1, AmAddIxPfx(Ay, 0)
    End If
Next
End Sub
Function IsEqAyFstN(A, B, FstN%) As Boolean
Dim J%: For J = 0 To FstN - 1
    If A(J) <> B(J) Then Exit Function
Next
IsEqAyFstN = True
End Function

Function IsEqAy(A, B) As Boolean
If Not IsArray(A) Then Exit Function
If Not IsArray(B) Then Exit Function
If Not IsEqSi(A, B) Then Exit Function
Dim J&, X
For Each X In Itr(A)
    If Not IsEqV(X, B(J)) Then Exit Function
    J = J + 1
Next
IsEqAy = True
End Function

Function AyAyy(AyOfAy())
If Si(AyOfAy) = 0 Then AyAyy = AvEmp: Exit Function
AyAyy = AyOfAy(0)
Dim J&: For J = 1 To UB(AyOfAy)
    PushAy AyAyy, AyOfAy(J)
Next
End Function


Function LikyKssy(Kssy$()) As String()
Dim Kss: For Each Kss In Itr(Kssy)
    PushIAy LikyKssy, Kss
Next
End Function
Function CvVy(Vy)
Const CSub$ = CMod & "CvVy"
Select Case True
Case IsStr(Vy): CvVy = SySs(CStr(Vy))
Case IsArray(Vy): CvVy = Vy
Case Else: Thw CSub, "VyzDicKK should either be string or array", "Vy-TypeName Vy", TypeName(Vy), Vy
End Select
End Function

Sub ChkNoDup(Ay, Optional N$ = "Ay", Optional Fun$ = "ChkNoDup")
' If there are 2 ele with same string (IgnCas), throw error
Dim Dup$()
    Dup = AwDup(Ay)
If Si(Dup) = 0 Then Exit Sub
Thw Fun, "There are dup in array", "AyNm Dup Ay", N, Dup, Ay
End Sub

Function MinEleGT0(Ay)
If Si(Ay) = 0 Then Exit Function
Dim O: O = Ay(0)
Dim V: For Each V In Ay
    If V > 0 Then
        If O = 0 Then
            O = V
        Else
            If V < O Then O = V
        End If
    End If
Next
MinEleGT0 = O
End Function

Function AySum#(AyNum)
Dim O#, V: For Each V In Itr(AyNum)
    O = O + V
Next
AySum = O
End Function

Function HasIntersect(A, B) As Boolean
Dim I: For Each I In Itr(A)
    If HasEle(B, I) Then HasIntersect = True: Exit Function
Next
End Function

Function SyIntersect(A$(), B$()) As String(): SyIntersect = AyIntersect(A, B): End Function
Function AyIntersect(A, B)
AyIntersect = AyNw(A)
If Si(A) = 0 Then Exit Function
If Si(A) = 0 Then Exit Function
Dim V
For Each V In A
    If HasEle(B, V) Then PushI AyIntersect, V
Next
End Function

Function SyMinus(A$(), B$()) As String()
SyMinus = AyMinus(A, B)
End Function

Private Sub B_AyMinus()
GoSub T1
Dim A(), B()
T1:
    A = Array(1, 2, 2, 2, 4, 5)
    B = Array(2, 2)
    Ept = Array(1, 2, 4, 5)
    GoTo Tst
T2:

    A = Array(1, 4, 5)
    B = Array(2, 1)
    Ept = Array(4, 5)
    GoTo Tst
Tst:
    Act = AyMinus(A, B)
    C
    Return
End Sub
Function AyMinus(A, B)
If Si(B) = 0 Then AyMinus = A: Exit Function
AyMinus = AyNw(A)
If Si(A) = 0 Then Exit Function
Dim V
For Each V In A
    If Not HasEle(B, V) Then
        PushI AyMinus, V
    End If
Next
End Function
