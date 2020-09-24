Attribute VB_Name = "MxVb_Ay_Op_AySrtQ"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_QSrt."
Enum eSrt: eSrtAsc: eSrtDes: End Enum
Private Sub B_AySrtQ()
Dim Ay, Ord As eSrt
GoSub T0
GoSub T1
Exit Sub
T0:
    Ay = Array(1, 2, 3, 4, 0, 1, 1, 5)
    Ord = eSrtAsc
    Ept = Array(0, 1, 1, 1, 2, 3, 4, 5)
    GoTo Tst
T1:
    Ay = Array(1, 2, 4, 87, 4, 2)
    Ord = eSrtDes
    Ept = Array(87, 4, 4, 2, 2, 1)
    GoTo Tst
Tst:
    Act = AySrtQ(Ay, Ord)
    C
    Return
End Sub

Function AySrtQ(Ay, Optional By As eSrt = eSrtAsc)
If Si(Ay) = 0 Then AySrtQ = Ay: Exit Function
AySrtQ = WSrt(Ay)
If By = eSrtDes Then AySrtQ = AyRev(AySrtQ)
End Function
Private Function WSrt(Ay)
Select Case Si(Ay)
Case 0, 1: WSrt = Ay
Case 2 And Ay(0) <= Ay(1)
    WSrt = Ay
Case 2
    Dim O: O = Ay
    Dim S: S = O(0)
    O(0) = O(1)
    O(1) = S
    WSrt = O
Case Else
    O = Ay
    Dim P: P = Pop(O)
    Dim L: L = AwLE(O, P)
    Dim H: H = AwGT(O, P)
    Dim A: A = WSrt(L)
    Dim B: B = WSrt(H)
    WSrt = A
    PushI WSrt, P
    PushIAy WSrt, B
End Select
End Function

Function SySrtQ(Sy$(), Optional By As eSrt = eSrtAsc, Optional C As eCas) As String()
If Si(Sy) = 0 Then Exit Function
Dim O$(): O = W2Srt(Sy, VbCprMth(C))
If By = eSrtDes Then O = AyRev(O)
SySrtQ = O
End Function
Private Function W2Srt(Sy$(), C As VbCompareMethod) As String()
Select Case Si(Sy)
Case 0, 1: W2Srt = Sy
Case 2 And StrComp(Sy(0), Sy(1), C) <> 1
    W2Srt = Sy
Case 2
    Dim O$(): O = Sy
    Dim S$: S = O(0)
    O(0) = O(1)
    O(1) = S
    W2Srt = O
Case Else
    O = Sy
    Dim P$: P = Pop(O)
    Dim L$(): L = W2AwLE(O, P, C)
    Dim H$(): H = W2AwGT(O, P, C)
    Dim A$(): A = W2Srt(L, C)
    Dim B$(): B = W2Srt(H, C)
    W2Srt = A
    PushI W2Srt, P
    PushIAy W2Srt, B
End Select
End Function
Private Function W2AwLE(Sy$(), P$, C As VbCompareMethod) As String()
Dim I: For Each I In Sy
    If StrComp(I, P, C) <= 0 Then PushI W2AwLE, I
Next
End Function
Private Function W2AwGT(Sy$(), P$, C As VbCompareMethod) As String()
Dim I: For Each I In Sy
    If StrComp(I, P, C) = 1 Then PushI W2AwGT, I
Next
End Function
