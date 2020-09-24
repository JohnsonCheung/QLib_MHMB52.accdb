Attribute VB_Name = "MxVb_Str_Brk_BrkBet"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Brk_BrkBet."
Type S3: A As String: B As String: C As String: End Type

Sub AsgBrkBet(S$, A$, B$, O1, O2, O3)
AsgS3 BrkBet(S, A, B), O1, O2, O3
End Sub

Sub AsgS3(A As S3, O1, O2, O3)
O1 = A.A
O2 = A.B
O3 = A.C
End Sub

Function BrkBet(S$, A$, B$) As S3
If S = "" Then Exit Function
Dim P1%, P2%, O As S3, LA%, LB%
LA = Len(A)
LB = Len(B)
P1 = InStr(S, A)
P2 = InStr(S, B)
Select Case True
Case P1 <> 0 And P2 <> 0 And P1 > P2:       Stop
Case P1 = 0 And P2 = 0: O.A = Trim(S)
Case P1 = 0:            O.A = Trim(Left(S, P2 - 1)): O.C = Trim(Mid(S, P2 + LB))
Case P2 = 0:            O.A = Trim(Left(S, P1 - 1)): O.B = Trim(Mid(S, P1 + LA))
Case Else:              O.A = Trim(Left(S, P1 - 1)): O.B = Trim(Mid(S, P1 + LA, P2 - P1 + LA - 2)): O.C = Trim(Mid(S, P2 + Len(LB)))
End Select
BrkBet = O
End Function
