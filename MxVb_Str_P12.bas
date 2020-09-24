Attribute VB_Name = "MxVb_Str_P12"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_P12_S."
Type P12: P1 As Integer: P2 As Integer: End Type 'Deriving(Ay Ctor)
Function P12SI&(P() As P12): On Error Resume Next: P12SI = UBound(P) + 1: End Function
Function UbP12&(P() As P12): UbP12 = P12SI(P) - 1: End Function
Function TakPosIf$(S, Pos%)
If Pos > 0 Then TakPosIf = Mid(S, Pos)
End Function
Function TakP12y(S, P() As P12) As String()
Dim J%: For J = 0 To UbP12(P)
    PushI TakP12y, BetP12(S, P(J))
Next
End Function
Function BetP12y(S, P() As P12, Optional Inl As Boolean) As String()
Dim J&: For J = 0 To UbP12(P)
    PushI BetP12y, BetP12(S, P(J), Inl)
Next
End Function
Sub PushP12(O() As P12, M As P12): Dim N&: N = P12SI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function P12yRepy(S, P() As P12) As String()
Dim J&: For J = 0 To UbP12(P)
    PushI P12yRepy, P12Rep(P(J))
Next
End Function
Function IsEqP12(A As P12, B As P12) As Boolean
With A
Select Case True
Case .P1 <> B.P1, .P2 <> B.P2: Exit Function
End Select
End With
IsEqP12 = True
End Function
Function IsEqP12y(A() As P12, B() As P12) As Boolean
Dim U&: U = UbP12(A): If U <> UbP12(B) Then Exit Function
Dim J&: For J = 0 To U
    If Not IsEqP12(A(J), B(J)) Then Exit Function
Next
IsEqP12y = True
End Function
Function LenP12%(P As P12, Optional Inl As Boolean)
If IsEmpP12(P) Then Exit Function
With P
    If Inl Then
        LenP12 = .P2 - .P1 + 1
    Else
        LenP12 = .P2 - .P1 - 1
    End If
End With
End Function
Function IsEmpP12(P As P12) As Boolean
With P
Select Case True
Case .P1 <= 0, .P2 <= 0, .P1 > .P2: IsEmpP12 = True: Exit Function
End Select
End With
End Function
Function P12(P1, P2) As P12
Const CSub$ = CMod & "P12"
If P1 <= 0 Then ThwPm CSub, "P1 Must >=1", "P1", P1
If P2 <= 0 Then ThwPm CSub, "P2 Must >=1", "P2", P2
If P1 > P2 Then ThwPm CSub, "P2 Must >= P1", "P1 P2", P1, P2
With P12
    .P1 = P1
    .P2 = P2
End With
End Function
Function EmpP12() As P12: End Function
Function P12Rep$(P As P12): P12Rep = "P12 " & P.P1 & " " & P.P2: End Function
Function RepP12(P12Rep$) As P12
Dim M$
Dim P$(): P = SySs(P12Rep): If Si(P) <> 3 Then M = "Should be 3 terms": GoTo M
If P(0) <> "P12" Then M = "Term1 <> 'P12'": GoTo M
If CanCvInt(P(1)) Then M = "P1 is not Int": GoTo M
If CanCvInt(P(2)) Then M = "P2 is not Int": GoTo M
Dim P1%: P1 = P(1)
Dim P2%: P2 = P(2)
If P1 > P2 Then M = "P2 > P1": GoTo M
If P1 < -1 Then M = "P1 < -1": GoTo M
If P2 < -1 Then M = "P2 < -1": GoTo M
If P1 = -1 And P2 <> -1 Then M = "P1 = -1 and P2 <> -1": GoTo M
If P2 = -1 And P1 <> -1 Then M = "P2 = -1 and P1 <> -1": GoTo M
RepP12.P1 = P1
RepP12.P2 = P2
Exit Function
M: ThwPm CSub, M, "@Rep12", P12Rep
End Function
