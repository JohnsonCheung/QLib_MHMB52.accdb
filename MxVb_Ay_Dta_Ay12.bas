Attribute VB_Name = "MxVb_Ay_Dta_Ay12"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_Ay12."
Type Ay2
    A As Variant
    B As Variant
End Type

Function Ay2(A, B) As Ay2
Const CSub$ = CMod & "Ay2"
ChkIsAy A, CSub
ChkIsAy B, CSub
With Ay2
    .A = A
    .B = B
End With
End Function

Function Ay2AyPfx(Ay, Pfx$) As Ay2
Dim O As Ay2
O.A = AyNw(Ay)
O.B = O.A
Dim S$, I
For Each I In Itr(Ay)
    S = I
    If HasPfx(S, Pfx) Then
        PushI O.B, S
    Else
        PushI O.A, S
    End If
Next
Ay2AyPfx = O
End Function

Function Ay2AyN(Ay, N&) As Ay2
Ay2AyN = Ay2(AwFstN(Ay, N), AeFstN(Ay, N))
End Function

Function Ay2Jn(A, B, Sep$) As String()
Dim J&: For J = 0 To Min(UB(A), UB(B))
    PushI Ay2Jn, A(J) & Sep & B(J)
Next
End Function

Function Ay2JnDot(A, B) As String()
Ay2JnDot = Ay2Jn(A, B, ".")
End Function

Function Ay2JnSngQ(A, B) As String()
Ay2JnSngQ = Ay2Jn(A, B, "'")
End Function

Function Ay3FmAyBei(Ay, B As Bei) As Ay3
Ay3FmAyBei = Ay3FmAyBE(Ay, B.Bix, B.Eix)
End Function

Function DyTWOAY(A, B) As Variant()
Dim J&: For J = 0 To Min(UB(A), UB(B))
    PushI DyTWOAY, Array(A(J), B(J))
Next
End Function

Function DrsTWOAY(A, B, Optional N1$ = "Ay1", Optional N2$ = "Ay2") As Drs
DrsTWOAY = Drs(Sy(N1, N2), DyTWOAY(A, B))
End Function

Function FmtAyabSpc(AyA, AyB) As String()
FmtAyabSpc = FmtAy12(AyA, AyB, " ")
End Function

Function FmtAy12(A, B, Optional Sep$, Optional FF$ = "Ay1 Ay2") As String()
FmtAy12 = FmtS12y(S12yAy12(A, B), FF)
End Function

Function JnAy12(A, B, Optional SepAB$ = " = ", Optional SepItm$ = vbCmaSpc)
Dim O$()
Dim J%: For J = 0 To UB(A)
    PushI O, A(J) & SepAB & B(J)
Next
JnAy12 = Jn(O, SepItm)
End Function

Function FmtAyabNEmpB(A, B, Optional Sep$ = " ") As String()
Dim J&, O$()
For J = 0 To UB(A)
    If Not IsEmp(B(J)) Then
        Push O, A(J) & Sep & B(J)
    End If
Next
FmtAyabNEmpB = O
End Function

Sub AsgAyaReSzMax(A, B, OA, OB)
OA = A
OB = B
ResiHigh OA, OB
End Sub

Function DyAy2(A, B) As Variant()
Const CSub$ = CMod & "DyAy2"
ChkIsEqAySi A, B, CSub
Dim I, J&: For Each I In Itr(A)
    PushI DyAy2, Array(I, B(J))
    J = J + 1
Next
End Function
