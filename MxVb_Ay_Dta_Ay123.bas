Attribute VB_Name = "MxVb_Ay_Dta_Ay123"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_Ay123."
Type Ay3
    A As Variant
    B As Variant
    C As Variant
End Type

Function Ay3(A, B, C) As Ay3
Const CSub$ = CMod & "Ay3"
ChkIsAy A, CSub
ChkIsAy A, CSub
ChkIsAy B, CSub
ChkIsAy C, CSub
With Ay3
    .A = A
    .B = B
    .C = C
End With
End Function
Function Ay3FmAyBE(Ay, Bix&, Eix&) As Ay3
Dim O As Ay3
Ay3FmAyBE = Ay3( _
    AwBE(Ay, 0, Bix), _
    AwBE(Ay, Bix, Eix), _
    AwFm(Ay, Eix))
End Function

Private Sub B_Ay3FmAyBei()
Dim Ay(): Ay = Array(1, 2, 3, 4)
Dim M As Bei: M = Bei(1, 2)
Dim Act As Ay3: Act = Ay3FmAyBei(Ay, M)
Ass IsEqAy(Act.A, Array(1))
Ass IsEqAy(Act.B, Array(2, 3))
Ass IsEqAy(Act.C, Array(4))
End Sub

Private Sub B_Ay3FmAyBE()
Dim Ay(): Ay = Array(1, 2, 3, 4)
Dim Act As Ay3: Act = Ay3FmAyBE(Ay, 1, 2)
Ass IsEqAy(Act.A, Array(1))
Ass IsEqAy(Act.B, Array(2, 3))
Ass IsEqAy(Act.C, Array(4))
End Sub
