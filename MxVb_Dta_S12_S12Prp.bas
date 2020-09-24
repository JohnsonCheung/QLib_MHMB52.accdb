Attribute VB_Name = "MxVb_Dta_S12_S12Prp"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_S12_Prp."

Function FstS2(S1, A() As S12) As Stropt
'Ret : Fnd S1 in A return S2 @@
Dim J&: For J = 0 To UbS12(A)
    With A(J)
        If .S1 = S1 Then FstS2 = SomStr(.S2): Exit Function
    End With
Next
End Function

Function IsS12Lines(A As S12) As Boolean
Select Case True
Case IsLines(A.S1), IsLines(A.S2): IsS12Lines = True: Exit Function
End Select
End Function

Function IsS12yLines(A() As S12) As Boolean
Dim J&: For J = 0 To UbS12(A)
    If IsS12Lines(A(J)) Then IsS12yLines = True: Exit Function
Next
End Function
Function S12yWhDif(A() As S12) As S12()
Dim J&: For J = 0 To UbS12(A)
    Dim M As S12: M = A(J)
    If IsS12Dif(M) Then PushS12 S12yWhDif, M
Next
End Function
Function IsS12Dif(S As S12) As Boolean: IsS12Dif = S.S1 <> S.S2: End Function
