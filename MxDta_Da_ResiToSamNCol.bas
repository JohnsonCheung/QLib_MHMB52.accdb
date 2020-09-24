Attribute VB_Name = "MxDta_Da_ResiToSamNCol"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Da_ResiToSamNCol."
Function ResiToSamNCol(A As Drs) As Drs
Dim N%: N = NDcDrs(A)
ResiToSamNCol = A
ResiToSamNCol.Dy = ResiToNCol(A.Dy, N - 1)
End Function

Function ResiToSamNDcDy(Dy()) As Variant()
If Si(Dy) = 0 Then Exit Function
ResiToSamNDcDy = ResiToNCol(Dy, NDcDy(Dy) - 1)
End Function

Function ResiToNCol(Dy(), U%) As Variant()
Dim Dr, O(), J&
O = Dy
For Each Dr In Itr(O)
    If UB(Dr) <> U Then
        ReDim Preserve Dr(U)
        O(J) = Dr
    End If
    J = J + 1
Next
ResiToNCol = O
End Function
