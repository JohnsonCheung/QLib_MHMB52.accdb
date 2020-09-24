Attribute VB_Name = "MxVb_Dta_Cnt256"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Cnt256."
Public Const Cnt256Ff$ = "Asc Chr Cnt"
Function Cnt256(S) As Long()
':Cnt256: :Lngy-with-256-ele
Dim O&(): ReDim O(255)
Dim J&: For J = 1 To Len(S)
    Dim A As Byte: A = Asc(Mid(S, J, 1))
    O(A) = O(A) + 1
Next
Cnt256 = O
End Function

Function Cnt256Drs(Cnt256&()) As Drs
Cnt256Drs = DrsFf(Cnt256Ff, Cnt256Dy(Cnt256))
End Function

Function Cnt256Dy(Cnt256&()) As Variant()
Dim J%: For J = 0 To 255
    If Cnt256(J) > 0 Then
        PushI Cnt256Dy, Array(J, Chr(J), Cnt256(J))
    End If
Next
End Function
