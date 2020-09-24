Attribute VB_Name = "MxVb_Dta_PosRRCC"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_PosRrcc."
Type Rrcc
    R1 As Long 'all started from 1
    R2 As Long
    C1 As Long
    C2 As Long
End Type
Type TRc
    R As Long
    C As Long
End Type
Function HasTRc(A As Rrcc, B As TRc) As Boolean
If IsNBet(B.R, A.R1, A.R2) Then Exit Function
If IsNBet(B.C, A.C1, A.C2) Then Exit Function
HasTRc = True
End Function
Function Rrcc(R1, R2, C1, C2) As Rrcc
If R1 < 0 Then Exit Function
If R2 < 0 Then Exit Function
If C1 < 0 Then Exit Function
If C2 < 0 Then Exit Function
With Rrcc
    .R1 = R1
    .R2 = R2
    .C1 = C1
    .C2 = C2
End With
End Function

Function IsEqRrcc(A As Rrcc, B As Rrcc) As Boolean
Dim A1 As Rrcc: A1 = RrccNrm(A)
Dim B1 As Rrcc: B1 = RrccNrm(B)
If A1.R1 <> B1.R1 Then Exit Function
If A1.R2 <> B1.R2 Then Exit Function
If A1.C1 <> B1.C1 Then Exit Function
If A1.C2 <> B1.C2 Then Exit Function
IsEqRrcc = True
End Function

Function RrccNrm(A As Rrcc) As Rrcc
Dim O As Rrcc
With O
    If A.R1 < 0 Then .R1 = 0
    If A.R2 < 0 Then .R2 = 0
    If A.C1 < 0 Then .C1 = 0
    If A.C2 < 0 Then .C2 = 0
    If .R1 > .R2 Then .R1 = 0: .R2 = 0
End With
End Function

Function IsEmpRrcc(A As Rrcc) As Boolean
IsEmpRrcc = IsEqRrcc(A, RrccEmp)
End Function

Function RrccEmp() As Rrcc: End Function

Function StrRrcc$(A As Rrcc)
With A
StrRrcc = FmtQQ("Rrcc(? ? ? ?)", .R1, .R2, .C1, .C2)
End With
End Function

Function RrccSq(Sq()) As Rrcc
With RrccSq
    .R1 = LBound(Sq, 1)
    .R2 = UBound(Sq, 1)
    .C1 = LBound(Sq, 2)
    .C2 = UBound(Sq, 2)
End With
End Function
