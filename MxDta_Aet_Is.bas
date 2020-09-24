Attribute VB_Name = "MxDta_Aet_Is"
Option Compare Text
Const CMod$ = "MxDta_Aet_Is."
Option Explicit
Function IsAet(V) As Boolean
Select Case True
Case TypeName(V) <> "Dictionary"
Case Not IsItrAllEmp(CvDi(V).Items)
Case Else: IsAet = True
End Select
End Function
Function IsEmpAet(Aet As Dictionary) As Boolean
IsEmpAet = Aet.Count = 0
End Function

Function IsEqAet(Aet1 As Dictionary, Aet2 As Dictionary) As Boolean
If Aet1.Cnt <> Aet2.Cnt Then Exit Function
Dim K1: For Each K1 In Aet1.Keys
    If Not Aet2.Exists(K1) Then Exit Function
Next
IsEqAet = True
End Function

Function IsEqAetInOrd(Aet1 As Dictionary, Aet2 As Dictionary) As Boolean
If Aet1.Count <> Aet2.Count Then Exit Function
IsEqAetInOrd = IsEqAy(AvAet(Aet1), AvAet(Aet2))
End Function
