Attribute VB_Name = "MxDta_Aet"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Aet."
Function CvAet(V) As Dictionary: Set CvAet = V: End Function

Function AetAetItr(Aet As Dictionary, Itr) As Dictionary
Set AetAetItr = AetClone(Aet)
PushAetItr AetAetItr, Itr
End Function

Function AetSrt(Aet As Dictionary) As Dictionary: Set AetSrt = AetAy(AySrtQ(AvAet(Aet))): End Function

Function AetAdd(Aet1 As Dictionary, Aet2 As Dictionary) As Dictionary
Set AetAdd = AetClone(Aet1)
Dim K: For Each K In Aet2.Keys
    PushAetEle AetAdd, K
Next
End Function

Function AetMinus(Aet1 As Dictionary, Aet2 As Dictionary) As Dictionary
Set AetMinus = New Dictionary
Dim E1: For Each E1 In Aet1.Keys
    If Not Aet2.Exists(E1) Then PushAetEle AetMinus, E1
Next
End Function
