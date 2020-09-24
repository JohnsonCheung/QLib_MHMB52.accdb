Attribute VB_Name = "MxDta_Aet_FmTo_AetXxx"
Option Compare Text
Const CMod$ = "MxDta_Aet_Nw."
Option Explicit

Function AetAy(Ay, Optional C As eCas) As Dictionary
Set AetAy = AetEmp(C)
PushAetAy AetAy, Ay
End Function
Function AetNw(Optional C As eCas) As Dictionary
Set AetNw = New Dictionary
AetNw
End Function
Function AetItm(Itm) As Dictionary
Set AetItm = AetEmp
PushAetEle AetItm, Itm
End Function

Function AetClone(Aet As Dictionary, Optional C As eCas) As Dictionary
Set AetClone = New Dictionary
Dim K: For Each K In Aet.Keys
    AetClone.Add K, Empty
Next
End Function

Function AetItr(Itr) As Dictionary: Set AetItr = AetAetItr(AetEmp, Itr): End Function
Function AetSs(Ss$) As Dictionary:   Set AetSs = AetAy(SySs(Ss)):        End Function
Function AetEmp(Optional C As eCas) As Dictionary
Set AetEmp = New Dictionary
AetEmp.CompareMode = VbCprMth(C)
End Function
Function AetAp(ParamArray Ap()) As Dictionary
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
Set AetAp = AetAy(Av)
End Function
