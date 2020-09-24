Attribute VB_Name = "MxVb_Ay_Nw_PrimyAp"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_PrimAyFmAp."

Function SyApNB(ParamArray ApItmOrAy()) As String()
If UBound(ApItmOrAy) = 0 Then Exit Function
Dim Av(): Av = ApItmOrAy
SyApNB = SyAvNB(Av)
End Function

Function Sy(ParamArray ApItmOrAy()) As String()
If UBound(ApItmOrAy) = -1 Then Exit Function
Dim Av(): Av = ApItmOrAy
Sy = SyAv(Av)
End Function

Function SyNB(ParamArray ApItmOrAy()) As String()
Dim Av(): Av = ApItmOrAy
SyNB = SyAvNB(Av)
End Function

Function SyAvNB(AvItmOrAy()) As String()
Dim I: For Each I In Itr(AvItmOrAy)
    If IsArray(I) Then
        PushNBAy SyAvNB, I
    Else
        PushNB SyAvNB, I
    End If
Next
End Function

Function SyAp(ParamArray ApItmOrAy()) As String()
Dim Av(): Av = ApItmOrAy
SyAp = SyAv(Av)
End Function
Function SyAv(AvOf_Itm_or_Ay) As String()
Dim I: For Each I In Itr(AvOf_Itm_or_Ay)
    If IsArray(I) Then
        PushIAy SyAv, I
    Else
        PushI SyAv, I
    End If
Next
End Function

Function Av(ParamArray Ap()) As Variant()
Av = Ap
End Function
