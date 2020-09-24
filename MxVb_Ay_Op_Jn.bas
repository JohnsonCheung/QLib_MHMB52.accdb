Attribute VB_Name = "MxVb_Ay_Op_Jn"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_Jn."

Function JnSpcApNB$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnSpcApNB = JnSpc(SyAyNB(Av))
End Function

Function JnVbarAp$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnVbarAp = JnVBar(Av)
End Function

Function JnVbarApSpc$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnVbarApSpc = JnVbarSpc(Av)
End Function

Function JnSpcAp$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnSpcAp = JnSpc(Av)
End Function

Function JnTabAp$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnTabAp = JnTab(Av)
End Function

Function JnSemiColonAp$(ParamArray Ap())
Dim Av(): If UBound(Ap) >= 0 Then Av = Ap
JnSemiColonAp = JnSemi(AeEmpEle(Av))
End Function

Function JnCmaSpcFf$(FF$)
JnCmaSpcFf = JnQSqCommaSpc(FnyFF(FF))
End Function
