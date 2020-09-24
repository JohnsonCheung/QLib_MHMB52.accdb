Attribute VB_Name = "MxVb_Ay_Nw_PriyNw"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_PriyNw."
Function Inty(ParamArray Ap()) As Integer():    Dim Av(): Av = Ap:    Inty = IntoyAy(IntyEmp, Av):  End Function
Function BoolyAp(ParamArray Ap()) As Boolean(): Dim Av(): Av = Ap: BoolyAp = IntoyAy(BoolyEmp, Av): End Function
Function Lngy(ParamArray Ap()) As Long():       Dim Av(): Av = Ap:    Lngy = IntoyAy(LngyEmp, Av):  End Function
Function Sngy(ParamArray Ap()) As Single():     Dim Av(): Av = Ap:    Sngy = IntoyAy(Sngy, Av):     End Function
Function Dtey(ParamArray Ap()) As Date():       Dim Av(): Av = Ap:    Dtey = IntoyAy(Dtey, Av):     End Function
Function Dbly(ParamArray Ap()) As Double():     Dim Av(): Av = Ap:    Dbly = IntoyAy(Dbly, Av):     End Function
Function Ccyy(ParamArray Ap()) As Double():     Dim Av(): Av = Ap:    Ccyy = IntoyAy(Ccyy, Av):     End Function
Function IntoyAy(Intoy, Ay)
IntoyAy = AyNw(Intoy)
Dim I: For Each I In Itr(Ay)
    Push IntoyAy, I
Next
End Function
