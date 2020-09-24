Attribute VB_Name = "MxXls_Dta"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Dta."

Function DrsFxq(Fx, Q) As Drs: DrsFxq = DrsArs(CnFx(Fx).Execute(Q)): End Function

Private Sub B_DrsFxq()
Dmp WnyFx(MHO.MHOMB52.FxiSalTxt)
BrwDrs DrsFxw(MHO.MHOMB52.FxiSalTxt, "Sheet1")
End Sub
