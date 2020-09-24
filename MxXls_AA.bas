Attribute VB_Name = "MxXls_AA"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_AA."

Function DtFx(Fx$, Optional Wsn$) As Dt
Dim W$: W = WsnDft(Wsn, Fx)
DtFx = DtDrs(DrsFxw(Fx, W), W)
End Function
Function DrsFxw(Fx$, W$) As Drs: DrsFxw = DrsArs(ArsFxw(Fx, W)): End Function
