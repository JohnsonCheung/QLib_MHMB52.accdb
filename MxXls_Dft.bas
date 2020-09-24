Attribute VB_Name = "MxXls_Dft"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Dft."
Function WsnDft$(Wsn$, Fx)
If Wsn = "" Then
    WsnDft = WsnFst(Fx)
Else
    WsnDft = Wsn
End If
End Function
