Attribute VB_Name = "MxXls_Op_Opn"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Op_Opn."

Sub OpnFx(Fx$):     OpnFxX XlsNw, Fx:   End Sub
Sub OpnFxy(Fxy$()): OpnFxyX Fxy, XlsNw: End Sub
Sub OpnFxAp(ParamArray ApFx())
Dim Av(): Av = ApFx
OpnFxy SyAy(Av)
End Sub
Sub OpnFxyX(Fxy$(), X As Excel.Application)
Minvn X
Dim F: For Each F In Fxy
    OpnFxX X, F
    X.Workbooks.Open F, UpdateLinks:=False
Next
ArrangeWbV X
End Sub

Sub OpnFxX(X As Excel.Application, Fx)  ' New Xls will be visible.  Don't open Xls invisible, always pass @X to open Fx, where @X should be visible
X.Workbooks.Open Fx, UpdateLinks:=False
End Sub
Sub OpnFcsv(Fcsv)
Xls.Workbooks.OpenText Fcsv
End Sub
