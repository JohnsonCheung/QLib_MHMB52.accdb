Attribute VB_Name = "MxXls_A_AppAct"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_A_ActApp."
Sub ActAppWb(B As Workbook):           ActAppFx B.Name:            End Sub
Sub ActAppFx(Fx$):                     ActApp Fn(Fx) & " - Excel": End Sub
Sub ActAppXls(X As Excel.Application): ActAppWb X.ActiveWorkbook:  End Sub
Sub ActApp(Tit$)
On Error GoTo X
Interaction.AppActivate Tit
Exit Sub
X: Inf CSub, "No window with @Tit", "@Tit", Tit
End Sub
