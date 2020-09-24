Attribute VB_Name = "MxIde_MthEmp"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_MthEmp."

Function EmpTstSubl$(Mthn) ' :Lines #Empty-Tst-Sub-Mth-Lines#
Dim L1$, L2$
L1 = FmtQQ("Private Sub ?__Tst()", Mthn)
L2 = "End Sub"
EmpTstSubl = L1 & vbCrLf & L2
End Function
