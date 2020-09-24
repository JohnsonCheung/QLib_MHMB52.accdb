Attribute VB_Name = "MxIde_Mth_Prp_MthCxt"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_Cxt."
Function CxtMth(Mthy$()) As String()
Dim N%: N = NContln(Mthy, 0)
CxtMth = AeLas(AeFstN(Mthy, N))
End Function
