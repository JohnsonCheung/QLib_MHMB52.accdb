Attribute VB_Name = "MxIde_Mthln_MthTy_MthTyCnst"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthln_MthTy_MthTyCnst."

Function ShtMthKdy() As String()
Static X$(): If Si(X) = 0 Then X = Sy("Fun", "Sub", "Prp")
ShtMthKdy = X
End Function

Function PrpTyy() As String()
Static X$()
If Si(X) = 0 Then X = Sy("Get", "Set", "Let")
PrpTyy = X
End Function

Function MthTyy() As String()
Static X$(): If Si(X) = 0 Then X = Sy("Function", "Sub", "Property Get", "Property Set", "Property Let")
MthTyy = X
End Function
Function ShtMthTyy() As String()
Static X$()
If Si(X) = 0 Then X = Sy("Fun", "Sub", "Get", "Set", "Let")
ShtMthTyy = X
End Function

Function MthKdy() As String()
Static X$()
If Si(X) = 0 Then X = Sy("Fun", "Sub", "Prp")
MthKdy = X
End Function
