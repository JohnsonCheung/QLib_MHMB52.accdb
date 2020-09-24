Attribute VB_Name = "MxIde_Mthn_Fun"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthn_Fun."

Function PrpTy$(ShtMthTy$) ' Ret ShtMthTy if they are Get|Set|Let else Blank
Select Case ShtMthTy
Case "Get", "Let", "Set": PrpTy = ShtMthTy
End Select
End Function
