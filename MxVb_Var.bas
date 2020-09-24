Attribute VB_Name = "MxVb_Var"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Var_."

Function ValTrue(Cond As Boolean, V):   ValTrue = IIf(Cond = True, V, Empty):  End Function
Function ValFalse(Cond As Boolean, V): ValFalse = IIf(Cond = False, V, Empty): End Function

Function EnsBet(I, A, B)
Select Case True
Case I < A: EnsBet = A
Case I > B: EnsBet = B
Case Else: EnsBet = I
End Select
End Function
