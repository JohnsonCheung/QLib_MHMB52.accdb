Attribute VB_Name = "MxVb_Var_Dft"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Dft."
Function ValDft(V, Dft)
If IsEmp(V) Then
   ValDft = Dft
Else
   ValDft = V
End If
End Function

Function StrDft$(Str, Dft): StrDft = IIf(Str = "", Dft, Str): End Function
Function ValLimit(V, A, B)
Select Case V
Case V > B: ValLimit = B
Case V < A: ValLimit = A
Case Else: ValLimit = V
End Select
End Function
