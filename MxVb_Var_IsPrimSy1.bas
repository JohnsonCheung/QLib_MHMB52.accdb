Attribute VB_Name = "MxVb_Var_IsPrimSy1"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Var_IsPrimSy1."


Function IsDteSy(Sy$()) As Boolean
Dim S: For Each S In Sy
    If Not IsDte(S) Then Exit Function
Next
IsDteSy = True
End Function

Function IsDblSy(Sy$()) As Boolean
Dim S: For Each S In Sy
    If Not IsVDblVdt(S) Then Exit Function
Next
IsDblSy = True
End Function
