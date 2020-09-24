Attribute VB_Name = "MxVb_Dta_ItrIs"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_ItrIs."

Function IsItrLines(Itr) As Boolean
Dim V: For Each V In Itr
    If Not IsStr(V) Then Exit Function
    If HasLf(V) Then IsItrLines = True: Exit Function
Next
End Function

Function IsItrLn(Itr) As Boolean
Dim V: For Each V In Itr
    If Not IsStr(V) Then Exit Function
    If HasLf(V) Then Exit Function
Next
IsItrLn = True
End Function

Function IsItrAy(Itr) As Boolean
Dim V: For Each V In Itr
    If Not IsArray(V) Then Exit Function
Next
IsItrAy = True
End Function
Function IsItrStr(Itr) As Boolean
Dim V: For Each V In Itr
    If Not IsStr(V) Then Exit Function
Next
IsItrStr = True
End Function

Function IsVPrimItr(Itr) As Boolean
Dim V: For Each V In Itr
    If Not IsPrim(V) Then Exit Function
Next
IsVPrimItr = True
End Function

Function IsSyItr(Itr) As Boolean
Dim V: For Each V In Itr
    If Not IsSy(V) Then Exit Function
Next
IsSyItr = True
End Function

Function IsBoolItr(Itr) As Boolean
Dim V: For Each V In Itr
    If Not IsBool(V) Then Exit Function
Next
IsBoolItr = True
End Function

Function IsNumItr(Itr) As Boolean
Dim V: For Each V In Itr
    If Not IsNumeric(V) Then Exit Function
Next
IsNumItr = True
End Function

Function IsDteItr(Itr) As Boolean
Dim V: For Each V In Itr
    If Not IsDte(V) Then Exit Function
Next
IsDteItr = True
End Function

Function IsItrEqNm(ItrA, ItrB)
IsItrEqNm = IsSamAy(Itn(ItrA), Itn(ItrB))
End Function
