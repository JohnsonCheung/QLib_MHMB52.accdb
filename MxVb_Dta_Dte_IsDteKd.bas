Attribute VB_Name = "MxVb_Dta_Dte_IsDteKd"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_DteIs."
Function IsHHMMDD(S) As Boolean
Select Case True
Case _
    Len(S) <> 6, _
    Not IsHH(Left(S, 2)), _
    Not Is0059(Mid(S, 3, 2)), _
    Not Is0059(Right(S, 2))
Case Else: IsHHMMDD = True
End Select
End Function

Function IsHH(S) As Boolean
Select Case True
Case _
    Len(S) <> 2, _
    Not IsAllDig(S), _
    "00" > S, S > "23"
Case Else: IsHH = True
End Select
End Function

Function Is0059(S) As Boolean
Select Case True
Case _
    Len(S) <> 2, _
    Not IsAllDig(S), _
    "00" > S, S > "59"
Case Else: Is0059 = True
End Select
End Function

Function IsYYYYMMDD(S) As Boolean
If Len(S) <> 8 Then Exit Function
If Not IsYYYY(Left(S, 4)) Then Exit Function
If Not IsMM(Mid(S, 5, 2)) Then Exit Function
If Not IsDD(Right(S, 2)) Then Exit Function
IsYYYYMMDD = True
End Function

Function IsMM(S) As Boolean
If Len(S) <> 2 Then Exit Function
If Not IsAllDig(S) Then Exit Function
If S < "00" Then Exit Function
If S > "12" Then Exit Function
IsMM = True
End Function
Function IsYYYY(S) As Boolean
Select Case True
Case Len(S) <> 4, Not IsAllDig(S), S < "2000"
Case Else: IsYYYY = True
End Select
End Function
Function IsDD(S) As Boolean
Select Case True
Case Len(S) <> 2, Not IsAllDig(S), S < "00", "31" < S
Case Else: IsDD = True
End Select
End Function

Function IsHMS(A$) As Boolean
If Len(A) <> 6 Then Exit Function
IsHMS = IsDate(Format(A, "00:00:00"))
End Function

Function IsDashYYMD(A$) As Boolean
Select Case True
Case Len(A) <> 10, Mid(A, 5, 1) <> "-", Mid(A, 8, 1) <> "-": Exit Function
End Select
IsDashYYMD = IsDate(A)
End Function
Function IsTimStr(S) As Boolean
Select Case True
Case Len(S) <> 19, _
    Mid(S, 11, 1) <> " ", _
    Not IsDashYYMD(Left(S, 10)), _
    Not IsHMS(Right(S, 6))
    Exit Function
End Select
IsTimStr = True
End Function
