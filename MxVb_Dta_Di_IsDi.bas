Attribute VB_Name = "MxVb_Dta_Di_IsDi"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_DicIs."

Function IsEqDic(V As Dictionary, B As Dictionary, Optional C As eCas) As Boolean
If V.Count <> B.Count Then Exit Function
If V.Count = 0 Then IsEqDic = True: Exit Function
Dim K1, K2
K1 = AySrtQ(V.Keys)
K2 = AySrtQ(B.Keys)
If Not IsEqAy(K1, K2) Then Exit Function
Dim K
For Each K In K1
   If B(K) <> V(K) Then Exit Function
Next
IsEqDic = True
End Function

Function IsDikNm(A As Dictionary) As Boolean
Dim K: For Each K In A.Keys
    If Not IsNm(K) Then Exit Function
Next
IsDikNm = True
End Function
Function IsDikStr(D As Dictionary) As Boolean
IsDikStr = IsSyItr(D.Keys)
End Function

Function IsEmpDic(D As Dictionary) As Boolean
IsEmpDic = D.Count = 0
End Function

Function IsDiiLy(D As Dictionary) As Boolean
IsDiiLy = IsDiiSy(D)
End Function

Function IsDiiSy(D As Dictionary) As Boolean
Select Case True
Case Not IsSyItr(D.Items), Not IsItrStr(D.Keys)
Case Else: IsDiiSy = True
End Select
End Function

Function IsDiiAy(D As Dictionary) As Boolean
If Not IsItrAy(D.Items) Then Exit Function
IsDiiAy = True
End Function

Sub ChkIsDi12SamKey(D As Dictionary, B As Dictionary, Fun$)
If Not IsDi12SamKey(D, B) Then Thw Fun, "Give 2 dictionary are not EqKey"
End Sub

Function IsDi12SamKey(D As Dictionary, B As Dictionary) As Boolean
If D.Count <> B.Count Then Exit Function
Dim K: For Each K In D.Keys
    If Not B.Exists(K) Then Exit Function
Next
IsDi12SamKey = True
End Function

Sub ChkIsDiiVStr(D As Dictionary, Fun$)
If Not IsDiiStr(D) Then Thw Fun, "Given Dic is not StrDic"
End Sub
Sub ChkIsDiiSy(D As Dictionary, Fun$)
If Not IsDiiSy(D) Then Thw Fun, "Given Dic is not SyDic"
End Sub

Sub ChkIsDiiLines(D As Dictionary, Fun$)
If Not IsDiiLines(D) Then Thw Fun, "Given Dic is not DiLines"
End Sub

Function IsDiiLines(D As Dictionary) As Boolean
Select Case True
Case Not IsItrLines(D.Items), Not IsItrStr(D.Keys)
Case Else: IsDiiLines = True
End Select
End Function

Function IsDiiStr(D As Dictionary) As Boolean
If Not IsItrStr(D.Keys) Then Exit Function
IsDiiStr = IsItrStr(D.Items)
End Function

Function IsDiiPrim(D As Dictionary) As Boolean
If Not IsVPrimItr(D.Keys) Then Exit Function
IsDiiPrim = IsVPrimItr(D.Items)
End Function
