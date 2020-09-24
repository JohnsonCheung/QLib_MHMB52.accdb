Attribute VB_Name = "MxVb_Var_IsS"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Var_IsS."
Const C_UCasA% = 65
Const C_UCasZ% = 90
Const C_LCasA% = 97
Const C_LCasZ% = 122

Function IsVItr(V) As Boolean: IsVItr = TypeName(V) = "Collection": End Function
Function IsLines(V) As Boolean
If Not IsStr(V) Then Exit Function
IsLines = HasLf(V)
End Function
Function IsLsy(V) As Boolean
If Not IsArray(V) Then Exit Function
If Not IsSy(V) Then Exit Function
Dim L: For Each L In Itr(V)
    If HasLf(L) Then
        IsLsy = True: Exit Function
    End If
Next
End Function
Function IsVAyy(V) As Boolean
If Not IsAv(V) Then Exit Function
Dim X
For Each X In Itr(V)
    If Not IsArray(X) Then Exit Function
Next
IsVAyy = True
End Function
Function IsVLy(V) As Boolean
If Not IsVLy(V) Then Exit Function
Dim L: For Each L In Itr(V)
    If IsVLn(L) Then Exit Function
Next
End Function
Function IsVLn(V) As Boolean
If Not IsStr(V) Then Exit Function
IsVLn = NoLf(V)
End Function
Function IsNothing(V) As Boolean
If Not IsObject(V) Then Exit Function
IsNothing = ObjPtr(V) = 0
End Function
Function IsSomething(V) As Boolean: IsSomething = Not IsNothing(V): End Function
Function IsPrim(V) As Boolean
Select Case VarType(V)
Case _
   VbVarType.vbBoolean, _
   VbVarType.vbByte, _
   VbVarType.vbCurrency, _
   VbVarType.vbDate, _
   VbVarType.vbDecimal, _
   VbVarType.vbDouble, _
   VbVarType.vbInteger, _
   VbVarType.vbLong, _
   VbVarType.vbSingle, _
   VbVarType.vbString
   IsPrim = True
End Select
End Function
Function IsSBlnkN(V) As Boolean
If Not IsStr(V) Then Exit Function
IsSBlnkN = V <> ""
End Function

Function IsStrDte(S$) As Boolean
On Error GoTo X
CVDate S
IsStrDte = True
Exit Function
X:
End Function

Function IsStr(V) As Boolean:       IsStr = VarType(V) = vbString:  End Function
Function IsBool(V) As Boolean:     IsBool = VarType(V) = vbBoolean: End Function
Function IsByt(V) As Boolean:       IsByt = VarType(V) = vbByte:    End Function
Function IsDte(V) As Boolean:       IsDte = VarType(V) = vbDate:    End Function
Function IsLng(V) As Boolean:       IsLng = VarType(V) = vbLong:    End Function
Function CanCvInt(V) As Boolean: CanCvInt = VarType(V) = vbInteger: End Function
Function IsDbl(V) As Boolean:       IsDbl = VarType(V) = vbDouble:  End Function
Function IsStrInt(V) As Boolean
On Error GoTo X
Dim A%: A = V
IsStrInt = True
X:
End Function
Private Sub B_IsSy()
Dim V$()
Dim B: B = V
Dim C()
Dim D
Ass IsSy(V) = True
Ass IsSy(B) = True
Ass IsSy(C) = False
Ass IsSy(D) = False
End Sub
Function IsSy(V) As Boolean: IsSy = VarType(V) = vbArray + vbString: End Function
Sub ChkIsSy(V, Fun$)
If IsSy(V) Then Exit Sub
Thw Fun, "V must be Sy", "TypeName(V)", TypeName(V)
End Sub
Function IsAv(V) As Boolean:       IsAv = VarType(V) = vbArray + vbVariant: End Function
Function IsDtey(V) As Boolean:   IsDtey = VarType(V) = vbArray + vbDate:    End Function
Function IsByty(V) As Boolean:   IsByty = VarType(V) = vbByte + vbArray:    End Function
Function IsLngy(V) As Boolean:   IsLngy = VarType(V) = vbArray + vbLong:    End Function
Function IsInty(V) As Boolean:   IsInty = VarType(V) = vbArray + vbInteger: End Function
Function IsBooly(V) As Boolean: IsBooly = VarType(V) = vbArray + vbBoolean: End Function
Function IsObjy(V) As Boolean:   IsObjy = VarType(V) = vbArray + vbObject:  End Function

Function IsEqVbt(A, B) As Boolean: IsEqVbt = VarType(A) = VarType(B): End Function
Function IsEqV(A, B, Optional C As eCas) As Boolean
If Not IsEqVbt(A, B) Then Exit Function
Select Case True
Case IsStr(A): IsEqV = IsEqStr(A, B, C)
Case IsArray(A): IsEqV = IsEqAy(A, B)
Case IsDi(A): IsEqV = IsEqDic(CvDi(A), CvDi(B), C)
Case IsObject(A): IsEqV = ObjPtr(A) = ObjPtr(B)
Case Else: IsEqV = A = B
End Select
End Function
Function IsTynPrim(Tyn$) As Boolean
Select Case Tyn
Case _
   "Boolean", _
   "Byte", _
   "Currency", _
   "Date", _
   "Decimal", _
   "Double", _
   "Integer", _
   "Long", _
   "Single", _
   "String"
   IsTynPrim = True
End Select
End Function

Function IsQuo(S, Q1$, Optional ByVal Q2$) As Boolean
If Q2 = "" Then Q2 = Q1
If ChrFst(S) <> Q1 Then Exit Function
IsQuo = ChrLas(S) = Q2
End Function

Function IsQuoSng(S) As Boolean: IsQuoSng = IsQuo(S, "'"):      End Function
Function IsQuoDbl(S) As Boolean: IsQuoDbl = IsQuo(S, vbQuoDbl): End Function

Function IsSngQuoNeed(S) As Boolean
If IsQuoSq(S) Then Exit Function
Select Case True
Case IsAscDig(Asc(ChrFst(S))), HasSpc(S), HasDot(S), HasHyp(S), HasPound(S): IsSngQuoNeed = True
End Select
End Function

Function IsVbtNum(A As VbVarType) As Boolean
Select Case A
Case vbByte, vbInteger, vbLong, vbSingle, vbDecimal, vbDouble, vbCurrency: IsVbtNum = True
End Select
End Function

Function IsBlnk(V) As Boolean
Select Case True
Case IsStr(V): IsBlnk = IsStrBlnk(CStr(V))
Case IsNull(V), IsEmpty(V), IsMissing(V): IsBlnk = True
End Select
End Function

Function IsStrBlnk(S$) As Boolean: IsStrBlnk = Trim(S) = "": End Function
Function IsOdd(N) As Boolean:          IsOdd = N Mod 2 = 1:  End Function
Function IsEven(N) As Boolean:        IsEven = N Mod 2 = 0:  End Function

Function IsBet(V, A, B) As Boolean
If A > V Then Exit Function
If V > B Then Exit Function
IsBet = True
End Function

Sub ChkIsBet(V, A, B, Fun$)
If IsBet(V, A, B) Then Exit Sub
Thw Fun, "V is not between A and B", "V A B", V, A, B
End Sub

Function IsErObj(A) As Boolean: IsErObj = TypeName(A) = "Error": End Function

Function IsEmp(V) As Boolean
Select Case True
Case IsStr(V):    IsEmp = Trim(V) = ""
Case IsArray(V):  IsEmp = Si(V) = 0
Case IsEmpty(V), IsNothing(V), IsMissing(V), IsNull(V): IsEmp = True
End Select
End Function

Function IsNBet(V, A, B) As Boolean:  IsNBet = Not IsBet(V, A, B): End Function
Function IsQuoSq(S) As Boolean:      IsQuoSq = IsQuo(S, "[", "]"): End Function

Function IsIxyOutRg(Ixy, U&) As Boolean '#Is-Ixy-Out-Range#
Dim Ix: For Each Ix In Itr(Ixy)
    If 0 > Ix Then IsIxyOutRg = True: Exit Function
    If Ix > U Then IsIxyOutRg = True: Exit Function
Next
End Function

Function IsEqStr(A, B, Optional C As eCas) As Boolean:    IsEqStr = StrComp(A, B, VbCprMth(C)) = 0:                     End Function
Function IsEqStrSen(A, B) As Boolean:                  IsEqStrSen = StrComp(A, B, VbCompareMethod.vbBinaryCompare) = 0: End Function

Function IsVDblVdt(V) As Boolean ' Is @S convertable to Dbl
On Error GoTo X
Dim A#: A = V
IsVDblVdt = True
Exit Function
X:
End Function

Function IsEmpAy(Ay) As Boolean:   IsEmpAy = Si(Ay) = 0:                  End Function
Function IsSamAy(A, B) As Boolean: IsSamAy = IsEqDic(DiCnt(A), DiCnt(B)): End Function

Function IsAyEleEqAll(Ay) As Boolean
Const CSub$ = CMod & "IsAyEleEqAll"
If Si(Ay) <= 1 Then IsAyEleEqAll = True: Exit Function
Dim A0, J&
A0 = Ay(0)
For J = 1 To UB(Ay)
    If A0 <> Ay(J) Then Exit Function
Next
IsAyEleEqAll = True
End Function

Function IsEqSi(A, B) As Boolean: IsEqSi = Si(A) = Si(B):              End Function
Function IsNeAy(A, B) As Boolean: IsNeAy = Not IsEqAy(A, B):           End Function
Function IsEqDy(A, B) As Boolean: IsEqDy = IsEqAy(A, B):               End Function
Function IsDi(V) As Boolean:        IsDi = TypeName(V) = "Dictionary": End Function
Function IsAllDig(S) As Boolean
Dim J%: For J = 1 To Len(S)
    If Not IsAscDig(Asc(Mid(S, J, 1))) Then Exit Function
Next
IsAllDig = True
End Function

