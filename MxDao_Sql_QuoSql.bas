Attribute VB_Name = "MxDao_Sql_QuoSql"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Sql_QuoSql."
Function QuoSqlStr$(Sqlstr$) ' If @SqlStr has only single-quote, quote by single.  If has only double, quote single.  Else, quote single and replace inside-single as 2-single.
Dim WiSng As Boolean, WiDbl As Boolean
WiSng = HasQuoSng(Sqlstr)
WiDbl = HasQuoDbl(Sqlstr)
Dim O$
Select Case True
Case WiSng And WiDbl: O = QuoSng(Replace(Sqlstr, vbQuoSng, vbQuoSng2))
Case WiSng: O = QuoDbl(Sqlstr)
Case WiDbl: O = QuoSng(Sqlstr)
Case Else:  O = QuoSng(Sqlstr)
End Select
QuoSqlStr = O
End Function

Function QuoSqlPrimy(Primy) As String()
Dim Prim: For Each Prim In Primy
    PushI QuoSqlPrimy, QuoSqlPrim(Prim)
Next
End Function

Function QuoSqlPrim$(Prim)
Const CSub$ = CMod & "QuoSqlPrim"
Dim V: V = Prim
Select Case True
Case IsStr(V): QuoSqlPrim = QuoSqlStr(CStr(V))
Case IsDte(V): QuoSqlPrim = QuoDte(V)
Case IsNumeric(V), IsBool(V): QuoSqlPrim = V
Case IsEmpty(V): QuoSqlPrim = "null"
Case Else
    Thw CSub, "V should be Dte Str, Numeric or Empty", "TypeName(V)", TypeName(V)
End Select
End Function

Function QuoSqlT$(T, Optional Alias$):          QuoSqlT = QuoSqlTorF(T, Alias):                   End Function
Function QuoSqlF$(F, Optional Alias$):          QuoSqlF = QuoSqlTorF(F, Alias):                   End Function
Function AliasIf$(Alias$):                      AliasIf = StrPfxIfNB(".", Alias):                 End Function
Function AsIf$(Alias$):                            AsIf = StrTrue(Alias <> "", " As " & Alias): End Function
Function QuoSqlTorF$(TorF, Optional Alias$): QuoSqlTorF = AliasIf(Alias) & WQuo(TorF):            End Function
Private Function WQuo$(Sqln)
If WShd(Sqln) Then WQuo = QuoSq(Sqln) Else WQuo = Sqln
End Function
Private Function WShd(Sqln) As Boolean: WShd = WRx.Test(Sqln): End Function
Private Function WRx() As RegExp
Static X As RegExp: If IsNothing(X) Then Set X = Rx("^[^A-Za-z]|[A-Za-z][^\W]+")
Set WRx = X
End Function

Function ChrQuoSqlzDaoTy$(A As Dao.DataTypeEnum)
Const CSub$ = CMod & "ChrQuoSqlzDaoTy"
Select Case A
Case _
    Dao.DataTypeEnum.dbBigInt, _
    Dao.DataTypeEnum.dbByte, _
    Dao.DataTypeEnum.dbCurrency, _
    Dao.DataTypeEnum.dbDecimal, _
    Dao.DataTypeEnum.dbDouble, _
    Dao.DataTypeEnum.dbFloat, _
    Dao.DataTypeEnum.dbInteger, _
    Dao.DataTypeEnum.dbLong, _
    Dao.DataTypeEnum.dbNumeric, _
    Dao.DataTypeEnum.dbSingle: Exit Function
Case _
    Dao.DataTypeEnum.dbChar, _
    Dao.DataTypeEnum.dbMemo, _
    Dao.DataTypeEnum.dbText: ChrQuoSqlzDaoTy = "'"
Case _
    Dao.DataTypeEnum.dbDate: ChrQuoSqlzDaoTy = "#"
Case Else
    Thw CSub, "Invalid Daoty", "Daoty", A
End Select
End Function
