Attribute VB_Name = "MxDao_Sql_Fmt_FmtSql"
Option Compare Text
Option Explicit
Private Sub B_FmtSqlStruC_()
GoSub T1
Exit Sub
Dim Q$
T1:
    Q = "Select CStr(x.Sku) As Sku,Val(Nz(x.M01,0)) As M01,Val(Nz(x.M02,0)) As M02,Val(Nz(x.M03,0)) As M03,Val(Nz(x.M04,0)) As M04,Val(Nz(x.M05,0)) As M05,Val(Nz(x.M06,0)) As M06,Val(Nz(x.M07,0)) As M07,Val(Nz(x.M08,0)) As M08,Val(Nz(x.M09,0)) As M09,Val(Nz(x.M10,0)) As M10,Val(Nz(x.M11,0)) As M11,Val(Nz(x.M12,0)) As M12,Co into [#Fc] from [#IFc] as x where CStr(Nz(Sku,''))<>'' and (Val(Nz(M01,0))<>0 or Val(Nz(M02,0))<>0 or Val(Nz(M03,0))<>0 or Val(Nz(M04,0))<>0 or Val(Nz(M05,0))<>0 or Val(Nz(M06,0))<>0 or Val(Nz(M07,0))<>0 or Val(Nz(M08,0))<>0 or Val(Nz(M09,0))<>0 or Val(Nz(M10,0))<>0 or Val(Nz(M11,0))<>0 or Val(Nz(M12,0))<>0)"
    Act = FmtSqlStruC(Q)
    Debug.Print Act
    Return
End Sub
Private Sub B_FmtSqlDb()
'GoSub Z4
GoSub ZZ
'GoSub Z1
'GoSub Z2
'GoSub Z3
'GoSub T1
Exit Sub
Dim Sql$
ZZ:
    BrwAy FmtSqlDb(CDb)
    Return
T1:
    Sql = "Update [$PHLSku] set Stm=IIf(Left(CdTopaz,3)='UDV','U','M')"
    Ept = RplVBar("Update [$PHLSku]|Set Stm = IIf(Left(CdTopaz,3)='UDV','U','M')")
    GoTo Tst
Dim Qn$
Z1: Qn = "qExpOH": GoTo ZTst
Z2: Qn = "qOHZM01": GoTo ZTst
Z3: Qn = "qOHMB52": GoTo ZTst
Z4: Qn = "dupPermitSku": GoTo ZTst
ZTst: Dmp Qn: Dmp ULnHdr(Qn): Debug.Print FmtSql(SqlQnC(Qn)): Debug.Print: Return
Tst:
    Act = FmtSql(Sql)
    C
    Return
End Sub
Function FmtSqlDb(D As Database) As String()
Dim OLy$(), Sql$, FmtSql$, J%
Dim Q: For Each Q In Itr(Qny(D))
    J = J + 1
    PushIAy OLy, LyUL("#" & J & " " & Q)
    Sql = SqlQn(D, Q)
    FmtSql = FmtSqlStruTQp(D, Sql)
    PushI OLy, LinesTab4Spc(TrimWhite(Sql))
    PushI OLy, vbSpc4 & ULnLines(Sql, "-")
    PushI OLy, LinesTab4Spc(FmtSql)
    PushI OLy, ""
Next
FmtSqlDb = OLy
End Function
Function FmtSqlStruC$(Q):               FmtSqlStruC = FmtSqlStru(CDb, Q):  End Function
Function FmtSqlStru$(D As Database, Q):  FmtSqlStru = FmtSqlStruTQp(D, Q): End Function
Function FmtSql$(Q):                         FmtSql = FmtSqlTQp(Q):        End Function
