Attribute VB_Name = "MxDao_Dbt_Da"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Dbt_Da."
Function DsDbC(Optional Dsn$) As Ds:               DsDbC = DsDb(CDb, Dsn):                        End Function
Function DsDb(D As Database, Optional Dsn$) As Ds:  DsDb = DsTny(D, Tny(D), StrDft(Dsn, D.Name)): End Function
Function DsTny(D As Database, Tny$(), Optional Dsn$) As Ds
Dim T: For Each T In Tny
    PushDt DsTny.Dty, DtT(D, T)
Next
End Function

Function DrsT(D As Database, T) As Drs:             DrsT = DrsRs(RsTbl(D, T)):          End Function
Function DrsTC(T) As Drs:                          DrsTC = DrsT(CDb, T):                End Function
Function DtT(D As Database, T) As Dt:                DtT = Dt(T, Fny(D, T), DyT(D, T)): End Function
Function DyT(D As Database, T) As Variant():         DyT = DyRs(RsTbl(D, T)):           End Function
Function DyFf(D As Database, T, FF$) As Variant():  DyFf = DyQ(D, SqlSelFf(T, FF)):     End Function
