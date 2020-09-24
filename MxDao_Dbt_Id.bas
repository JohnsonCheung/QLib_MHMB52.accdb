Attribute VB_Name = "MxDao_Dbt_Id"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Dbt_Id."

Function IdRecLas&(D As Database, T)
'@T ! Assume it has a field <T>Id and a "PrimaryKey", using the field as Key
ChkHasRid D, T
Dim R As Dao.Recordset: Set R = D.TableDefs(T).OpenRecordset
R.Index = "PrimaryKey"
R.MoveLast
IdRecLas = R.Fields(0).Value
End Function

Sub ChkHasRid(D As Database, T, Optional Fun$ = "ChkHasRid")
If Not HasRid(D, T) Then Thw Fun, "Given table is not Id-Table (should have Id-Fld Id-Pk)", "Db T", D.Name, T
End Sub

Function HasRid(D As Database, T) As Boolean
Select Case True
Case NoFldRid(D, T): Exit Function
Case NoPkRid(D, T): Exit Function
End Select
HasRid = True
End Function

Function HasFldRid(D As Database, T) As Boolean: HasFldRid = D.TableDefs(T).Fields(0).Name = T & "Id": End Function
Function NoFldRid(D As Database, T) As Boolean:   NoFldRid = Not HasFldRid(D, T):                      End Function
Function HasPkRid(D As Database, T) As Boolean:   HasPkRid = IsEqAy(FnyPk(D, T), Sy(T & "Id")):        End Function
Function NoPkRid(D As Database, T) As Boolean:     NoPkRid = HasPkRid(D, T):                           End Function

Function RidC&(T, N$):               RidC = Rid(CDb, T, N):                                               End Function
Function Rid&(D As Database, T, N$):  Rid = ValQ(D, FmtQQ("Select ?Id from ? where ?n='?'", T, T, T, N)): End Function
