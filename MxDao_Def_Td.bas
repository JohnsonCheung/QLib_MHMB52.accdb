Attribute VB_Name = "MxDao_Def_Td"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Def_Td."

Sub AddPk(D As Database, T As Dao.TableDef)
If WHasIdFld(T) Then
    T.Indexes.Append WPk(D, T.Name)
End If
End Sub
Private Function WHasIdFld(T As Dao.TableDef) As Boolean
If T.Fields(0).Name <> T.Name & "Id" Then Exit Function
If T.Attributes <> Dao.FieldAttributeEnum.dbAutoIncrField Then Exit Function
If T.Fields(0).Type <> dbLong Then Exit Function
WHasIdFld = True
End Function
Private Function WPk(D As Database, T$) As Dao.Index
Dim O As New Dao.Index
O.Name = "PrimaryKey"
O.Primary = True
CvIdxFds(O.Fields).Append FdId(T & "Id")
Set WPk = O
End Function

Sub AddSk(T As Dao.TableDef, Skff$): Stop '          T.Indexes.Append WSk(T, Tml(Skff)): End Sub
End Sub
Private Function WSk(Td As Dao.TableDef, Fny$()) As Dao.Index: Set WSk = IdxNw(Td, Td.Name, Fny, IsUKy:=True): End Function

Function Fdy(FF$, T As Dao.DataTypeEnum) As Dao.Field2()
Dim F: For Each F In FnyFF(FF)
    PushObj Fdy, FdC(F, T)
Next
End Function

Function FnyTd(T As Dao.TableDef) As String(): FnyTd = Itn(T.Fields): End Function

Function IsTdEq(T As Dao.TableDef, B As Dao.TableDef) As Boolean
With T
Select Case True
Case .Name <> B.Name
Case .Attributes <> B.Attributes
Case Not IsEqIdxs(.Indexes, B.Indexes)
'Case Not FdsIsEq(.Fields, B.Fields)
Case Else: IsTdEq = True
End Select
End With
End Function

Function IsTdHid(T As Dao.TableDef) As Boolean: IsTdHid = (T.Attributes And Dao.TableDefAttributeEnum.dbHiddenObject) <> 0:  End Function
Function IsTdSys(T As Dao.TableDef) As Boolean: IsTdSys = (T.Attributes And Dao.TableDefAttributeEnum.dbSystemObject) <> 0:  End Function
Function IsTdLnk(T As Dao.TableDef) As Boolean: IsTdLnk = (T.Attributes And Dao.TableDefAttributeEnum.dbAttachedTable) <> 0: End Function

Sub ChkIsEqStruT(D As Database, T1, T2)
Const CSub$ = CMod & "ChkIsEqStruT"
Dim A$: A = StruT(D, T1)
Dim B$: B = StruT(D, T2)
If A <> B Then Thw CSub, "Two 2 Td as diff", "Td-A Td-B", A, B
End Sub

Function SqlQnC$(Qn):               SqlQnC = SqlQn(CDb, Qn):  End Function
Function SqlQn$(D As Database, Qn):  SqlQn = Qd(CDb, Qn).Sql: End Function
Function SqyQry(D As Database) As String()
Dim Qd As QueryDef: For Each Qd In D.QueryDefs
    PushI SqyQry, Qd.Sql
Next
End Function
Function SqyQryC() As String(): SqyQryC = SqyQry(CDb): End Function

Function Fd(D As Database, T, F) As Dao.Field:  Set Fd = Td(D, T).Fields(F): End Function
Function FdC(T, F) As Dao.Field:               Set FdC = Fd(CDb, T, F):      End Function
Function Td(D As Database, T) As Dao.TableDef:
D.TableDefs.Refresh
Set Td = D.TableDefs(T)
End Function
Function TdC(T) As Dao.TableDef: Set TdC = Td(CDb, T): End Function
