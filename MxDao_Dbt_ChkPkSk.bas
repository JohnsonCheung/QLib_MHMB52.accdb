Attribute VB_Name = "MxDao_Dbt_ChkPkSk"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Dbt_ChkPkSk."
Sub ChkPk(D As Database, T): ChkEr ErPk(D, T): End Sub
Function ErPk$(D As Database, T)
If HasPk(D, T) Then Exit Function
Dim Pk$(): Pk = FnyPk(D, T)
Select Case True
Case Si(Pk) = 0: ErPk = FmtQQ("[?] does not have PrimaryKey-Idx", T)
Case Si(Pk) <> 1: ErPk = FmtQQ("There is PrimaryKey-Idx, but it has [?] fields[?]", Si(Pk), Tml(Pk))
Case Pk(0) <> T & "Id": ErPk = FmtQQ("There is One-field-PrimaryKey-Idx of Fldn(?), but it should named as ?Id", Pk(0), T)
Case Fd(D, T, 0).Name <> T & "Id": ErPk = FmtQQ("The Pk-field(?Id) should be first fields, but now it is (?)", T, Fd(D, T, T & "Id").OrdinalPosition)
End Select
End Function
Function ErPkSk$(D As Database, T)
ErPkSk = ErPk(D, T): If ErPkSk <> "" Then Exit Function
ErPkSk = ErSk(D, T)
End Function
Function ErSk$(D As Database, T)

End Function
Function EryPkSkC() As String(): EryPkSkC = EryPkSk(CDb): End Function
Function EryPkSk(D As Database) As String()
Dim T: For Each T In Tny(D)
    PushIAy EryPkSk, ErPkSk(D, T)
Next
End Function

Function IdxSk(D As Database, T) As Dao.Index
Dim I As Dao.Index: For Each I In D.TableDefs(T).Indexes
    If Not I.Unique Then GoTo N
    If I.Name <> T Then GoTo N
    Set IdxSk = I
    Exit Function
N: Next
End Function
Function IdxPPk(D As Database, T) As Dao.Index
'PPk: :Dao.Index|Nothing #(P)rimary-PrimaryKey(Pk)-(Idx)# A Pk of @D-@T having (1)IsPk (2)1-Fld (3)LngTy (4)Nm=Tbn&Id (5)AutoIncr
Dim I As Dao.Index: For Each I In D.TableDefs(T).Indexes
    If Not I.Primary Then GoTo N
    If I.Fields.Count <> 1 Then GoTo N
    Dim F As Dao.Field: Set F = I.Fields(0)
    If F.Name <> T & "Id" Then GoTo N
    If F.Type <> Dao.dbLong Then GoTo N
    If Not (F.Attributes And Dao.FieldAttributeEnum.dbAutoIncrField) Then GoTo N
    Set IdxPPk = I
    Exit Function
N: Next
End Function
Function ChkSk$(D As Database, T) '#(ChkSk)-Exist-in-@D-@T#
If Not HasSk(D, T) Then
    ChkSk = FmtQQ("No SecondaryKey for Table[?] in Db[?]", T, D.Name)
    Exit Function
End If
End Function

Function ChkSsk$(D As Database, T) '#(ChkSsk)-is-exist-in-@D-@T# See :Ssk
Dim O$, Sk$(): Sk = FnySk(D, T)
O = ChkSk(D, T): If O <> "" Then ChkSsk = O: Exit Function
If Si(Sk) <> 1 Then
'    ChkSsk = FmtQQ("Secondary is not single field. Tbl[?] Db[?] SkFfn[?]", T, D.Name, JnTmy(Sk))
End If
End Function
Function HasPk(D As Database, T) As Boolean:      HasPk = HasTruePrp(Td(D, T).Indexes, "Primary"): End Function
Function HasPkTd(A As Dao.TableDef) As Boolean: HasPkTd = HasTruePrp(A.Indexes, "Primary"):        End Function
Function HasSk(D As Database, T) As Boolean:      HasSk = Not IsNothing(IdxSk(D, T)):              End Function
Function HasPPk(D As Database, T) As Boolean:    HasPPk = HasPPkTd(D.TableDefs(T)):                End Function

Function HasPPkTd(T As Dao.TableDef) As Boolean
Dim Pk$(): Pk = FnyPkTd(T)
If Si(Pk) <> 1 Then Exit Function
HasPPkTd = Pk(0) = T.Name & "Id"
End Function
Function HasSkTd(T As Dao.TableDef) As Boolean
Const CSub$ = CMod & "HasSkTd"
If Not HasItn(T.Indexes, T.Name) Then Exit Function
If Not T.Indexes(T.Name).Unique Then Thw CSub, "Table has Index name same as table name, but it is not unique", "Tbn", T.Name
HasSkTd = True
End Function
