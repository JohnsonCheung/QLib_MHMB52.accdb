Attribute VB_Name = "MxDao_Db_Sk"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Db_Sk."
Public Const Pkn$ = "PrimaryKey"

Sub InsRecSskVet(D As Database, SskTbn, ToInsSskVet As Dictionary) _
'Insert Single-Field-Secondary-Key-Aet into Dbt
'Assume T has single-fld-sk and can be inserted by just giving such SSk-value
Dim ShouldInsVet As Dictionary
    Set ShouldInsVet = AetMinus(ToInsSskVet, AetSsk(D, SskTbn))
If ShouldInsVet.Count = 0 Then Exit Sub
Dim F$: F = Sskn(D, SskTbn)
With RsTbl(D, SskTbn)
    Dim I: For Each I In ShouldInsVet
        .AddNew
        .Fields(F).Value = I
        .Update
    Next
    .Close
End With
End Sub

Function FnySkC(T) As String(): FnySkC = FnySk(CDb, T): End Function
Function FnySk(D As Database, T) As String()
Dim I As Dao.Index: Set I = IdxSk(D, T): If IsNothing(I) Then Exit Function
FnySk = Itn(I.Fields)
End Function
Function FnySkTd(T As Dao.TableDef) As String()
Dim I As Dao.Index: Set I = WIdxSk(T): If IsNothing(I) Then Exit Function
FnySkTd = Itn(I.Fields)
End Function
Private Function WIdxSk(T As Dao.TableDef) As Dao.Indexes: Set WIdxSk = ItoFstNm(T.Indexes, T.Name): End Function

Function IdxNwSk(T As Dao.TableDef, Fny$()) As Dao.Index: Set IdxNwSk = IdxNw(T, T.Name, Fny, IsUKy:=True):                    End Function
Function IdxNwPk(T As Dao.TableDef) As Dao.Index:         Set IdxNwPk = IdxNw(T, "PrimaryKey", Sy(T.Name & "id"), IsPk:=True): End Function
Function IdxNw(T As Dao.TableDef, K$, Fny$(), Optional IsPk As Boolean, Optional IsUKy As Boolean) As Dao.Index
Dim O As Dao.Index: Set O = T.CreateIndex(K)
If IsPk Then
    O.Primary = True
    O.Unique = True
ElseIf IsUKy Then
    O.Unique = True
End If
Dim Fds As Dao.IndexFields: Set Fds = O.Fields
Dim F: For Each F In Fny
    Fds.Append O.CreateField(F)
Next
Set IdxNw = O
End Function

Function Sskn$(D As Database, T)
Const CSub$ = CMod & "Sskn"
Dim Sk$(): Sk = FnySk(D, T): If Si(Sk) = 1 Then Sskn = Sk(0): Exit Function
Thw CSub, "FnySk-Si<>1", "Db T, FnySk-Si FnySk", D.Name, T, Si(Sk), Sk
End Function

Function AetSsk(D As Database, T) As Dictionary: Set AetSsk = AetF(D, T, Sskn(D, T)): End Function
