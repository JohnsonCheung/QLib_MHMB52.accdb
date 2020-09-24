Attribute VB_Name = "MxDao_Def_TdStr"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Def_TdStr."

Function SpTd$(T As Dao.TableDef)
Dim Tbn$, Id$, S$, R$
    Tbn = T.Name
    If HasPPkTd(T) Then Id = "*Id"
    Dim Pk$(): Pk = Sy(Tbn & "Id")
    Dim Sk$(): Sk = FnySkTd(T)
    If HasSkTd(T) Then S = TmlAy(AmRpl(Sk, Tbn, "*")) & " |"
    R = TmlAy(CvSy(AyMinusAp(FnyTd(T), Pk, Sk)))
SpTd = JnSpc(SyNB(T.Name, Id, S, R))
End Function

Function Tdrep$(D As Database, T): Tdrep = SpTd(D.TableDefs(T)): End Function
Function HasPPk(T As Dao.TableDef) As Boolean
'If Not HasPk(A) Then Exit Function
Dim Pk$(): Pk = FnyPkTd(T): If Si(Pk) <> 1 Then Exit Function
Dim P$: P = T.Name & "Id"
If Pk(0) <> P Then Exit Function
HasPPk = T.Fields(0).Name <> P
End Function
