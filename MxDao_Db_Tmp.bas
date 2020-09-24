Attribute VB_Name = "MxDao_Db_Tmp"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Db_Tmp."

Function IsDbTmp(D As Database) As Boolean: IsDbTmp = PthDb(D) = PthTmpDb: End Function
Sub DrpDbTmp(D As Database)
If IsDbTmp(D) Then DrpDb D
End Sub
Sub DrpDb(D As Database)
Dim N$
N = D.Name
D.Close
DltFfn N
End Sub
Function PthTmpDb$(): PthTmpDb = PthEns(PthTmp & "DbTmp\"): End Function
Sub BrwDbTmpLas():               BrwDb DbTmpLas:            End Sub

Function DbTmpLas() As Database: Set DbTmpLas = Db(FbTmpLas): End Function
Function FbTmpLas$()
Const CSub$ = CMod & "FbTmpLas"
Dim P$: P = PthTmpDb
Dim Fn$: Fn = EleMax(Fnay(P, "*.accdb"))
If Fn = "" Then Thw CSub, "No *.accdb PthTmpDb", "PthTmpDb", PthTmpDb
FbTmpLas = PthTmpDb & Fn
End Function

Function DbTmp(Optional Pfx$ = "N") As Database: Set DbTmp = CrtFb(FbTmp(Pfx)):               End Function
Function FbTmp$(Optional Pfx$ = "N"):                FbTmp = PthTmpDb & Tmpn(Pfx) & ".accdb": End Function
