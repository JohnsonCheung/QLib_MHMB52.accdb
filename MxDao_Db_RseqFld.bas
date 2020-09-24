Attribute VB_Name = "MxDao_Db_RseqFld"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Db_RseqFld."

Sub CrtTboRseq(T, Nseq$(), OrdBy$)
Dim F$: F = Join(AmQuoSq(Nseq), ",")
Dim SelF$: SelF = "Select " & F
Dim Into$: Into = " Into [@" & T & "]"
Dim Fm$: Fm = " From [#" & T & "]"
Dim Sql$: Sql = SelF & Into & Fm & " Order By " & OrdBy
RunqC Sql
End Sub

Function IsLnkTbl(D As Database, T) As Boolean: IsLnkTbl = Td(D, T).Connect <> "": End Function
Sub RseqFldC(T, FF$):                                      RseqFld CDb, T, FF:     End Sub
Sub RseqFld(D As Database, T, FF$)
'@Ff can be partial of @T->Fny, but all fields should exist in @T->Fny
Const CSub$ = CMod & "RseqFld"
If IsLnkTbl(D, T) Then Thw CSub, "Given table is a linked, cannot Rseq", "T", T
Dim BefF$(): BefF = Fny(D, T)
Dim NewF$():
    Dim GivF$(): GivF = FnyFF(FF)
    Dim ExcessF$(): ExcessF = AyMinus(GivF, BefF)
    If Si(ExcessF) > 0 Then Thw CSub, "Given Ff some excess field than given T", "Excess-field Given-Ff T-Ff T", TmlAy(ExcessF), FF, TmlAy(BefF), T
    Dim MisF$(): MisF = AyMinus(BefF, GivF)
    NewF = SyAdd(GivF, MisF)

Dim Td1 As Dao.TableDef: Set Td1 = Td(D, T)
WRseq Td1, NewF
WResetOrdinalPosition Td1
WVerify Td1, NewF, BefF, FF, D
End Sub
Private Sub WRseq(T As Dao.TableDef, Fny$())
Dim M%: M = MaxOrdPosTd(T)
Dim J%: For J = UB(Fny) To 0 Step -1
    T.Fields(Fny(J)).OrdinalPosition = M + 1 + J
Next
End Sub
Private Sub WResetOrdinalPosition(T As Dao.TableDef)
Dim J%: J = 0
Dim F As Dao.Field: For Each F In T.Fields
    J = J + 1
    F.OrdinalPosition = J
Next
End Sub

Private Sub WVerify(T As Dao.TableDef, NewFny$(), BefFny$(), GivenFf$, D As Database)
Const CSub$ = CMod & "WVerify"
Dim J%: For J = 0 To UB(NewFny)
    If T.Fields(J).OrdinalPosition <> J + 1 Then
        Thw CSub, "Table not reseq as expected", _
            "Given-Ff Bef-Srt-Given-Tbl-Ff Aft-Srt-Given-Tbl-Ff Aft-Srt-Given-Tbl-OrdPos", _
            GivenFf, TmlAy(BefFny), TmlAy(Itn(T.Fields)), FmtOrdPos(D, T)
    End If
Next
End Sub
