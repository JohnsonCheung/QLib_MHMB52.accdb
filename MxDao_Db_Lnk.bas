Attribute VB_Name = "MxDao_Db_Lnk"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Db_Lnk."
Type TLnkTblFx: T As String: Fx As String: WsnFx As String: TbnAs As String: End Type ' Deriving(Ctor Ay)
Type TLnkTblFb: T As String: Fb As String: TbnFb As String: TbnAs As String: End Type ' Deriving(Ctor Ay)
Type TLnkTbl: Fx() As TLnkTblFx: Fb() As TLnkTblFb: End Type
Sub LnkTLnkTbl(D As Database, U As TLnkTbl)
Dim J%
For J = 0 To UbTLnkTblFx(U.Fx): With U.Fx(J): LnkFxw D, .TbnAs, .Fx, .WsnFx: End With: Next
For J = 0 To UbTLnkTblFb(U.Fb): With U.Fb(J): LnkFbt D, .TbnAs, .Fb, .TbnFb: End With: Next
End Sub
Sub LnkFxwC(T, Fx, Optional WsnFx$ = "Sheet1"):               LnkFxw CDb, T, Fx, WsnFx:                    End Sub
Sub LnkFbtC(T, Fb, Optional TbnFb$):                          LnkFbt CDb, T, Fb, TbnFb:                    End Sub
Sub LnkFxw(D As Database, T, Fx, Optional WsnFx$ = "Sheet1"): LnkTbl D, T, WsnFx & "$", CnsFxDao(Fx):      End Sub
Sub LnkFbt(D As Database, T, Fb, Optional TbnFb):             LnkTbl D, T, StrDft(TbnFb, T), CnsFbDao(Fb): End Sub
Sub LnkTbl(D As Database, T, TbnSrc$, Cn$) ' Crt Tb-T as Lnk Tbl with @S::SrcTbn & @Cn::Cns
Const CSub$ = CMod & "LnkTbl"
On Error GoTo X
Drp D, T
Dim Td As New Dao.TableDef
    With Td
        .Connect = Cn
        .Name = T
        .SourceTableName = TbnSrc
    End With
D.TableDefs.Append Td
Exit Sub
X:
    Dim Er$: Er = Err.Description
    Thw CSub, "Error in linking table", "Er Db T TbnSrc Cn", Er, D.Name, T, TbnSrc, Cn
End Sub

Sub TLnkTblFxPush(O() As TLnkTblFx, M As TLnkTblFx): Dim N&: N = SiTLnkTblFx(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiTLnkTblFx&(A() As TLnkTblFx): On Error Resume Next: SiTLnkTblFx = UBound(A) + 1: End Function
Function UbTLnkTblFx&(A() As TLnkTblFx): UbTLnkTblFx = SiTLnkTblFx(A) - 1: End Function
Function TLnkTblFx(T, Fx, WsnFx, TbnAs) As TLnkTblFx
With TLnkTblFx
    .T = T
    .Fx = Fx
    .WsnFx = WsnFx
    .TbnAs = TbnAs
End With
End Function
Sub PushTLnkTblFx(O() As TLnkTblFx, M As TLnkTblFx): Dim N&: N = SiTLnkTblFx(O): ReDim Preserve O(N): O(N) = M: End Sub
Sub PushTLnkTblFb(O() As TLnkTblFb, M As TLnkTblFb): Dim N&: N = SiTLnkTblFb(O): ReDim Preserve O(N): O(N) = M: End Sub
Function SiTLnkTblFb&(A() As TLnkTblFb): On Error Resume Next: SiTLnkTblFb = UBound(A) + 1: End Function
Function UbTLnkTblFb&(A() As TLnkTblFb): UbTLnkTblFb = SiTLnkTblFb(A) - 1: End Function
Function TLnkTblFb(T, Fb, TbnFb, TbnAs) As TLnkTblFb
With TLnkTblFb
    .T = T
    .Fb = Fb
    .TbnFb = TbnFb
    .TbnAs = TbnAs
End With
End Function
