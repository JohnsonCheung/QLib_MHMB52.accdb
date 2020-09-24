Attribute VB_Name = "MxXls_Lo_Op_RfhLo"
' RfhLo using @Da-Way.  Rpl *Lo->DataBodyRange by the table in @Fx with Handling Fml..
'Assume the Lo->Name is in Format of Lo_* and there is such @* in @Fb, else throw
'Assume the Lo->Fny aft exl can all be found in @Fb->Tbl->Fny, else throw
'        where aft exl is (1) Exl Filler* (2) with Fml
'     get Data from @D->@* will not formatted in a Sq match the fields in *Lo with handling fml
'     Put data to *Lo
'     put Fml back
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Lo_Op_RfhLo."
Private Sub B_RfhLoFxDaC()
Dim T$: T = FxTmp
CpyFfn MH.FcSamp.FxRpt, T
RfhLoFxDaC T
OpnFxAp T, MH.FcSamp.FxRpt
End Sub
Sub RfhLoFxDaC(Fx$): RfhLoFxDa CDb, Fx: End Sub
Sub RfhLoFxDa(D As Database, Fx)
Dim B As Workbook: Set B = WbFx(Fx)
RfhLoWbDa D, B
B.Close True
End Sub
Sub RfhLoWbDaC(B As Workbook): RfhLoWbDa CDb, B: End Sub
Sub RfhLoWbDa(D As Database, B As Workbook)
ChkTDtaSrcNoEr Ept:=TDtaSrcLo(B), Act:=TDtaSrcTblOup(D)
Dim L() As ListObject: L = LoyTbl(B)
Dim I: For Each I In Itr(L)
    RfhLoDa D, CvLo(I)
Next
End Sub
Sub RfhLoDaC(L As ListObject): RfhLoDa CDb, L: End Sub
Sub RfhLoDa(D As Database, L As ListObject)
Dim Fmllny$(): Fmllny = FmllnyLo(L)
Dim Sq(): Sq = SqQ(D, SqlLo(L))
ResiLo L, UBound(Sq, 1)
L.DataBodyRange.Value = Sq
SetLoFmllny L, Fmllny
End Sub

Private Sub B_RfhLoFxRsC()
Dim F$: F = FxTmp
CpyFfn MH.FcSamp.FxRpt, F
RfhLoFxRsC F
OpnFxAp F, MH.FcSamp.FxRpt
End Sub
Sub RfhLoFxRsC(Fx): RfhLoFxRs CDb, Fx: End Sub ' #(Rfh)-(C)Db-@Tb-Into-all-(Lo)-of-@Fx-using-(Ars)-approach#
Sub RfhLoFxRs(D As Database, Fx):
RfhLoWbRs D, WbFx(Fx)
CWbLas.Close True
End Sub
Sub RfhLoWbRsC(B As Workbook): RfhLoWbRs CDb, B: End Sub
Sub RfhLoWbRs(D As Database, B As Workbook)
ChkTDtaSrcNoEr Ept:=TDtaSrcLo(B), Act:=TDtaSrcTblOup(D)
Dim L: For Each L In Itr(LoyTbl(B))
    RfhLoRs D, CvLo(L)
Next
End Sub
Sub RfhLoRsC(L As ListObject): RfhLoRs CDb, L: End Sub
Sub RfhLoRs(D As Database, L As ListObject)
Dim Fmllny$(): Fmllny = FmllnyLo(L)
Dim R As Dao.Recordset: Set R = Rs(D, SqlLo(L))
ResiLo L, NRecRs(R)
L.DataBodyRange.CopyFromRecordset R
SetLoFmllny L, Fmllny
End Sub

Private Sub B_SqlLo()
Dim B As Workbook: Set B = MH.FcSamp.WbRpt
Dim L As ListObject: Set L = LoWb(B, "Lo_FcStm")
Debug.Print SqlLo(L)
B.Close
End Sub
Function SqlLo$(L As ListObject)
Dim M$
Dim T$: T = TbnLo(L): If T = "" Then M = "Lon-of-@L does have valid Lon patn mapping to a Tbn.": GoTo M
Dim X$
    Dim WFnyLo$(): WFnyLo = FnyLo(L)
    Dim WFnyFml$(): WFnyFml = FnyFml(L)
    X = WQL_FldLis(WFnyLo, WFnyFml)
SqlLo = SqlSelX(T, X)
Exit Function
M:
    Const MsglVdtLon$ = "PatnLonVdt:Lo_YyyTTT where Yyy is Inp|Tmp|Oup; TTT is Tbn; Patn-Tbn is ?TTT; where ? is >|$|@ map from Yyy"
    Thw CSub, M, "Lon-of-@L [Valid Lon]", L.Name, MsglVdtLon

End Function
Private Sub WQL__SqlLo(): End Sub
Private Function WQL_FldLis$(Fny$(), FnyFiller$())
Dim O$()
Dim F: For Each F In Fny
    PushI O, WQL_FldItm(F, HasEle(FnyFiller, F))
Next
WQL_FldLis = JnCma(O)
End Function
Private Function WQL_FldItm$(F, IsFiller As Boolean)
Dim SqlItm$: SqlItm = QuoSqlF(F)
If IsFiller Then
    WQL_FldItm = "null as " & SqlItm
Else
    WQL_FldItm = SqlItm
End If
End Function
