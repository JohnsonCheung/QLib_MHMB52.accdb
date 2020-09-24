Attribute VB_Name = "MxDao_UsrPrm"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_TbUsrPrm."

Function PmvSetIfC(Pmn$, V): PmvSetIfC = PmvSetIf(CDb, Pmn, V): End Function
Function PmvSetIf(D As Database, Pmn$, V):
If IsEmp(V) Then Exit Function
UpdRsV Rs(D, WWSqlPmn(Pmn)), V
PmvSetIf = V
End Function
Sub SetPmvIfC(Pmn$, V): SetPmvIf CDb, Pmn, V: End Sub
Sub SetPmvIf(D As Database, Pmn$, V):
If IsEmp(V) Then Exit Sub
UpdRsV Rs(D, WWSqlPmn(Pmn)), V
End Sub
Function PmvStrC$(Pmn$):                   PmvStrC = PmvStr(CDb, Pmn):                                       End Function
Function PmvStr$(D As Database, Pmn$):      PmvStr = Nz(Pmv(D, Pmn), ""):                                    End Function
Function Pmv(D As Database, Pmn$):             Pmv = ValQ(D, WWSqlPmn(Pmn)):                                 End Function '#Parameter-Value#
Function Pmny(D As Database) As String():     Pmny = AySrtQ(FnyRs(RsTbl(D, "UsrPrm"))):                      End Function
Function PmvC(Pmn$):                          PmvC = Pmv(CDb, Pmn):                                          End Function '#CDb-Parameter-Value: :Val
Function PmnyC() As String():                PmnyC = Pmny(CDb):                                              End Function
Private Function WWSqlPmn$(Pmn$):         WWSqlPmn = FmtQQ("Select ? from UsrPrm where Usr='?'", Pmn, CUsr): End Function

Sub TglPm(D As Database, Pmn$): PmvSetIf D, Pmn, Not Pmv(D, Pmn): End Sub
Sub TglPmC(Pmn$):               TglPm CDb, Pmn:                   End Sub
Function PmFxSelSetC$(Pmn$, Optional FxSpecDes$ = "Select a Xlsx file"): PmFxSelSet CDb, Pmn, FxSpecDes: End Function

Function PmFxSelSet$(D As Database, Pmn$, Optional FxSpecDes$ = "Select a Xlsx file")
PmFxSelSet = FxSel(Nz(Pmv(D, Pmn), ""), FxSpecDes): If PmFxSelSet <> "" Then PmvSetIf D, Pmn, PmFxSelSet
End Function
Sub SelSet_PmFx(D As Database, Pmn$, Optional FxSpecDes$ = "Select a Xlsx file")
SetPmvIf D, Pmn, FxSel(Nz(Pmv(D, Pmn), ""), FxSpecDes)
End Sub
Sub SelSet_PmFxC(Pmn$, Optional FxSpecDes$ = "Select a Xlsx file"): SelSet_PmFx CDb, Pmn, FxSpecDes: End Sub
Function PmPthSelSetC(Pmn$): PmPthSelSet CDb, Pmn: End Function
Function PmPthSelSet$(D As Database, Pmn$): PmPthSelSet = PmvSetIf(D, Pmn, PthSel(PmvStr(D, Pmn))): End Function
