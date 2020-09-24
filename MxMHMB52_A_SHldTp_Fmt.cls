VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_SHldTp_Fmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_SHldTp_Fmt."
Private Const WdtFcSC% = 8
Private Const WdtFcRemSC% = 8
Private Const WdtFcStkDays% = 5
Private Const WdtFcM15% = 6

Sub Fmt()
Dim B As Workbook: Set B = WbFx(MH.SHldTp.Tp)
MH.SHldFmtDteTit15Mth.Fmt B, Ym(19, 5)  'At B-Lvl  'This is needed to be called each time generating the report
'--- no need to call when gen rpt
B.Application.WindowState = xlMinimized
B.Application.Visible = True
Stop
WFmtKeyCnt B                'At B-Lvl
WFmtSHldWb B    'StkHld
WFmtFcWb B    'Forecast
WFmtStkDaysWb B    'StkDays
SetWbOutLinSum B
MiniWbOLvl B
'SetWbNoAutoColWdt B
Stop
B.Save
B.Application.WindowState = xlMaximized
End Sub

Private Sub B_WFmtFcWb(): WFmtFcWb X_ShWs: End Sub
Private Sub WFmtFcWb(B As Workbook)
Dim Ws As Worksheet: For Each Ws In B.Sheets
    If MH.SHldWsn.IsFc(Ws.Name) Then WFmtFcWs Ws
Next
End Sub
Private Sub WFmtFcWs(S As Worksheet)
S.Activate
Dim Lo As ListObject: Set Lo = LoFst(S)
Lcc_Wdt(Lo, FcSSM15) = WdtFcM15
Lc_Wdt(Lo, "SC") = WdtFcSC
Lc_Wdt(Lo, "StkDays") = WdtFcStkDays
Lc_Wdt(Lo, "RemSC") = WdtFcRemSC
SetLcAsSum Lo, "SC RemSC"
SetLcFmt Lo, FcSSM15, "#,###"
SetLcFmt Lo, "SC StkDays RemSC", "#,###"
End Sub

Private Sub B_WFmtStkDaysWb(): WFmtStkDaysWb X_WbTpTmp: End Sub
Private Sub WFmtStkDaysWb(B As Workbook)
'#Fmt-StkDays-Ws-es#
Dim Ws As Worksheet: For Each Ws In B.Sheets
    If MH.SHldWsn.IsSd(Ws.Name) Then WFmtStkDaysWs Ws
Next
End Sub
Private Sub WFmtStkDaysWs(S As Worksheet)
S.Activate
Dim L As ListObject: Set L = LoFst(S)
Lcc_Wdt(L, SdSSStkDays) = 6
Lcc_Wdt(L, SdSSRemSc) = 6
Lc_Wdt(L, "SC") = 7
SetLcAsSum L, SdSSRemSc
SetLcFmt L, SdSSRemSc, "#,###"
SetLcFmt L, SdSSStkDays, "#,###"
End Sub

Private Sub B_WFmtSHldWb(): WFmtSHldWb X_WbTpTmp: End Sub
Private Sub WFmtSHldWb(B As Workbook)
Dim Ws As Worksheet: For Each Ws In B.Sheets
    If MH.SHldWsn.IsSh(Ws.Name) Then WFmtWsSHld Ws
Next
End Sub
Private Sub WFmtWsSHld(S As Worksheet)
S.Activate
Dim Lo As ListObject: Set Lo = LoFst(S)
Lcc_Wdt(Lo, "F1 F2") = 0.5
Lcc_Wdt(Lo, VnnnSHldHkd) = 8
Lcc_Wdt(Lo, VnnnSHldSc) = 7
Lcc_Wdt(Lo, FfSHldKpi) = 5
SetLcAsSum Lo, VnnnSHldSc
SetLcAsSum Lo, VnnnSHldHkd
SetLcFmt Lo, VnnnSHldHkd, "#,###K"
SetLcFmt Lo, VnnnSHldSc, "#,###"
End Sub

Private Sub B_WFmtKeyCnt()
Dim B As Workbook: Set B = WbFx(MH.SHldTp.Tp)
WFmtKeyCnt B
Maxv B.Application
End Sub
Private Sub WFmtKeyCnt(B As Workbook)
'Each SHRptWs using the Wsn to find the Key Column, set this KeyCol.Calc=Cnt
Dim S As Worksheet: For Each S In B.Sheets
    If MH.SHldWsn.IsSHld(S.Name) Then
        SetLcAsCnt LoFst(S), WColnKey(S.Name)
    End If
Next
End Sub
Private Function WColnKey$(RptWsn$)
':RptWsn: :Nm ! #Rpt-Ws-Nm# must in format of [S3 Ph7], where S3 = SktHld | Fc | StkDays & Ph7 is one of Ph7Ss
Dim O$
Select Case AftSpc(RptWsn)
Case "Sku": O = "Sku"
Case "L4": O = "PHQly"
Case "L3": O = "PHQGp"
Case "L2": O = "PHBrd"
Case "L1": O = "PHNam"
Case "Bus": O = "PHBus"
Case "Stm": O = "Stream"
Case Else: Stop
End Select
WColnKey = O
End Function

Private Function ShSS$():               ShSS = "*Key ScCsg HkdCsg .. ScTot HkdTot StkDays StkMths RemSC TarStkMths": End Function
Private Function VnnnSHldHkd$(): VnnnSHldHkd = NnQtp2("Hkd?", X_1Vn5):                                               End Function
Private Function VnnnSHldSc$():   VnnnSHldSc = NnQtp2("Sc?", X_1Vn5):                                                End Function
Private Function FfSHldKpi$():     FfSHldKpi = "StkDays StkMths RemSC TarStkMths":                                   End Function
Private Function X_1Vn5$():           X_1Vn5 = "Csg Df Dp Git Tot":                                                  End Function

Private Function FcSS$():       FcSS = "*Key SC StkDays RemSC *M15": End Function
Private Function FcSSM15$(): FcSSM15 = JnPfxSno("M", 1, 15, "00"):   End Function

Private Function SdSS$():               SdSS = "*Key SC *StkDays/RemSC15":       End Function
Private Function SdSSRemSc$():     SdSSRemSc = JnPfxSno("RemSC", 1, 15, "00"):   End Function
Private Function SdSSStkDays$(): SdSSStkDays = JnPfxSno("StkDays", 1, 15, "00"): End Function

Private Function X_ShWs() As Worksheet: Set X_ShWs = X_WbTpTmp.Sheets("StkHld Stm"): End Function

Private Function X_WbTpTmp() As Workbook
Dim F$: F = FxTmp
CpyFfn MH.MB52Tp.Tp, F
Set X_WbTpTmp = WbFx(F)
Maxv X_WbTpTmp.Application
End Function
