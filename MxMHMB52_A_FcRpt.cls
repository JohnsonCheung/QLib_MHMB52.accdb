VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_FcRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
Const CMod$ = "MxMHMB52_A_FcRpt."
Sub RptFcLas(): RptFcYm MH.FcLas.Ym: End Sub
Sub RptFcYm(A As Ym)
WOup A
WFxo A
StsDone
End Sub
Private Sub WOup(A As Ym)
MH.FcTmp7.CrtYm A
MH.PhTmp5.Crt
MH.Ph7Tb.CpyAddAtr "$Fc?", "@Fc?"
MH.Ph7Tb.RplStmByStream "@Fc?"
MH.Ph7Tb.Rseq "@Fc?"
MH.FcTmp7.Drp
End Sub
Private Sub WFxo(A As Ym)
Dim Fxo$: Fxo = MH.FcIO.Fxo(A): If AskOpnFx(Fxo) Then Exit Sub
DltFfnIf Fxo
MH.FcTp.Exp Fxo
Dim B As Workbook: Set B = WbFx(Fxo)
RfhLoWbDaC B
EnsIxws B
B.Save
Maxv B.Application
End Sub
