Attribute VB_Name = "MxXls_Lo_ParChd_Ini"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Lo_ParChd_Ini."

Private Sub B_IniLoParChd()
GoSub Z
Exit Sub
Dim L As ListObject, FfPar$
Exit Sub
Z:
    ClsWbAllNoSav
    IniLoParChd L, FfPar
    Return
End Sub
Sub IniLoParChd(L As ListObject, FfPar$)
'A1Lo(@L) must be A1
'Insert WsParChd, which have 2 Lo: LoPar & LoChd
'Insert Worksheet_SelectionChange Cdl
Const CSub$ = CMod & "IniLoParChd"
If A1Lo(L).Address <> "A1" Then Thw CSub, "Given A1-@Lo should be A1", "@L-A1-Adr", A1Lo(L).Address
Dim FnyPar$(): Stop 'FnyPar = Tml(FfPar)
Dim Fny$(): Fny = FnyLo(L)
Dim FnyChd$(): FnyChd = SyMinus(Fny, FnyPar)
Dim SqPar(): SqPar = WSqPar(L, FnyPar)
Dim SqChd(): SqChd = WSqChd(L, FnyChd)
Dim WsParChd As Worksheet: Set WsParChd = WWsParChd(L)
Dim LoPar As ListObject: Set LoPar = WLoPar(WsParChd, SqPar)
                                     WFmtLoParChd LoPar, L
Dim LoChd As ListObject: Set LoChd = WLoChd(WsParChd, SqChd, UBound(SqPar, 2))
                                     WFmtLoParChd LoChd, L
                                     AddCdlWs WsParChd, WCdlWsSelChg
End Sub
Private Function WWsParChd(L As ListObject) As Worksheet

End Function
Private Function WSqPar(L As ListObject, FnyPar$()) As Variant()

End Function
Private Function WSqChd(L As ListObject, FnyChd$()) As Variant()

End Function
Private Function WLoPar(WsParChd As Worksheet, Sq()) As ListObject

End Function
Private Function WLoChd(WsParChd As Worksheet, Sq(), NDcPar&) As ListObject

End Function
Private Sub WFmtLoParChd(LoParChd As ListObject, L As ListObject)

End Sub
Private Function WCdlWsSelChg$()

End Function
