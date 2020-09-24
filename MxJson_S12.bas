Attribute VB_Name = "MxJson_S12"
Option Compare Text
Option Explicit
Const CMod$ = "MxJson_S12."
Function S12Json(JsonS12$) As S12
'Dim L$: L = TrimWhite(JsonS12)
'Stop 'With StroptzRmvBktBig(L)
''    If Not .Som Then Thw CSub, "Given JsonS12 is not StrObj", "JsonS12", JsonS12
''End With
'Dim PrpnL$: Stop ' PrpnL = PrpnJsonShf(L, "S1 S2"): If Not HasEle(Ay, PrpnL) Then Thw CSub, "Given JsonS12 does not has prpn S1 or S2", "JsonS12", JsonS12
'Stop 'With OptJstrzShf(L)
'    If Not .Som Then
'        Thw CSub, "Expected Jstr", "StrWrk JsonS12", StrWrk, JsonS12
'    End If
'    Dim Str$: Str = 1
''End With
End Function
Function JsonS12$(S As S12)
End Function

Function OptJstrzShf$(OStrWrk$)
'OStrWrk = WhiteTrim
End Function
