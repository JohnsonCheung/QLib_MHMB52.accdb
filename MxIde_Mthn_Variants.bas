Attribute VB_Name = "MxIde_Mthn_Variants"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthn_Variants."


Function MthnPub$(L): MthnPub = MthnMay(L, "", ""):       End Function
Function SubnPub$(L): SubnPub = MthnMay(L, "", "Sub"):    End Function
Function FunnPub$(L): FunnPub = MthnMay(L, "", "Fun"):    End Function
Function MthnPrv$(L): MthnPrv = MthnMay(L, "Prv", ""):    End Function
Function SubnPrv$(L): SubnPrv = MthnMay(L, "Prv", "Sub"): End Function
Function FunnPrv$(L): FunnPrv = MthnMay(L, "Prv", "Fun"): End Function
Private Function MthnMay$(L, ShtMdy$, ShtTy$)
With TMthL(L)
    If .Mthn = "" Then Exit Function
    If ShtMdy <> .ShtMdy Then Exit Function
    If ShtTy = "" Then
        MthnMay = .Mthn
    Else
        If .ShtTy = ShtTy Then MthnMay = .Mthn
    End If
End With
End Function
Function PrpnyCmp(A As VBComponent) As String(): PrpnyCmp = Itn(A.Properties): End Function
