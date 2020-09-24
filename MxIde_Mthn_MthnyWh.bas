Attribute VB_Name = "MxIde_Mthn_MthnyWh"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthn_MthnyWh."

Private Function WMthnyRxAy(C As VBComponent, Rxay() As RegExp) As String()
Dim S$(): S = SrcCmp(C): If Si(S) = 0 Then Exit Function
WMthnyRxAy = AwRxAyAnd(MthnyPub(S), Rxay)
End Function
Function MthnWhMdy(Src$(), M As eWhMdy) As String()
Select Case True
Case M = eWhMdyAll: MthnWhMdy = Mthny(Src)
Case M = eWhMdyPub: MthnWhMdy = MthnyPub(Src)
Case M = eWhMdyPrv: MthnWhMdy = MthnyPrv(Src)
Case Else: ThwEnm CSub, M, EnmmMdy
End Select
End Function
