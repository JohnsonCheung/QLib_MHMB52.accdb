Attribute VB_Name = "MxXls_ParChd_Lo"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_ParChd_Lo."
Sub PutPChdLo(LoSrc As ListObject, Gpcc$, At As Range)
Dim FnyGp$(): FnyGp = SySs(Gpcc)
Dim FnySrc$(): FnySrc = FnyLo(LoSrc)
Dim AtChd As Range: Set AtChd = WAtChd(At, FnySrc, FnyGp)
LoNwSq At, WSqPar(LoSrc, FnyGp)
LoNwSq AtChd, WSqChd(LoSrc, FnyGp)
AddCdlWs WsLo(LoSrc), WSrcl
End Sub

Private Function WSqPar(LoSrc As ListObject, GpFny$()) As Variant()

End Function
Private Function WSqChd(LoSrc As ListObject, GpFny$()) As Variant()

End Function
Private Function WAtChd(At As Range, FnySrc$(), FnyGp$()) As Range

End Function
Private Function WSrcl$()

End Function
