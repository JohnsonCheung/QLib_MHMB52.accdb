Attribute VB_Name = "MxDao_Fun_RseqFny"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_RseqFny."

Function RseqFnyEnd(Fny$(), EndFny$()) As String()
Dim OkEnd$(): OkEnd = AyIntersect(Fny, EndFny)
RseqFnyEnd = SyAdd(SyMinus(Fny, OkEnd), OkEnd)
End Function

Function RseqFnyFront(Fny$(), FrontFny$()) As String()
Dim Front$(): Front = AyIntersect(FrontFny, Fny)
RseqFnyFront = SyAdd(Front, SyMinus(Fny, Front))
End Function
