Attribute VB_Name = "MxVb_Str_InsSsub"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_InsSsub."

Function StrInsAt$(S, At, By$)
If At <= 0 Then StrInsAt = S: Exit Function
StrInsAt = Left(S, At - 1) & By & Mid(S, At)
End Function
Function StrInsAtCrLf$(S, At): StrInsAtCrLf = StrInsAt(S, At, vbCrLf): End Function
