Attribute VB_Name = "MxVb_Str_JnQuo"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_JnQuo."


Function JnQDblComma$(Sy$())
JnQDblComma = JnCma(AmQuoDbl(Sy))
End Function

Function JnQDblSpc$(Sy$())
JnQDblSpc = JnSpc(AmQuoDbl(Sy))
End Function

Function JnQSngComma$(Sy$())
JnQSngComma = JnCma(AmQus(Sy))
End Function

Function JnQSngSpc$(Sy$())
JnQSngSpc = JnSpc(AmQus(Sy))
End Function

Function JnQSqCommaSpc$(Sy$())
JnQSqCommaSpc = JnCmaSpc(AmQuoSq(Sy))
End Function

Function JnQSqBktSpc$(Ay)
JnQSqBktSpc = JnSpc(AmQuoSq(Ay))
End Function
