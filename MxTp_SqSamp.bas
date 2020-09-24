Attribute VB_Name = "MxTp_SqSamp"
Option Compare Text
Option Explicit
Const CMod$ = "MxTp_SqSamp."
#If False Then

Function sampSqLLny() As LLn()
Dim O$()
'PushI O, LLn(1, "sel ?MbrCnt RecCnt TxCnt Qty Amt")
PushI O, "into #Cnt"
PushI O, "fm   #Tx"
PushI O, "wh   RecCnt bet @XX @XX"
PushI O, "and  RecCnt bet @XX @XX"

PushI O, "$"
PushI O, "?MbrCnt ?Count(Distinct Mbr)"
PushI O, "RecCnt  Count(*)"
PushI O, "TxCnt   Sum(TxCnt)"
PushI O, "Qty     Sum(Qty)"
PushI O, "Amt     Sum(Amt)"
'SampSqLLn = DoLLn(DyoLLnLy(O))
End Function
#End If
