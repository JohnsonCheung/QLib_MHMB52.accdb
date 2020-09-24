Attribute VB_Name = "MxDao_Sql_QpWh"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Sql_Wh."

Function WhFnyEq$(Fny$(), Eqvy):             WhFnyEq = Wh(BeprFnyEq(Fny, Eqvy)):     End Function
Function WhFeq(F$, Eqval, Optional Alias$):    WhFeq = Wh(BeprFeq(F, Eqval, Alias)): End Function
Function WhFldIn$(F$, VyIn):                 WhFldIn = Wh(BeprFldIn(F, VyIn)):       End Function
Function WhFfeq$(FF$, Eqvy):                  WhFfeq = WhFnyEq(FnyFF(FF), Eqvy):     End Function
Function WhTblEqK$(T, K&):                  WhTblEqK = WhFeq(T & "Id", K):           End Function
Function WhTblId$(T, Id):                    WhTblId = Wh(FmtQQ("[?]Id=?", T, Id)):  End Function

Function WhBet$(F$, FmV, ToV, Optional Alias$): WhBet = C_Wh & QuoSqlF(F, Alias) & C_Bet & QuoSqlT(FmV) & C_And & QuoSqlT(ToV): End Function

Function Wh$(Bepr$): Wh = StrPfxIfNB(C_Wh, Bepr): End Function

Private Sub B_WhFldIn()
Dim F$, Vy()
F = "A"
Vy = Array(1, "2", #2/1/2017#)
Ept = " where A=1 and B='2' and C=#2017-2-1#"
GoSub Tst
Exit Sub
Tst:
    Act = WhFldIn(F, Vy)
    C
    Return
End Sub
