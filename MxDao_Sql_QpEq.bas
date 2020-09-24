Attribute VB_Name = "MxDao_Sql_QpEq"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Sql_QpEq."
Function QpFeqVy(Fny$(), Vy(), Optional Alias$): QpFeqVy = QpEqLisEpr(Fny, W2EpryVy(Vy), Alias): End Function
Private Function W2EpryVy(Vy()) As String()
Dim V: For Each V In Vy
    PushI W2EpryVy, QuoSqlPrim(V)
Next
End Function
Function QpEqLisEpr(Fny$(), Epry$(), Optional Alias$)
Dim O$()
Dim J%: For J = 0 To UB(Fny)
    PushI O, QpFeq(Fny(J), Epry(J), Alias)
Next
QpEqLisEpr = JnCmaSpc(O)
End Function

Function QpFeq$(F, Epr$, Optional Alias$): QpFeq = QuoSqlF(F, Alias) & " = " & Epr: End Function
