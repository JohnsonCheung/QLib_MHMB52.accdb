Attribute VB_Name = "MxDao_Sql_Bepr"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Sql_Bepr."
Const QQ_F_In$ = "? in (?)"
Const QQ_F_Eq$ = "? = ?"
Function BeprSkvy$(D As Database, T, Skvy):  BeprSkvy = BeprFnyEq(FnySk(D, T), Skvy):                                  End Function
Function BeprIn$(Epr$, VyPrimIn):              BeprIn = FmtQQ(QQ_F_In, Epr, VyPrimIn):                                 End Function
Function BeprFldIn$(F, VyPrimIn):           BeprFldIn = FmtQQ(QQ_F_In, QuoSq(F), QuoBkt(JnCma(QuoSqlPrim(VyPrimIn)))): End Function
Function BeprFnyEq$(Fny$(), Eqvy, Optional Alias$)
Dim SyFeq$()
Dim J%: For J = 0 To UB(Fny)
    PushI SyFeq, BeprFeq(Fny(J), Eqvy(J), Alias)
Next
BeprFnyEq = QpAnd(SyFeq)
End Function
Function BeprFeq$(F, Eqval, Optional Alias$)
Dim Fld$: Fld = QuoSqlF(F, Alias)
Dim V$: V = QuoSqlPrim(Eqval)
BeprFeq = FmtQQ(QQ_F_Eq, Fld, V)
End Function
Function BeprFldIsBlnk$(F, Optional Alias$)
BeprFldIsBlnk = FmtQQ("Trim(Nz(?,''))=''", QuoSqlF(F, Alias))
End Function
