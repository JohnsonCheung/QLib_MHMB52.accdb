Attribute VB_Name = "MxDao_Sql_QpJn"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Sql_QpJn."

Function QpJnInr$(T$, FF$, Optional AliasX$, Optional AliasA$):              QpJnInr = QpJn(C_IJn, T, FF, AliasX, AliasA): End Function
Function QpJnLeft$(T, FF$, Optional AliasX$ = "x", Optional AliasA$ = "a"): QpJnLeft = QpJn(C_IJn, T, FF, AliasX, AliasA): End Function

Private Function QpJn(KwJn$, T, FF$, AliasX$, AliasA$)
Dim X$, A$, O$()
Dim F: For Each F In FnyFF(FF)
    X = QuoSqlT(F, AliasX)
    A = QuoSqlT(F, AliasA)
    PushI O, FmtQQ("? = ?", A, X)
Next
Dim TT$: TT = QuoSpc(QuoSq(T)) & AliasA & " "
QpJn = " " & KwJn & TT & QpAnd(O) & ")"
End Function

Function QpOn(TmlJn$, Optional Alias12$ = "x a")
QpOn = C_On & QpFldXAEqLis(TmlJn, Alias12)
End Function

Private Sub B_QpOn()
GoSub T1
Exit Sub
Dim JnTml$, Alias12$
T1:
    JnTml = "A B C D|E"
    Alias12 = "x a"
    Ept = " on x.A = a.A and x.B = a.B and x.C = a.C and x.D = a.E"
    GoTo Tst
Tst:
    Act = QpOn(JnTml, Alias12)
    C
    Return
End Sub
