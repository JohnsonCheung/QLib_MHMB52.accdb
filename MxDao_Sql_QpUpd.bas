Attribute VB_Name = "MxDao_Sql_QpUpd"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Sql_QpUpd."
Function QpSetEpr$(Fny$(), Epry$())
Dim F$(): F = AmQuoSq(Fny)
QpSetEpr = QpSetX(QpEqLisEpr(Fny, Epry))
End Function

Function QpSetX$(SetX$): QpSetX = C_Set & SetX: End Function

Function QpSetXA(FnyX$(), FnyA$())
Dim X$(): X = AmAddPfxSfx(FnyX, "x.[", "]")
Dim A$(): A = AmAddPfxSfx(FnyA, "a.[", "]")
Dim J$(): J = FmtAy12(X, A, " = ")
          J = AmAddPfx(J, " ")
Dim S$:   S = Jn(J, "," & " ")
QpSetXA = QpSetX(S)
End Function

Function QpSetXATml(TmlSet$)
Dim S As SyPair: S = S12yTmlColon(TmlSet)
QpSetXATml = QpSetXA(S.Sy1, S.Sy2)
End Function

Function QpUpd$(T):                       QpUpd = C_Upd & QuoSqlT(T):                End Function
Function QpUpdX$(TblX$):                 QpUpdX = C_Upd & TblX:                      End Function
Function QpUpdXA$(X$, A$, TmlColonJn$): QpUpdXA = QpUpdX(QpTblXA(X, A, TmlColonJn)): End Function
Function QpTblXA(X$, A$, TmlColonJn$): Stop: End Function

Function QpSetFfEprAy$(FF$, Ey$())
Const CSub$ = CMod & "QpSetFfEprAy"
Dim Fny$(): Fny = FnyFF(FF)
Ass IsVbly(Ey)
If Si(Fny) <> Si(Ey) Then Thw CSub, "[Ff-Sz} <> [Si-Ey], where [Ff],[Ey]", Si(Fny), Si(Ey), FF, Ey
Dim AFny$()
    AFny = AmAli(Fny)
    AFny = AmStrSfx(AFny, " = ")
Dim W%
    'W = VblWdty(Ey)
Dim Ident%
    W = AyWdt(AFny)
Dim Ay$()
    Dim J%, U%, S$
    U = UB(AFny)
    For J = 0 To U
        If J = U Then
            S = ""
        Else
            S = ","
        End If
        'Push Ay, VblAli(Ey(J), Pfx:=AFny(J), IdentOpt:=Ident, WdtOpt:=W, Sfx:=S)
    Next
Dim Vbl$
    Dim Ay1$()
    Dim P$
    For J = 0 To U
        If J = 0 Then P = "|  Set" Else P = ""
'        Push Ay1, VblAli(Ay(J), Pfx:=P, IdentOpt:=6)
    Next
    Vbl = JnVBar(Ay1)
QpSetFfEprAy = Vbl
End Function

Function QpSetValFf$(FF$, Vy())
End Function
Function QpSetVal$(Fny$(), Vy())
End Function

Private Sub B_QpSetEpry()
Dim Fny$(), Vy()
Ept = LinesVbl("|  Set|" & _
"    [A xx] = 1                     ,|" & _
"    B      = '2'                   ,|" & _
"    C      = #2018-12-01 12:34:56# ")
Stop 'Fny = Tmy("[A xx] B C"): Vy = Array(1, "2", #12/1/2018 12:34:56 PM#): GoSub Tst
Exit Sub
Tst:
'    Act = QpSetEpr(Fny, Vy)
    C
    Return
End Sub

Private Sub B_QpSetFnyErpy()
Dim Fny$(), VblyEpr$()
Fny = SySs("a b c d")
Push VblyEpr, "1sdfkl|lskdfj|skldfjskldfjs dflkjsdf| sdf"
Push VblyEpr, "2sdfkl|lskdfjdf| sdf"
Push VblyEpr, "3sdfkl|fjskldfjs dflkjsdf| sdf"
Push VblyEpr, "4sf| sdf"
    Act = QpSetEpr(Fny, VblyEpr)
'Debug.Print LinesVbl(Act)
End Sub
