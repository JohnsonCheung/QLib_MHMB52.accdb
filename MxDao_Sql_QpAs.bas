Attribute VB_Name = "MxDao_Sql_QpAs"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Sql_QpAs."

Function QpSelAs$(M() As FldMap)
Dim O$()
Dim J%: For J = 0 To FldMapUB(M)
    PushI O, WAs(M(J))
Next
QpSelAs = JnCmaSpc(O)
End Function

Private Function WAs$(M As FldMap)
With M
    WAs = QuoSq(.Extn) & " As " & .Intn
End With
End Function
Function QpSelX$(X$, Optional IsDis As Boolean): QpSelX = C_Sel & QpDis(IsDis) & X: End Function

Private Sub B_QpSel()
Dim Fny$(), VblyEpr$()
VblyEpr = Sy("F1-Epr", "F2-Epr   AA|BB    X|DD       Y", "F3-Epr  x")
Fny = SplitSpc("F1 F2 F3xxxxx")
'Debug.Print LinesVbl(QpSelFfFldLvs(Fny, VblyEpr))
End Sub

Private Sub B_QpSelFnyExtny()
Dim Fny$()
Dim Extny$()
GoSub Z
Exit Sub
Z:
    Fny = SySs("Sku CurRateAc VdtFm VdtTo HKD Per CA_Uom")
    Extny = Tmy("Sku [     Amount] [Valid From] [Valid to] Unit per Uom")
    Debug.Print QpSelFnyExtny(Fny, Extny)
    Return
End Sub

Function QpSelDisTF$(T$, F$, Optional IsDis As Boolean):            QpSelDisTF = QpSelDis(IsDis) & "[" & F & "] From [" & T & "]": End Function
Function QpSelDis$(IsDis As Boolean):                                 QpSelDis = C_Sel & QpDis(IsDis):                             End Function
Function QpSelStar$():                                               QpSelStar = C_Sel & "*":                                      End Function
Function QpSelF$(F, Optional IsDis As Boolean):                         QpSelF = C_Sel & QpDis(IsDis) & QuoSq(F):                  End Function
Function QpSelFf$(FF$, Optional IsDis As Boolean):                     QpSelFf = QpSelFny(FnyFF(FF), IsDis):                       End Function
Function QpSelFfExtny$(FF$, Extny$(), Optional IsDis As Boolean): QpSelFfExtny = QpSelX(QpSelFnyExtny(FnyFF(FF), Extny)):          End Function
Function QpSelFny$(Fny$(), Optional IsDis As Boolean):                QpSelFny = C_Sel & QpDis(IsDis) & JnCmaSpc(Fny):             End Function
Function QpSelFnyExtny$(Fny$(), Extny$(), Optional IsDis As Boolean)
Dim O$(), J%, E$, F$
For J = 0 To UB(Fny)
    F = Fny(J)
    E = Trim(Extny(J))
    Select Case True
    Case E = "", E = F: PushI O, F
    Case Else: PushI O, QuoSq(E) & " As " & F
    End Select
Next
QpSelFnyExtny = QpSelDis(IsDis) & JnCmaSpc(O)
End Function
