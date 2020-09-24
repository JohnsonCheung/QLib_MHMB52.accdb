Attribute VB_Name = "MxDao_Sql_Qp"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Sql_Qp."
Const KwBet$ = "between"
Const KwSet$ = "set"
Const KwDis$ = "distinct"
Const KwUpd$ = "update"
Const KwInto$ = "into"
Const KwSel$ = "select"
Const KwFm$ = "from"
Const KwAs$ = "as"
Const KwGp$ = "group"
Const KwBy$ = "by"
Const KwWh$ = "where"
Const KwAnd$ = "and"
Const KwOn$ = "on"
Const KwJn$ = "join"
Const KwInr$ = "inner"
Const KwOr$ = "or"
Const KwOrd$ = "order"
Const KwLeft$ = "left"
Public Const C_Upd$ = KwUpd & " "
Public Const C_On$ = " " & KwOn & " "
Public Const C_Sel$ = KwSel & " "
Public Const C_SelDis$ = KwSel & " " & KwDis & " "
Public Const C_As$ = " " & KwAs & " "
Public Const C_Bet$ = " " & KwBet & " "
Public Const C_Ord$ = " " & KwOrd & " "
Public Const C_And$ = " " & KwAnd & " "
Public Const C_IJn$ = " " & KwInr & " " & KwJn & " "
Public Const C_LJn$ = " " & KwLeft & " " & KwJn & " "
Public Const C_Fm$ = " " & KwFm & " "
Public Const C_Set$ = " " & KwSet & " "
Public Const C_IsDis$ = " " & KwDis & " "
Public Const C_Gp$ = " " & KwGp & " "
Public Const C_Wh$ = " " & KwWh & " "
Public Const C_Into$ = " " & KwInto & " "
Type FldMap: Extn As String: Intn As String: End Type ' Deriving(Ctor Ay)

Function QpFm$(T, Optional Alias$):    QpFm = C_Fm & QuoSq(T) & AsIf(Alias): End Function
Function QpInsT$(T):                 QpInsT = "Insert into [" & T & "]":     End Function
Function QpBktFf$(FF$):             QpBktFf = QuoBkt(QpFf(FF)):              End Function
Function FmtVblyEpr(VblyEpr$(), Optional Pfx$, Optional IdentOpt%, Optional Sep$ = ",") As String()
Ass IsVbly(VblyEpr)
Dim Ident%
    If IdentOpt > 0 Then
        Ident = IdentOpt
    Else
        Ident = 0
    End If
    If Ident = 0 Then
        If Pfx <> "" Then
            Ident = Len(Pfx)
        End If
    End If
Dim O$(), P$, S$, U&, J&
U = UB(VblyEpr)
Dim W%
'    W = VblWdty(VblyEpr)
For J = 0 To U
    If J = 0 Then P = Pfx Else P = ""
    If J = U Then S = "" Else S = Sep
'    Push O, VblAli(VblyEpr(J), IdentOpt:=Ident, Pfx:=P, WdtOpt:=W, Sfx:=S)
Next
FmtVblyEpr = O
End Function

Function QpBktVy$(SqlVy)
Dim O$()
Dim V: For Each V In SqlVy
    PushI O, QuoSqlPrim(V)
Next
QpBktVy = QuoBkt(JnCma(O))
End Function

Function QpAnd$(Bepry$()): QpAnd = Jn(Bepry, C_And): End Function

Function QpFldXAEqLis$(TmlColon$, Alias12)
Dim X$, A$:    AsgAy SplitSpc(Alias12), X, A
Dim P As SyPair: P = S12yTmlColon(TmlColon)
QpFldXAEqLis = JnAy12(P.Sy1, P.Sy2)
End Function

Function QpOrd$(By$): QpOrd = StrPfxIfNB(C_Ord, By): End Function
Function QpOrdFfMinusSfx$(FfHypSfxOrd$)
If FfHypSfxOrd = "" Then Exit Function
Dim O$(): O = SySs(FfHypSfxOrd)
Dim I, J%
For Each I In O
    If HasSfx(O(J), "-") Then
        O(J) = RmvSfx(O(J), "-") & " desc"
    End If
    J = J + 1
Next
QpOrdFfMinusSfx = QpOrd(JnCmaSpc(O))
End Function
Function QpDis$(IsDis As Boolean): QpDis = StrTrue(IsDis, C_IsDis & " "): End Function


Private Function X_QpXAIJn(TblX$, TblA$, TmlJn$): X_QpXAIJn = FmtQQ(" [?] x inner join [?] a on ?", TblX, TblA, QpOn(TmlJn)): End Function




Function QpIntoSelStar$(Into): QpIntoSelStar = "Select * Into [" & Into & "]": End Function
Function QpInto$(T):                  QpInto = C_Into & "[" & T & "]":         End Function
Function QpIntoFm$(T, Fm):          QpIntoFm = QpInto(T) & QpFm(Fm):           End Function

Function QpAndFeq$(F$, Eqval, Optional Alias$):  QpAndFeq = vbCrLf & " and " & BeprFeq(F, Eqval, Alias): End Function
Function QpInsInto$(T):                         QpInsInto = "Insert Into [" & T & "]":                   End Function
Function QpValues$(Dr):                          QpValues = JnCma(QuoSqlPrim(Dr)):                       End Function
Function QpFf$(FF$, Optional Alias$)
Dim A$(): A = FnyFF(FF)
If Alias <> "" Then A = AmAddPfx(A, Alias & ".")
QpFf = JnCmaSpc(A)
End Function
Function QpFfAs$(FF$, DiEpr As Dictionary)
Dim UFny%, Fny$()
    Fny = FnyFF(FF)
    UFny = UB(Fny)
Dim SyFldAs$() '(UFny%,Fny$(),EpryO$()
    ReDim SyFldAs(UFny)
    Dim J%: For J = 0 To UFny
        Dim F$: F = Fny(J)
        If DiEpr.Exist(F) Then
            PushI SyFldAs, DiEpr(F) & " as " & QuoSqlF(F)
        Else
            PushI SyFldAs, QuoSqlF(F)
        End If
    Next
QpFfAs = " " & JnCmaSpc(SyFldAs) & " "
End Function
Function FldMap(Extn, Intn) As FldMap
With FldMap
    .Extn = Extn
    .Intn = Intn
End With
End Function
Function AddFldMap(A As FldMap, B As FldMap) As FldMap(): PushFldMap AddFldMap, A: PushFldMap AddFldMap, B: End Function
Sub PushFldMapAy(O() As FldMap, A() As FldMap): Dim J&: For J = 0 To FldMapUB(A): PushFldMap O, A(J): Next: End Sub
Sub PushFldMap(O() As FldMap, M As FldMap): Dim N&: N = FldMapSI(O): ReDim Preserve O(N): O(N) = M: End Sub
Function FldMapSI&(A() As FldMap): On Error Resume Next: FldMapSI = UBound(A) + 1:   End Function
Function FldMapUB&(A() As FldMap): FldMapUB = FldMapSI(A) - 1: End Function

Function S12yTmlColon(TmlColon$) As SyPair
Dim S1$(), S2$()
Dim S: For Each S In ItrTml(TmlColon)
    With BrkBoth(S, ":")
    PushI S1, .S1
    PushI S2, .S2
    End With
Next
S12yTmlColon.Sy1 = S1
S12yTmlColon.Sy2 = S2
End Function
