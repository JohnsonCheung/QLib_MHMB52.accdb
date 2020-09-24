Attribute VB_Name = "MxTp_SqSqy"
Option Compare Text
Option Explicit
Const CMod$ = "MxTp_SqSqy."
#If False Then
Enum eSqBlkTy: eErBlk: eSqBlk: ePmBlk: eSwBlk: End Enum
Enum eSqSpect: eSelStmt: eRUpdStmt: eDrpStmt: End Enum
Public Const eSqSpect$ = "SelStmt UpdStmt DrpStmt"
Const KwInto$ = "INTO"
Const KwSel$ = "SEL"
Const KwSelDis$ = "SELECT DISTINCT"
Const KwFm$ = "FM"
Const KwGp$ = "GP"
Const KwWh$ = "WH"
Const KwAnd$ = "AND"
Const KwJn$ = "JN"
Const KwLeftJn$ = "LEFT JOIN"
Const SqTpKwSS$ = "Pm Sw Sql Rmk"
Type Pmsw2
    Pm As Dictionary
    stmtSw As Dictionary
    fldSw As Dictionary
End Type

Function SqyD(B As SqBld) As String()

End Function
Private Function SqlnErn$(A As LLn, B As Pmsw2)
Dim S As S12: S = BrkSpc(A.Ln)
Select Case True
Case HasPfx(S.S1, ">?")
    Select Case S.S2
    Case "0", "1"
    Case SqlnErn = LLnMsg(A, "If K='>?xxx', V should be 0 or 1.")
    End Select
Case HasPfx(S.S1, ">")
Case Else: SqlnErn = LLnMsg(A, FmtQQ("Pm line should beg with (>? | >)"))
End Select
End Function

Private Function LLnMsg$(A As LLn, Msg$)

End Function

Private Sub B_SqyTpLy()
Dim SqTpLy$()
GoSub Z
Exit Sub
Z:
    B SqyTpLy(SampSqTpLy)
    Return
End Sub

Function SqyTpLy(SqTpLy$()) As String()
Const CSub$ = CMod & "SqyTpLy"
Dim S As SqTpSrc: S = SqTpSrcT(SqTpLy)
Dim S As SqTpSrc: S = SqTpSrcT(SqTpLy)
ChkEry SqTpEr(S), CSub
SqyTpLy = UUSqy(UUBld(UUDta(S)))

Function SqlSel$(Sel$(), EDic As Dictionary, fldSw As Dictionary)
'BrwKLys Sel
Dim LFm$, LInto$, LSel$, LOrd$, LWh$, LGp$, LAndOr$(), LAlias$()
'    LSel = ShfKLyMLin(X, "Sel")
'    LInto = ShfKLyMLin(X, "Into")
'    LFm = ShfKLyMLin(X, "Fm")
'    LJn = ShfKLyMLyKK(X, "Jn LJn")
'    LWh = ShfKLyOLin(X, "Wh")
'    LAndOr = ShfKLyMLyKK(X, "And Or")
'    LGp = ShfKLyOLin(X, "Gp")
'    LOrd = ShfKLyOLin(X, "Ord")
Dim ADic As Dictionary: Set ADic = DiVkkLy(LAlias)
Dim Ffny$(), FGp$()
'    Ffny = QpSelFny(LSel, fldSw)
    FGp = QpSelF(LGp, fldSw)
Dim OX$, Into$, OT$, OWh$, OGp$, OOrd$
'    Dim Fny$()

'    '
    Into = RmvA1T(LInto)
Stop '    OGp = QpEprLis(FGp, EDic, ADic)
'    OOrd = QpEprLis(FOrd, EDic, ADic)
'    OWh = Wh()
'    OT = RmvA1T(LFm)
'SqlSel = SqlSel_X_Into_T_Wh_Gp_Ord(OX, Into, OT, OGp, OOrd)
End Function

Function QpSelF(FF$, fldSw As Dictionary) As String()
Const CSub$ = CMod & "QpSelF"
Dim Fny$(): Fny = FnyFF(FF)
Dim F1$, F
For Each F In Fny
    F1 = ChrFst(F)
    Select Case True
    Case F1 = "?"
        If Not fldSw.Exists(F) Then Thw CSub, "An option fld not found in fldSw", "Opt-Fld Ff fldSw", F, FF, fldSw
        If fldSw(F) Then
            'PushI XFny, RmvFst(F)
        End If
    Case F1 = "$"
        'PushI XFny, RmvFst(F)
    Case Else
        'PushI XFny, F
    End Select
Next
Stop
End Function

Private Function IsSkip(FstSqln$, Sqlny$(), T As eSqSpect, stmtSw As Dictionary) As Boolean
Const CSub$ = CMod & "IsSkip"
If ChrFst(FstSqln) <> "?" Then Exit Function
Dim Key$: Key = WSwStmtKey(Sqlny, T)
If Not stmtSw.Exists(Key) Then Thw CSub, "stmtSw does not contain the WSwStmtKey", "Sqlny WSwStmtKey stmtSw", Sqlny, Key, stmtSw
IsSkip = Not stmtSw(Key)
End Function

Function SqyL(L() As SqDtaln, P As Pmsw2) As String()
Dim J%: For J = 0 To SqSrclnUB(L)
    PushI SqyL, SqlL(L(J), P)
Next
End Function

Function SqlSp$(Sqsp, P As Pmsw2)
Dim Sqlny$(): Sqlny = Sqsp
Dim FstSqln$:    FstSqln = Sqlny(0)
Dim Ty As eSqSpect:     Ty = SqSpect(FstSqln)
Dim Skip As Boolean:  Skip = IsSkip(FstSqln, Sqlny, Ty, P.Sw2.stmtSw)
                             If Skip Then Exit Function
Dim S$():                S = QpRmvEprLin(Sqlny)
Dim E As Dictionary: Set E = QpEprDic(Sqlny)
Dim O$
    Select Case True
    Case Ty = eDrpStmt: O = sqlDrp(S)
    Case Ty = eRUpdStmt: O = SqlUpd(S, E, P.Sw2.fldSw)
    Case Ty = eSelStmt: O = SqlSel(S, E, P.Sw2.fldSw)
    Case Else: ThwImposs CSub
    End Select
SqlSp = O
End Function

Private Function sqlDrp$(Drp$())
End Function
Private Function SqlUpd$(Upd$(), EDic As Dictionary, fldSw As Dictionary)
End Function



Function QpRmvEprLin(Sqlny$()) As String()
QpRmvEprLin = AePfx(Sqlny, "$")
End Function


Function QpEprDic(Sqlny$()) As Dictionary
Set QpEprDic = Dic(CvSy(AwPfx(Sqlny, "$")))
End Function

Function SqSpect(FstSqln$) As eSqSpect
Dim L$: L = RmvPfx(T1(FstSqln), "?")
Select Case L
Case "SEL", "SELDIS": SqSpect = eSelStmt
Case "UPD": SqSpect = eRUpdStmt
Case "DRP": SqSpect = eDrpStmt
Case Else: Stop
End Select
End Function

Private Function WSwStmtKey$(Sqlny$(), T As eSqSpect) ' #statment-switch-key#
Const CSub$ = CMod & "WSwStmtKey"
Dim O$
Select Case T
Case eSelStmt: O = WSwStmtKeySel(Sqlny)
Case eRUpdStmt: O = WSwStmtKeyUpd(Sqlny)
Case Else:  ThwEnm CSub, "eSqSpect", eSqSpectSS, T
End Select
WSwStmtKey = "?:" & O
End Function

Private Function WSwStmtKeySel$(SelSqlny$()): WSwStmtKeySel = EleFstwRmvT1(SelSqlny, "into"): End Function

Private Function WSwStmtKeyUpd$(UpdSqlny$())
Dim Lin1$
    Lin1 = UpdSqlny(0)
If RmvPfx(ShfTm(Lin1), "?") <> "upd" Then Stop
WSwStmtKeyUpd = Lin1
End Function

Function IsXXX(A$(), XXX$) As Boolean
IsXXX = UCase(T1(A(UB(A)))) = XXX
End Function

Function QpAnd(A$(), E As Dictionary)
'and f bet xx xx
'and f in xx
Dim F$, I, L$, Ix%
For Each I In Itr(A)
    'Set M = I
    'LnxAsg M, L, Ix
    If ShfTm(L) <> "and" Then Stop
    F = ShfTm(L)
    Select Case ShfTm(L)
    Case "bet":
    Case "in"
    Case Else: Stop
    End Select
Next
End Function

Function QpGp$(GG$, fldSw As Dictionary, E As Dictionary)
If GG = "" Then Exit Function
Dim EprAy$(), Ay$()
Stop
'    EprAy = DicSelIntoSy(EDic, Ay)
'XGp = SqpGp(EprAy)
End Function

Function QpJnOrLeftJn(A$(), E As Dictionary) As String()

End Function

Function QpSel$(A$, E As Dictionary)
Dim Fny$()
    Dim T1$, L$
    L = A
    T1 = RmvPfx(ShfTm(L), "?")
    'Fny = XSelFny(SySS(L), fldSw)
Select Case T1
'Case KwSel:    XSel = X.Sel_FnSampEDic(Fny, E)
'Case KwSelDis: XSel = X.Sel_FnSampEDic(Fny, E, IsDis:=True)
Case Else: Stop
End Select
End Function
Function QpSelFny(Fny$(), fldSw As Dictionary) As String()
Dim F
For Each F In Fny
    If ChrFst(F) = "?" Then
        If Not fldSw.Exists(F) Then Stop
        'If fldSw(F) Then PushI XSelFny, F
    Else
        'PushI XSelFny, F
    End If
Next
End Function

Private Function QpSet(DroLLn(), E As Dictionary, OEr$())

End Function

Private Function QpUpd(DroLLn(), E As Dictionary, OEr$())

End Function
Private Function Wh$() ' (L$, E As Dictionary)
'L is following
'  ?Fld in @ValLis  -
'  ?Fld bet @V1 @V2
Dim F$, Vy$(), V1, V2, IsBet As Boolean
If IsBet Then
'    If Not FndValPair(F, E, V1, V2) Then Exit Function
    'XWh = SWhBet(F, V1, V2)
    Exit Function
End If
'If Not FndVy(F, E, Vy, Q) Then Exit Function
'XWh = SWhFldInVSampStr(F, Vy)
End Function

Function WhBetNum$(DroLLn(), E As Dictionary, OEr$())

End Function

Function WhEpr(DroLLn(), E As Dictionary, OEr$())

End Function

Function WhInNumLis$(DroLLn(), E As Dictionary, OEr$())

End Function

Function CvVSampToTFFm01(A As Dictionary) As Dictionary
Dim O As Dictionary: Set O = DiClone(A)
Dim K
For Each K In O.Keys
    Select Case O(K)
    Case "0": O(K) = False
    Case "1": O(K) = True
    End Select
Next
Set CvVSampToTFFm01 = O
End Function

Private Sub B_SqlSel()
Dim E As Dictionary, Ly$(), fldSw As Dictionary

'---
Erase XX
    X "?XX Fld-XX"
    X "BB Fld-BB-LINE-1"
    X "BB Fld-BB-LINE-2"
    Set E = Dic(XX)           '<== Set EprDic
Erase XX
    X "?XX 0"
    Set fldSw = Dic(XX)
    Set fldSw = CvVSampToTFFm01(fldSw)
Erase XX
    X "sel ?XX BB CC"
    X "into #AA"
    X "fm   #AA"
    X "jn   #AA"
    X "jn   #AA"
    X "wh   A bet $a $b"
    X "and  B in $c"
    X "gp   D C"        '<== LySq
GoSub Tst
Exit Sub
Tst:
    Act = SqlSel(Ly, E, fldSw)
    C
    Return
End Sub

Private Sub B_EprDic()
Dim Ly$()
Dim D As New Dictionary
'-----

Erase Ly
PushI Ly, "aaa bbb"
PushI Ly, "111 222"
PushI Ly, "$"
PushI Ly, "A B0"
PushI Ly, "A B1"
PushI Ly, "A B2"
PushI Ly, "B B0"
D.RemoveAll
    D.Add "A", JnCrLf(SySs("B0 B1 B2"))
    D.Add "B", "B0"
    Set Ept = D
GoSub Tst
Exit Sub
Tst:
    Set Act = QpEprDic(Ly)
    Ass IsEqDic(CvDi(Act), CvDi(Ept))
    
    Return
End Sub

Private Sub B_WSwStmtKey()
Dim Ly$(), Ty As eSqSpect
GoSub T0
GoSub T1
Exit Sub
'---
T0:
    Erase Ly
    PushI Ly, "sel sdflk"
    PushI Ly, "fm AA BB"
    PushI Ly, "into XX"
    Ept = "XX"
    Ty = eSelStmt
    GoTo Tst
T1:
    Erase Ly
    PushI Ly, "?upd XX BB"
    PushI Ly, "fm dsklf dsfl"
    Ept = "XX BB"
    Ty = eRUpdStmt
    GoTo Tst
Tst:
    Act = WSwStmtKey(Ly, Ty)
    C
    Return
End Sub
#End If
