Attribute VB_Name = "MxDao_Sql"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Sql."

Function SqlColAddFnyDi$(T, Fny$(), DiFqSqlTy As Dictionary)
Const CSub$ = CMod & "SqlColAddFnyDi"
Dim O$()
Dim F: For Each F In Fny
    PushI O, F & " " & ValDiThw(DiFqSqlTy, F, "a field in @Fny not found the @DiFqSqlTy", CSub)
Next
SqlColAddFnyDi = FmtQQ("Alter Table [?] add column ?", T, JnCma(O))
End Function

Function SqlColAddAy$(T, ColAy$()):  SqlColAddAy = SqlColAdd(T, JnCmaSpc(ColAy)):                                   End Function
Function SqlColDrpFf(T, FF$):        SqlColDrpFf = "Alter Table [" & T & "] Drop Column " & QpFf(FF):               End Function
Function SqlColAdd$(T, LisCol$):       SqlColAdd = FmtQQ("Alter Table [?] add column ?", T, LisCol):                End Function
Function SqlCrtTbl$(T, X$):            SqlCrtTbl = FmtQQ("Create Table [?] (?)", T, X):                             End Function
Function SqlDlt$(T, Optional Bepr$):      SqlDlt = "Delete * from [" & T & "]" & Wh(Bepr):                          End Function
Function SqlDrpColFf$(T, FF$):       SqlDrpColFf = FmtQQ("Alter Table [?] drop column ?", T, QpFf(FF)):             End Function
Function SqlDrpCol$(T, X$):            SqlDrpCol = FmtQQ("Alter Table [?] drop column ?", T, X$):                   End Function
Function SqlDrpFny$(T, Fny$()):        SqlDrpFny = "Alter Table [" & T & "] drop column " & JnCmaSpc(AmQuoSq(Fny)): End Function
Function SqlDrpTbl$(T):                SqlDrpTbl = "Drop Table [" & T & "]":                                        End Function

Function SqlInsFfDr$(T, FF$, Dr): SqlInsFfDr = FmtQQ("Insert Into [?] (?) Values(?)", T, QpFf(FF), QpValues(Dr)): End Function
Function SqlInsFfVap$(T, FF$, ParamArray Vap())
Dim Vy(): Vy = Vap
SqlInsFfVap = QpInsT(T) & QpBktFf(FF) & " Values" & QpBktVy(Vy)
End Function

Function SqlSelFldWhFldIn(T, F, WhF, VyIn):                           SqlSelFldWhFldIn = SqlSelFld(T, F, BeprFldIn(WhF, VyIn)):   End Function
Function SqlSelFldWhFeq(T, F, WhF, Eqval):                              SqlSelFldWhFeq = SqlSelFld(T, F, BeprFeq(WhF, Eqval)):    End Function
Function SqlSelFldWhFnyEq(T, F, WhFny$(), Eqvy):                      SqlSelFldWhFnyEq = SqlSelFld(T, F, BeprFnyEq(WhFny, Eqvy)): End Function
Function SqlSelFld$(T, F, Optional Bepr$, Optional IsDis As Boolean):        SqlSelFld = QpSelF(F, IsDis) & QpFm(T) & Wh(Bepr):   End Function
Function SqlSelTFdot$(TF$, Optional Bepr$, Optional IsDis As Boolean)
With BrkTFdot(TF)
SqlSelTFdot = SqlSelFld(.S1, .S2, Bepr, IsDis)
End With
End Function

Function SqlSelX$(T, X$, Optional Bepr$):                                    SqlSelX = QpSelX(X) & QpFm(T) & Wh(Bepr): End Function
Function SqlSelFfAs$(T, FF$, E As Dictionary, Optional IsDis As Boolean): SqlSelFfAs = QpSelX(QpFfAs(FF, E), IsDis):   End Function
Function SqlSelFf(T, FF$, Optional IsDis As Boolean, Optional Bepr$, Optional FfHypSfxOrd$)
SqlSelFf = QpSelFf(FF, IsDis) & QpFm(T) & Wh(Bepr) & QpOrdFfMinusSfx(FfHypSfxOrd)
End Function
Function SqlSelFfWhFldIn$(T, FF$, WhF$, VyIn, Optional IsDis As Boolean, Optional FfHypSfxOrd$)
SqlSelFfWhFldIn = SqlSelFf(T, FF$, IsDis, BeprFldIn(WhF$, VyIn), FfHypSfxOrd)
End Function
Function SqlSelFny(T, Fny$(), Optional Bepr$, Optional IsDis As Boolean):        SqlSelFny = QpSelFny(Fny, IsDis) & QpFm(T) & Wh(Bepr): End Function
Function SqlSelFnyWhFnyEq$(Fny$(), T, WhFny$(), Eqvy):                    SqlSelFnyWhFnyEq = SqlSelFny(T, Fny, WhFnyEq(WhFny, Eqvy)):   End Function

Function SqlIntoSelExt$(Into, T, Fny$(), Extny$(), Optional Bepr$):         SqlIntoSelExt = QpSelFnyExtny(Fny, Extny) & QpInto(Into) & QpFm(T) & Wh(Bepr): End Function
Function SqlIntoSelStarWhFalse$(Into, T):                           SqlIntoSelStarWhFalse = FmtQQ("Select * Into [?] from [?] where false", Into, T):      End Function
Function SqlIntoSelStar$(Into$, T$, Optional Bepr$):                       SqlIntoSelStar = QpSelStar & QpInto(Into) & QpFm(T) & Wh(Bepr):                 End Function
Function SqlIntoSelFfEDi$(Into$, T, FF$, EDic As Dictionary, Optional Bepr$)
Dim Fny$(): Fny = FnyFF(FF)
Dim EprAy$(): EprAy = SyDicKy(EDic, Fny)
SqlIntoSelFfEDi = SqlIntoSelExt(Into, T, Fny, EprAy, Bepr)
End Function

Function SqlSelCnt$(T, Optional Bepr$):  SqlSelCnt = "select Count(*) from [" & T & "]" & Wh(Bepr):                                End Function
Function SqlSelTFeq(TF$, TFeq):         SqlSelTFeq = FmtQQ("Select * from [?] where [?]=?", TzTF(TF), FzTF(TF), QuoSqlPrim(TFeq)): End Function
Function SqlIntoSelX$(Into$, T, X$, Optional Bepr$, Optional GpBy$, Optional FfHypSfxOrd$, Optional IsDis As Boolean)
SqlIntoSelX = QpSelX(X) & QpInto(Into) & QpFm(T) & Wh(Bepr) & QpGp(GpBy) & QpOrd(FfHypSfxOrd$)
End Function
Function SqlSelStarWhId$(T, Id&):                                SqlSelStarWhId = QpIntoSelStar(T) & WhTblId(T, Id):                   End Function
Function SqlSelStarFeq$(T, F$, Feqv):                             SqlSelStarFeq = QpSelStar & QpFm(T) & WhFeq(F, Feqv):                End Function
Function SqlSelStarFnyEq$(T, Fny$(), Eqvy):                     SqlSelStarFnyEq = QpSelStar & QpFm(T) & WhFnyEq(Fny, Eqvy):            End Function
Function SqlSelStarFfeq$(T, FF$, Eqvy):                          SqlSelStarFfeq = SqlSelStarFnyEq(T, FnyFF(FF), Eqvy):                 End Function
Function SqlSelStarSkvy$(D As Database, T, Skvy()):              SqlSelStarSkvy = SqlSelStarFnyEq(T, FnySk(D, T), Skvy):               End Function
Function SqlSelStar$(T, Optional Bepr$, Optional FfHypSfxOrd$):      SqlSelStar = QpSelStar & QpFm(T) & Wh(Bepr) & QpOrd(FfHypSfxOrd): End Function

Function SqlIntoSelFf$(Into$, T, FF$, Optional Bepr$, Optional IsDis As Boolean): SqlIntoSelFf = QpSelFf(FF, IsDis) & QpInto(Into) & QpFm(T) & Wh(Bepr): End Function
Function SqlUpdFfWhFfeq$(T, FF$, Dr(), WhFf$, Eqvy)
SqlUpdFfWhFfeq = QpUpd(T) & QpSetValFf(FF, Dr) & WhFfeq(WhFf, Eqvy)
End Function

Function SqlUpdFnyDrSk$(T, Fny$(), Dr, Sk$())
If Si(Sk) = 0 Then Stop
Dim MUpd$, MSet$, MWh$: GoSub XMain
SqlUpdFnyDrSk = MUpd & MSet & MWh
Exit Function
XMain:
    Dim Fny1$(), Dr1(), Skvy(): GoSub X_Fny1_Dr1_SkVy
    MUpd = "Update [" & T & "]"
    MSet = QpSetVal(Fny1, Dr1)
    MWh = WhFnyEq(Sk, Skvy)
    Return
X_Ay:
    Dim L$(), R$()
    L = AmAliQusq(Fny)
    R = QuoSqlPrimy(Dr)
    Return
X_Fny1_Dr1_SkVy:
    Dim Ski, J%, Ixy%(), I%
    For Each Ski In Sk
'        I = IxEle(Fny, Ski)
        If I = -1 Then Stop
        Push Ixy, I
        Push Skvy, Dr(I)    '<====
    Next
    Dim F
    For Each F In Fny
        If Not HasEle(Ixy, J) Then
            Push Fny1, F        '<===
            Push Dr1, Dr(J)     '<===
        End If
        J = J + 1
    Next
    Return
End Function

Function SqlUpdFnyEpry$(T, Fny$(), Ey$(), Optional Bepr$): SqlUpdFnyEpry = QpUpd(T) & QpSetEpr(Fny, Ey) & Wh(Bepr): End Function

Function SqlUpdXA(X$, A$, TmlJn$, TmlSet$)
SqlUpdXA = QpUpdXA(X, A, TmlJn) & QpSetXATml(TmlSet)
End Function

Function SqlUpdSet$(TbX$, SetX$, Optional Bepr$): SqlUpdSet = QpUpd(TbX) & QpSetX(SetX) & Wh(Bepr): End Function

Private Sub B_SqlIntoSelExt()
Dim Fny$(), Ey$(), Into$, T$, Bepr$
GoSub Z
Exit Sub
Z:
    Fny = SySs("Sku CurRateAc VdtFm VdtTo HKD Per CA_Uom")
    Ey = Tmy("Sku [     Amount] [Valid From] [Valid to] Unit per Uom")
    Into = "#IZHT086"
    T = ">ZHT086"
    Bepr = ""
    Debug.Print SqlIntoSelExt(Into, T, Fny, Ey, Bepr)
    Return
End Sub
Function SqlUpd$(T, SetEq$): SqlUpd = "Update [" & T & "] set " & vbCrLf & SetEq: End Function
