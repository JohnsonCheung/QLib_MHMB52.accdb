Attribute VB_Name = "MxDao_Dbt_PrpInf"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Dbt_PrpInf."
Const Skn$ = "SecondaryKey"

Sub AddFld(D As Database, T, F$, Ty As DataTypeEnum, Optional Si%, Optional Precious%)
If HasFld(D, T, F) Then Exit Sub
Dim S$, SqSpect$
SqSpect = SqlTyzDao(Ty, Si, Precious)
S = FmtQQ("Alter Table [?] Add Column [?] ?", T, F, Ty)
D.Execute S
End Sub
Function SqlTyzDao$(T As Dao.DatabaseTypeEnum, Si%, Precious%)

End Function
Sub AsgColApDrsFf(D As Drs, FF$, ParamArray OColAp())
Dim F, J%
For Each F In FnyFF(FF)
    OColAp(J) = DcDrs(D, CStr(F))
    J = J + 1
Next
End Sub

Sub BrwDtT(D As Database, T):                                    BrwDt DtT(D, T):        End Sub
Function AetF(D As Database, T, F$) As Dictionary:    Set AetF = AetRs(RsF(D, T, F)):    End Function
Function AetTF(D As Database, TF$) As Dictionary:    Set AetTF = AetRs(RsTF(D, TF)):     End Function
Function RsF(D As Database, T, F$) As Dao.Recordset:   Set RsF = Rs(D, SqlSelFld(T, F)): End Function
Function CsyTC(T) As String():                           CsyTC = CsyT(CDb, T):           End Function
Function CsyT(D As Database, T) As String():              CsyT = CsyRs(RsTbl(D, T)):     End Function

Function DaotyF(D As Database, T, F) As Dao.DataTypeEnum:   DaotyF = D.TableDefs(T).Fields(F).Type: End Function
Function DaotyFC(T, F) As Dao.DataTypeEnum:                DaotyFC = DaotyF(CDb, T, F):             End Function
Function DaotyTFC(TF$) As Dao.DataTypeEnum:               DaotyTFC = DaotyTF(CDb, TF):              End Function
Function BrkTFdot(TF$) As S12
Dim A As S12: A = BrkDot(TF)
BrkTFdot = S12(RmvBktSq(A.S1), RmvBktSq(A.S2))
End Function
Function DaotyTF(D As Database, TF$) As Dao.DataTypeEnum
Dim A As S12: A = BrkTFdot(TF)
DaotyTF = DaotyF(D, A.S1, A.S2)
End Function

Function DicntTF(D As Database, T, F$) As Dictionary: Set DicntTF = DicntRs(RsF(D, T, F$)): End Function

Sub DrpCol(D As Database, T$, F$):    D.Execute SqlDrpCol(T, F):    End Sub
Sub DrpColFf(D As Database, T$, FF$): D.Execute SqlDrpColFf(T, FF): End Sub
Sub DrpColC(T$, F$):                  DrpCol CDb, T, F:             End Sub
Sub DrpColFfC(T$, FF$):               DrpColFf CDb, T, FF:          End Sub

Function Fds(D As Database, T) As Dao.Fields: Set Fds = D.TableDefs(T).OpenRecordset.Fields: End Function
Function FnyC(T) As String():                    FnyC = Fny(CDb, T):                         End Function
Function Fny(D As Database, T) As String():       Fny = Itn(D.TableDefs(T).Fields):          End Function
Function FnyIf(D As Database, T) As String(): On Error Resume Next: FnyIf = Fny(D, T):                                End Function
Function FstUniqIdx(D As Database, T) As Dao.Index: Set FstUniqIdx = ItoFstPrpTrue(D.TableDefs(T).Indexes, "Unique"): End Function

Sub ChkFldExist(D As Database, T, F$, Optional Fun$ = "ChkFldExist")
If HasFld(D, T, F) Then Exit Sub
Thw Fun, "Fld should be found in Tbl", "Fld Tbl Db", F, T, D.Name
End Sub
Sub ChkFldNExi(D As Database, T, F$, Optional Fun$ = "ChkFldNExi")
If Not HasFld(D, T, F) Then Exit Sub
Thw Fun, "Fld should not be found in Tbl", "Fld Tbl Db", F, T, D.Name
End Sub
Function HasFld(D As Database, T, F$) As Boolean:    HasFld = HasItn(DbRfhTd(D).TableDefs(T).Fields, F): End Function
Function HasIdx(D As Database, T, Idxn$) As Boolean: HasIdx = HasItn(D.TableDefs(T).Indexes, Idxn):      End Function
Function HasId(D As Database, T, Id&) As Boolean
If HasPk(D, T) Then HasId = HasRec(RsId(D, T, Id))
End Function

Sub AddIdxF(D As Database, T, Idxn$, F): AddIdxFny D, T, Idxn, Sy(F): End Sub
Sub AddIdxFny(D As Database, T, Idxn$, Fny$())
Dim TT As Dao.TableDef, I As Dao.Index
Set TT = Td(D, T)
Set I = T
TT.Indexes.Append NwIdxTd(TT, Idxn, Fny)
TT.CreateIndex (Idxn)
End Sub

Function NwIdxTd(Td As Dao.TableDef, Idxn$, Fny$()) As Dao.Index: Td.CreateIndex Idxn: End Function

Function IsHidTbl(D As Database, T) As Boolean: IsHidTbl = (D.TableDefs(T).Attributes And Dao.TableDefAttributeEnum.dbHiddenObject) <> 0: End Function
Function IsSysTbl(D As Database, T) As Boolean: IsSysTbl = (D.TableDefs(T).Attributes And Dao.TableDefAttributeEnum.dbSystemObject) <> 0: End Function

Function IsTblLnk(D As Database, T) As Boolean:     IsTblLnk = IsTblLnkFb(D, T) Or IsTblLnkFx(D, T): End Function
Function IsTblLnkFb(D As Database, T) As Boolean: IsTblLnkFb = HasPfx(CnsT(D, T), ";Database="):     End Function
Function IsTblLnkFx(D As Database, T) As Boolean: IsTblLnkFx = HasPfx(CnsT(D, T), "Excel"):          End Function

Function Idx(D As Database, T, Idxn) As Dao.Index: Set Idx = IdxTd(Td(D, T), Idxn): End Function
Function IdxTd(Td As Dao.TableDef, Idxn) As Dao.Index:
Dim I As Dao.Index: For Each I In Td.Indexes
    If I.Name = Idxn Then Set IdxTd = I: Exit Function
Next
End Function
Function DcIntF(D As Database, T, F$) As Integer(): DcIntF = DcIntQ(D, FmtQQ("Select [?] from [?]", F, T)): End Function

Function IsMemCol(DcDrs) As Boolean
Dim I: For Each I In DcDrs
    If IsStr(I) Then
        If Len(I) > 255 Then IsMemCol = True: Exit Function
    End If
Next
End Function

Function IxF%(D As Database, T, F)
Dim O%, I As Dao.Field: For Each I In D.TableDefs(T).Fields
    If I.Name = F Then IxF = O: Exit Function
    O = O + 1
Next
IxF = -1
End Function

Function JnQSqCommaSpcAp$(ParamArray Ap())
Dim Av(): Av = Ap
JnQSqCommaSpcAp = JnQSqCommaSpc(SyAy(Av))
End Function

Sub KillIfDbTmp(D As Database)
If IsDbTmp(D) Then
    Dim Fb$: Fb = D.Name
    ClsDb D
    Kill Fb
End If
End Sub

Function TimlnyLasUpdC() As String(): TimlnyLasUpdC = TimlnyLasUpd(CDb): End Function

Function TimlnyLasUpd(D As Database) As String()
Dim T: For Each T In Tni(D)
    PushI TimlnyLasUpd, T & " = " & TimLasUpd(D, T)
Next
End Function
Function TimLasUpd(D As Database, T) As Date: TimLasUpd = PvT(D, T, "LastUpdated"): End Function

Function TimlnyCrtC() As String(): TimlnyCrtC = TimlnyCrt(CDb): End Function
Function TimlnyCrt(D As Database) As String()
Dim T: For Each T In Tni(D)
    PushI TimlnyCrt, T & " = " & TimCrt(D, T)
Next
End Function
Function TimCrt(D As Database, T) As Date:  TimCrt = PvT(D, T, "DateCreated"): End Function
Function LnklnyC() As String():            LnklnyC = Lnklny(CDb):              End Function
Function Lnklny(D As Database) As String()
Dim T: For Each T In Tni(D)
    PushNB Lnklny, LnklnT(D, T)
Next
End Function

Function LnklnT$(D As Database, T)
If IsTblLnk(D, T) Then LnklnT = T & FmtQQ(";SrcTbn=?;Cns=?", SrcTbnT(D, T), CnsT(D, T))
End Function
Function SrcTbnT$(D As Database, T): SrcTbnT = D.TableDefs(T).SourceTableName: End Function

Function LoflDbt$(D As Database, T): LoflDbt = PvT(D, T, "Lofl"): End Function

Function VbtMax(A As VbVarType, B As VbVarType) As VbVarType
Const CSub$ = CMod & "VbtMax"
If A = B Then VbtMax = A: Exit Function
If Not IsVbtNum(B) Then Thw CSub, "Given B is not NumVbt", "B-VarType", B
Dim O As VbVarType
Select Case A
Case VbVarType.vbByte:      O = B
Case VbVarType.vbInteger:   O = IIf(B = vbByte, A, B)
Case VbVarType.vbLong:      O = IIf((B = vbByte) Or (B = vbInteger), A, B)
Case VbVarType.vbSingle:    O = IIf((B = vbByte) Or (B = vbInteger) Or (B = vbLong), A, B)
Case VbVarType.vbDecimal:   O = IIf((B = vbByte) Or (B = vbInteger) Or (B = vbLong) Or (B = vbSingle), A, B)
Case VbVarType.vbDouble:    O = IIf((B = vbByte) Or (B = vbInteger) Or (B = vbLong) Or (B = vbSingle) Or (B = vbDecimal), A, B)
Case VbVarType.vbCurrency:  O = IIf((B = vbByte) Or (B = vbInteger) Or (B = vbLong) Or (B = vbSingle) Or (B = vbDecimal) Or (B = vbDouble), A, B)
Case Else:                  Thw CSub, "Given A is not NumVbt", "A-VarType", A
End Select
VbtMax = O
End Function

Function NDcT&(D As Database, T):                      NDcT = D.TableDefs(T).Fields.Count:                          End Function
Function NRecFxw&(Fx, Wsn, Optional Bepr$):         NRecFxw = ValCnq(CnFx(Fx), SqlSelCnt(Axtn(Wsn), Bepr)):         End Function
Function NRecT&(D As Database, T, Optional Bepr$):    NRecT = ValQ(D, SqlSelCnt(T, Bepr)):                          End Function
Function NRecTFeq&(D As Database, T, F, Eqval):    NRecTFeq = NRecT(D, T, BeprFeq(F, Eqval)):                       End Function
Function NRecTC&(T, Optional Bepr$):                 NRecTC = NRecT(CDb, T, Bepr):                                  End Function
Function NxtId&(D As Database, T):                    NxtId = ValQ(D, FmtQQ("select Max(?Id) from [?]", T, T)) + 1: End Function

Function FnyPk(D As Database, T) As String():           FnyPk = FnyIdx(PkIdx(D, T)):      End Function
Function FnyPkTd(A As Dao.TableDef) As String():      FnyPkTd = FnyIdx(PkIdxTd(A)):       End Function
Function PkIdx(D As Database, T) As Dao.Index:      Set PkIdx = PkIdxTd(D.TableDefs(T)):  End Function
Function PkIdxn$(D As Database, T):                    PkIdxn = Objn(PkIdx(D, T)):        End Function
Function PkIdxTd(A As Dao.TableDef) As Dao.Index: Set PkIdxTd = ItoFstNm(A.Indexes, Pkn): End Function

Sub RenFlds(D As Database, T, FmFf$, ToFf$)
Dim FmFny$(): FmFny = FnyFF(FmFf)
Dim ToFny$(): ToFny = FnyFF(ToFf)
Dim J%: For J = UBound(FmFny) To 0 Step -1
    RenFld D, T, FmFny(J), ToFny(J)
Next
End Sub
Sub RenFldFf(T, FmFf$, ToFf$):            RenFlds CDb, T, FmFf$, ToFf$:          End Sub
Sub RenFld(D As Database, T, F$, ToFld$): D.TableDefs(T).Fields(F).Name = ToFld: End Sub
Sub RenTblStrPfx(D As Database, T, Pfx$): RenT D, T, Pfx & T:                    End Sub

Function RsId(D As Database, T, Id&) As Dao.Recordset:                Set RsId = Rs(D, SqlSelStarWhId(T, Id)):           End Function
Function RsTbl(D As Database, T, Optional Bepr$) As Dao.Recordset:   Set RsTbl = Rs(D, SqlSelStar(T, Bepr)):             End Function
Function RsTblC(T, Optional Bepr$) As Dao.Recordset:                Set RsTblC = RsTbl(CDb, T, Bepr):                    End Function
Function RsTF(D As Database, TF$, Optional Bepr$) As Dao.Recordset:   Set RsTF = D.OpenRecordset(SqlSelTFdot(TF, Bepr)): End Function
Function FfRst(D As Database, T, FF$) As Dao.Recordset:              Set FfRst = RsTFny(D, T, FnyFF(FF)):                End Function
Function RsTFny(D As Database, T, Fny$()) As Dao.Recordset:         Set RsTFny = D.OpenRecordset(SqlSelFny(T, Fny)):     End Function
Function DiRs(A As Dao.Recordset) As Dictionary
Set DiRs = New Dictionary
Dim F As Dao.Field
For Each F In A.Fields
    DiRs.Add F.Name, F.Value
Next
End Function


Function FbTdCn$(D As Database, T): FbTdCn = BetS1Opt2(D.TableDefs(T).Connect, "Database=", ";"): End Function
Function SrcTbn$(D As Database, T): SrcTbn = D.TableDefs(T).SourceTableName:                      End Function


Private Sub B_CrttDup()
GoSub Z
Exit Sub
Dim D As Database
Z:
    Set D = DbTmp
    DrpTmp D
    CrttDrs D, "#D", sampDrs
    CrttDup D, "#DDup", "#D", ""
    BrwDb D
    Return
End Sub

Private Sub B_FnyPk()
Z:
    Dim D As Database
    Stop 'Set D = Db(FbDtaMHDuty)
    Dim Dr(), Dy(), T, I
    For Each I In Tny(D)
        T = I
        Erase Dr
        Push Dr, T
        PushIAy Dr, FnyPk(D, T)
        PushI Dy, Dr
    Next
    BrwDy Dy
    Exit Sub
End Sub
