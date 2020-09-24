Attribute VB_Name = "MxDao_Dbt_DbtPrp"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Dbt_Prp."
Function PthDbC$():             PthDbC = PthDb(CDb):  End Function
Function PthDb$(D As Database):  PthDb = Pth(D.Name): End Function

Function HasTC(T) As Boolean: HasTC = HasT(CDb, T): End Function
Function HasT(D As Database, T) As Boolean
D.TableDefs.Refresh
HasT = HasItn(D.TableDefs, T):
End Function
Function HasTByMSys(D As Database, T) As Boolean
Dim W$: W = FmtQQ("Type in (1,6) and Name='?'", T)
HasTByMSys = HasRecQ(D, SqlSelStar("MSysObjects", W))
End Function

Function FFzT$(D As Database, T):  FFzT = TmlAy(Fny(D, T)): End Function
Function FFzTC$(T):               FFzTC = FFzT(CDb, T):     End Function

Function HasFF(D As Database, T, FF$) As Boolean
HasFF = HasSubAy(Fny(D, T), FnyFF(FF))
End Function

Function HasQnC(Qn) As Boolean: HasQnC = HasQn(CDb, Qn): End Function
Function HasQn(D As Database, Qn) As Boolean
Dim W$: W = FmtQQ("Name='?' and Type=5", Qn)
HasQn = HasRecQ(D, SqlSelStar("MSysObjects", W))
End Function

Function HasRecQn(D As Database, Qn) As Boolean:  HasRecQn = HasRec(RsQn(D, Qn)): End Function
Function HasRecQ(D As Database, Q) As Boolean:     HasRecQ = HasRec(Rs(D, Q)):    End Function
Function HasRecQC(Q) As Boolean:                  HasRecQC = HasRecQ(CDb, Q):     End Function
Function HasRecQnC(Qn) As Boolean:               HasRecQnC = HasRecQn(CDb, Qn):   End Function
Function HasRecTC(T) As Boolean:                  HasRecTC = HasRecT(CDb, T):     End Function
Function HasRecT(D As Database, T) As Boolean:     HasRecT = HasRec(RsTbl(D, T)): End Function


Function RsQQ(D As Database, QQ$, ParamArray Ap()) As Dao.Recordset
Dim Av():  Av = Ap
Set RsQQ = Rs(D, FmtQQAv(QQ, Av))
End Function

Function Tdrepy(D As Database, TT$) As String()
Dim T: For Each T In ItrTml(TT)
    PushI Tdrepy, Tdrep(D, T)
Next
End Function

Function FbyTmp() As String(): FbyTmp = Ffny(PthTmpDb, "*.accdb"): End Function
Sub ChkFbtNExi(Fb$, T, Optional Fun$ = CMod & "ChkTExist")
If HasFbt(Fb, T) Then Thw Fun, "In @Fb, should not have @T", "@Fb @T", Fb, T
End Sub
Sub ChkFbtExi(Fb$, T, Optional Fun$ = CMod & "ChkTExist")
If Not HasFbt(Fb, T) Then Thw Fun, "In @Fb, should not have @T", "@Fb @T", Fb, T
End Sub

Function CnsT$(D As Database, T):       CnsT = D.TableDefs(T).Connect:  End Function
Function CnsTC$(T):                    CnsTC = CnsT(CDb, T):            End Function
Function DbnCnnT$(D As Database, T): DbnCnnT = ScvDatabase(CnsT(D, T)): End Function

Private Sub B_BrwT()
Dim D As Database
DrpTT D, "#A #B"
Runq D, "Select Distinct Sku,BchNo,CLng(Rate) as RateRnd into [#A] from [#T]"
BrwTT D, "#A #T #B"
End Sub

Private Sub B_DsDb()
Dim D As Database, Tny0, Act As Ds, Ept As Ds
Stop
YY1:
    Stop 'Set D = Db(MHODuty.DbDta)
    Act = DsDb(D)
    BrwDs Act
    Exit Sub
YY2:
    Tny0 = "Permit PermitD"
    'Set Act = Ds( Tny0)
    Stop
End Sub

Function CDb() As Database
Static X As Database
If Not IsDbOk(X) Then Set X = CurrentDb
Set CDb = X
End Function
Function CDbn$():                      CDbn = CDb.Name:                  End Function
Function CFb$():                        CFb = CurrentDb.Name:            End Function
Function CCn() As ADODB.Connection: Set CCn = CurrentProject.Connection: End Function
Sub DmpFldSnoC(T):                            DmpFldSno CDb, T:          End Sub
Sub DmpFldSno(D As Database, T)
Dim J%
Dim F As Dao.Field: For Each F In Td(D, T).Fields
    Debug.Print J, F.Name, F.OrdinalPosition, StrTrue(J <> F.OrdinalPosition, "  <===== Bad")
    J = J + 1
Next
End Sub

Function DiFldDes(D As Database) As Dictionary
Dim Des$
Set DiFldDes = New Dictionary
Dim T: For Each T In Tni(D)
    Dim F: For Each F In Fny(D, T)
        Des = PvFDes(D, T, F)
        If Des <> "" Then DiFldDes.Add T & "." & F, D
    Next
Next
End Function
Function SqT(D As Database, T, Optional ExlFldn As Boolean) As Variant():  SqT = SqRs(RsTbl(D, T), ExlFldn): End Function
Function SqTC(T, Optional IsInlFldn As Boolean = True) As Variant():      SqTC = SqT(CDb, T, IsInlFldn):     End Function
