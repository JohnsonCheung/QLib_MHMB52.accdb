Attribute VB_Name = "MxDao_Db_Inf"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Db_Inf."
Private Sub B_BrwDbInf()
'strDdl = "GRANT SELECT ON MSysObjects TO Admin;"
'CurrentProject.Connection.Execute strDdlDim A As DBEngine: Set A = dao.DBEngine
'not work: dao.DBEngine.Workspaces(1).Databases(1).Execute "GRANT SELECT ON MSysObjects TO Admin;"
BrwDbInf MHO.MHODuty.DbDta
End Sub
Sub BrwDbInfC():             BrwDbInf CDb:   End Sub
Sub BrwDbInf(D As Database): BrwDs DbInf(D): End Sub

Function DbInfC() As Ds: DbInfC = DbInf(CDb): End Function
Function DbInf(D As Database) As Ds
Dim O() As Dt, T$()
T = Tny(D)
PushDt O, WTbl(D, T)
PushDt O, WTblDes(D, T)
PushDt O, WLnk(D, T)
PushDt O, WTblF(D, T)
PushDt O, WPrp(D)
PushDt O, WFld(D, T)
With DbInf
    .Dsn = D.Name
    .Dty = O
End With
End Function
Private Function WTblfDr(T, Seq%, F As Dao.Field2) As Variant(): WTblfDr = Array(T, Seq, F.Name, DtaTy(F.Type)): End Function
Private Function WFld(D As Database, Tny$()) As Dt
Dim Dy(), T
For Each T In Tni(D)
Next
WFld = DtFf("DbFld", "Tbl Fld Pk Ty Si Dft Req Des", Dy)
End Function
Private Function WLnk(D As Database, Tny$()) As Dt
Dim Dy(), C$
Dim T: For Each T In Tni(D)
   C = D.TableDefs(T).Connect
   If C <> "" Then Push Dy, Array(T, C)
Next
Dim O As Dt
WLnk = DtFf("DbLnk", "Tbl Connect", Dy)
End Function
Private Function WLnkLy(D As Database) As String()
Dim T$, I
For Each I In Tny(D)
    T = I
    PushNB WLnkLy, CnsT(D, T)
Next
End Function
Private Function WPrp(D As Database) As Dt
Dim Dy()
WPrp = DtFf("DbPrp", "Prp Ty Val", Dy)
End Function
Private Function WTblDes(D As Database, Tny$()) As Dt
Dim Dy(), Des$
Dim T: For Each T In Tny
    Des = PvTDes(D, T)
    If Des <> "" Then
        Push Dy, Array(T, Des)
    End If
Next
WTblDes = DtFf("PvTDes", "Tbl Des", Dy)
End Function
Private Sub B_WTbl(): DmpDt WTbl(CDb, TnyC): End Sub
Private Function WTbl(D As Database, Tny$()) As Dt
Dim T, Dy()
For Each T In Tny
    Push Dy, Array(T, NRecT(D, T), StruT(D, T))
Next
WTbl = DtFf("Tbl", "Tbl RecCnt Stru", Dy)
End Function
Private Function WTblF(D As Database, Tny$()) As Dt
Dim Dy()
Dim T$, I
For Each I In Tni(D)
    T = I
    PushIAy Dy, WTblfDy(D, T)
Next
WTblF = DtFf("TblFld", "Tbl Seq Fld Ty Si ", Dy)
End Function
Private Function WTblfDy(D As Database, T) As Variant()
Dim F$, Seq%, I
For Each I In Fny(D, T)
    F = I
    Seq = Seq + 1
    Push WTblfDy, WTblfDr(T, Seq, Fd(D, T, F))
Next
End Function
