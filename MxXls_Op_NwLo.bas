Attribute VB_Name = "MxXls_Op_NwLo"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Op_NwLo."
Enum eLoAddTblWay: eLoAddTblWayRs: eLoAddTblWayWc: eLoAddTblWaySq: End Enum ' Adding data to ws way

Private Sub B_LoT()
Dim R As Range: Set R = A1Nw
Dim At As Range: Set At = A1Nw
Dim Fb$: Fb = FbTmp: CpyFfn CFb, Fb
LoT At, CDb, "OH", eLoAddTblWayWc
Maxv At.Application
End Sub
Function LoT(At As Range, D As Database, T, Optional Way As eLoAddTblWay) As ListObject
Dim Fb$: Fb = D.Name
Select Case True
Case Way = eLoAddTblWayRs, Way = eLoAddTblWaySq
    Set LoT = LoRg(RgFmT(At, Db(Fb), T, Way))
Case Way = eLoAddTblWayWc
    Dim Cns$: Cns = CnsFbOle(Fb)
    Set LoT = WsRg(At).ListObjects.Add(xlSrcExternal, Cns, , xlYes, At)
    With LoT.QueryTable
        .CommandType = xlCmdTable
        .CommandText = T
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = False
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        '.SourceDataFile = D.Name
        .ListObject.DisplayName = LonTbn(T)
        .Refresh BackgroundQuery:=True
    End With
End Select
'    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=Array( _
'        "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=C:\Users\user\Desktop\SAPAccessReports\DutyPrepay5\DutyP" _
'        , _
'        "repay5_Data.mdb;Mode=Share Deny Write;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:D" _
'        , _
'        "atabase Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Glob" _
'        , _
'        "al Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=Fals" _
'        , _
'        "e;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Supp" _
'        , _
'        "ort Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceFiel" _
'        , "d Validation=False"), Destination:=Range("$H$4")).QueryTable
'        .CommandType = xlCmdTable
'        .CommandText = Array("@RptM")
'        .RowNumbers = False
'        .FillAdjacentFormulas = False
'        .PreserveFormatting = True
'        .RefreshOnFileOpen = False
'        .BackgroundQuery = True
'        .RefreshStyle = xlInsertDeleteCells
'        .SavePassword = False
'        .SaveData = True
'        .AdjustColumnWidth = True
'        .RefreshPeriod = 0
'        .PreserveColumnInfo = True
'        .SourceDataFile = _
'        "C:\Users\user\Desktop\SAPAccessReports\DutyPrepay5\DutyPrepay5_Data.mdb"
'        .ListObject.DisplayName = "Table_DutyPrepay5_Data_1"
'        .Refresh BackgroundQuery:=False
'    End With
End Function

Private Sub B_RgFmT()
Dim At As Range: Set At = A1Nw
RgFmT At, CDb, "OH", eLoAddTblWaySq
Maxv At.Application
End Sub
Function RgFmT(At As Range, D As Database, T, Optional Way As eLoAddTblWay) As Range
Select Case True
Case Way = eLoAddTblWayRs
    Dim Rs As Dao.Recordset: Set Rs = RsTbl(D, T)
    RgAyH FnyRs(Rs), At
    RgRC(At, 2, 1).CopyFromRecordset Rs
    Set RgFmT = RgRCRC(At, 1, 1, Rs.RecordCount + 1, Rs.Fields.Count)
Case Way = eLoAddTblWaySq
    Dim Sq(): Sq = SqT(D, T, True)
    Set RgFmT = RgRCRC(At, 1, 1, NDrSq(Sq), NDcSq(Sq))
    RgFmT.Value = Sq
End Select
End Function
