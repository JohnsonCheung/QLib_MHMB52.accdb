Attribute VB_Name = "MxXls_Wb_AddWs_FmTbl"
Option Compare Text
Const CMod$ = "MxXls_Wb_AddWs_FmTbl."
Option Explicit

Function WsTbllC(T, Optional Way As eLoAddTblWay) As Worksheet:     Set WsTbllC = WsTbl(CDb, T, Way):       End Function
Function WsTbl(D As Database, T, Way As eLoAddTblWay) As Worksheet:   Set WsTbl = WsTblAt(A1Nw, D, T, Way): End Function
Function WsTblAt(At As Range, D As Database, T, Optional Way As eLoAddTblWay) As Worksheet
Dim L As ListObject: Set L = LoT(At, D, T, Way)
Set WsTblAt = WsLo(L)
SetWsn WsTblAt, CStr(T)
SetLon L, LonTbn(T)
End Function

Function WsTblWb(B As Workbook, D As Database, T, Optional Way As eLoAddTblWay) As Worksheet: Set WsTblWb = WsTblAt(A1Ws(WsAdd(B)), D, T, Way): End Function

Function WsAddWc(C As WorkbookConnection, B As Workbook) As Worksheet
Dim S As Worksheet: Set S = WsAdd(B, C.Name)
Set WsAddWc = WsLo(WLoWc(S, C))
End Function
Private Function WLoWc(S As Worksheet, C As WorkbookConnection) As ListObject
Dim Ty As XlListObjectSourceType
Dim L As ListObject: 'Set L = S.ListObjects.Add(Ty, Src, Dest)
'WSetQt L.QueryTable.CommandText = C.Delete
Set WLoWc = L
'    Application.CutCopyMode = False
'    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=Array( _
'        "OLEDB;Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=C:\Users\user\Documents\Projects\Vba\QLib\QLib_StockHolding8_Backup." _
'        , _
'        "accdb;Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=" _
'        , _
'        "6;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Datab" _
'        , _
'        "ase Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=F" _
'        , _
'        "alse;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass Us" _
'        , _
'        "erInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False" _
'        ), Destination:=Range("$A$1")).QueryTable
'        .CommandType = xlCmdTable
'        .CommandText = Array("Fc")
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
'        .SourceConnectionFile = _
'        "C:\Users\user\Documents\My Data Sources\(Default) Fc.odc"
'        .ListObject.DisplayName = "Table_Default__Fc"
'        .Refresh BackgroundQuery:=False
'    End With
End Function
