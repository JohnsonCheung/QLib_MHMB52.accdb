Attribute VB_Name = "MxXls_Wc"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Wc."
Function WsWc(Wc As WorkbookConnection) As Worksheet
Dim Wb As Workbook, Ws As Worksheet
Set Wb = Wc.Parent
Set Ws = WsAdd(Wb, Wc.Name)
PutWc Wc, A1Ws(Ws)
Set WsWc = Ws
End Function

Sub RplWc(Wc As WorkbookConnection, Fb)
CvCn(Wc.OLEDBConnection.ADOConnection).ConnectionString = CnsFbAdo(Fb)
End Sub



Function WcnyWb(B As Workbook) As String(): WcnyWb = Itn(B.Connections): End Function

Private Sub B_WbWcStrAyWbOLE()
'D WcsyOleWb(WbFx(TpFx))
End Sub
Function WcsyOleWb(B As Workbook) As String()
Dim O() As OLEDBConnection
Dim Wc As WorkbookConnection: For Each Wc In B.Connections
    If Not IsNothing(Wc.OLEDBConnection) Then
        PushI WcsyOleWb, Wc.OLEDBConnection.Connection
    End If
Next
End Function

Sub DltWc(B As Workbook)
Dim Wc As Excel.WorkbookConnection
For Each Wc In B.Connections
    Wc.Delete
Next
End Sub
Sub ClsWc(C As WorkbookConnection)
If IsNothing(C.OLEDBConnection) Then Exit Sub
CvCn(C.ODBCConnection.Connection).Close
End Sub
Sub ClsWcWb(Wb As Workbook)
'After Qt refresh the Fb will be locked and cannot save after edit the source.
'To close the lock, it is not "To close all the workbookconnects", but "set the Qt.MaintainConnection= False" See !W?RfhQt
Dim C As WorkbookConnection: For Each C In Wb.Connections
    ClsWc C
Next
End Sub

Sub SetWcFb(C As WorkbookConnection, Fb)
If IsNothing(C.OLEDBConnection) Then Exit Sub
Dim Cn$
#Const A = 2
#If A = 1 Then
    Dim S$
    S = A.OLEDBConnection.Connection
    Cn = RplBet(S, Fb, "Data Source=", ";")
#ElseIf A = 2 Then
    Cn = CnsFbOle(Fb)
#End If
C.OLEDBConnection.Connection = Cn
End Sub

Sub PutWc(Wc As WorkbookConnection, At As Range)
Dim Lo As ListObject
Set Lo = WsRg(At).ListObjects.Add(SourceType:=0, Source:=Wc.OLEDBConnection.Connection, Destination:=At)
With Lo.QueryTable
    .CommandType = xlCmdTable
    .CommandText = Wc.Name
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .BackgroundQuery = True
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .PreserveColumnINF = True
    .ListObject.DisplayName = LonTbn(Wc.Name) ' Wc.Name should be @*
    .Refresh BackgroundQuery:=False
End With
End Sub

Function WcsyFx(Fx) As String()
Dim B As Workbook: Set B = Xls.Workbooks.Open(Fx)
WcsyFx = WcsyOleWb(B)
B.Close False
End Function
