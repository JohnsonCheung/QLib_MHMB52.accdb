Attribute VB_Name = "MxDao"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao."
Function DbEng() As DBEngine: Set DbEng = Dao.DBEngine: End Function

Sub RplLoCnFbt(Lo As ListObject, Fb, T)
With Lo.QueryTable
    RplWc .Connection, Fb '<==
    .CommandType = xlCmdTable
    .CommandText = T '<==
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .BackgroundQuery = False
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .PreserveColumnINF = True
    .ListObject.DisplayName = LonTbn(T) '<==
    .Refresh BackgroundQuery:=False
End With
End Sub

Function WbTmpInp(D As Database) As Workbook: Set WbTmpInp = WbTny(D, TnyTmpInp(D)): End Function

Function WbTny(D As Database, Tny$()) As Workbook
Dim WbO As Workbook: Set WbO = WbNw
Dim T: For Each T In Itr(Tny)
    WsTblWb WbO, D, T
Next
DltSheet1 WbO
Set WbTny = WbO
End Function
Function WbFb(Fb) As Workbook
Dim D As Database: Set D = Db(Fb)
Set WbFb = WbTny(D, Tny(D))
End Function
Function HasWsn(B As Workbook, Wsn) As Boolean: HasWsn = HasItn(B.Sheets, Wsn): End Function
Sub SetWsn(S As Worksheet, Nm$)
Const CSub$ = CMod & "SetWsn"
If Nm = "" Then Exit Sub
If S.Name = Nm Then Exit Sub
If HasWsn(WbWs(S), Nm) Then
    Dim B As Workbook: Set B = WbWs(S)
    Thw CSub, "Wsn exists in Wb", "Wsn Wbn Wny-in-Wb", Nm, B.Name, Wny(B)
End If
S.Name = Nm
End Sub


Sub AutoFitLo(A As ListObject)
A.DataBodyRange.EntireColumn.AutoFit
End Sub

Sub CrtFxTny(Fx, Db As Database, Tny$()): WbTny(Db, Tny).SaveAs Fx: End Sub
