Attribute VB_Name = "MxXls_Op__Dao"
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Op__Dao."

Private Sub B_CrtFxDbC()
Dim P$: P = PthTmpFdr("CrtFxDbC")
ClrPth P
Dim FxRs$: FxRs = P & "ByRs.xlsx"
Dim Fxwc$: Fxwc = P & "ByWc.xlsx"
Dim FxSq$: FxSq = P & "BySq.xlsx"
CrtFxDbC FxRs, eLoAddTblWayRs
CrtFxDbC Fxwc, eLoAddTblWayWc
CrtFxDbC FxSq, eLoAddTblWaySq
End Sub
Sub CrtFxDbC(Fx$, Optional Way As eLoAddTblWay):               CrtFxDb Fx, CDb:   End Sub
Sub CrtFxDb(Fx$, D As Database, Optional Way As eLoAddTblWay): WbDb(D).SaveAs Fx: End Sub

Function WbDbC(Optional Way As eLoAddTblWay) As Workbook:       Set WbDbC = WbDb(CDb):         End Function
Function WbDbOupC(Optional Way As eLoAddTblWay) As Workbook: Set WbDbOupC = WbDbOup(CDb, Way): End Function
Function WbDbInpC(Optional Way As eLoAddTblWay) As Workbook: Set WbDbInpC = WbDbInp(CDb, Way): End Function
Function WbDbTmpC(Optional Way As eLoAddTblWay) As Workbook: Set WbDbTmpC = WbDbTmp(CDb, Way): End Function
Function WbDbHshC(Optional Way As eLoAddTblWay) As Workbook: Set WbDbHshC = WbDbHsh(CDb, Way): End Function

Function WbDb(D As Database, Optional Way As eLoAddTblWay) As Workbook:       Set WbDb = X_Wb(D, Tny(D), Way):    End Function
Function WbDbOup(D As Database, Optional Way As eLoAddTblWay) As Workbook: Set WbDbOup = X_Wb(D, TnyOup(D), Way): End Function
Function WbDbInp(D As Database, Optional Way As eLoAddTblWay) As Workbook: Set WbDbInp = X_Wb(D, TnyInp(D), Way): End Function
Function WbDbTmp(D As Database, Optional Way As eLoAddTblWay) As Workbook: Set WbDbTmp = X_Wb(D, TnyTmp(D), Way): End Function
Function WbDbHsh(D As Database, Optional Way As eLoAddTblWay) As Workbook: Set WbDbHsh = X_Wb(D, TnyHsh(D), Way): End Function

Function WbFbOup(Fb, Optional Way As eLoAddTblWay) As Workbook: Set WbFbOup = WbDbOup(Db(Fb), Way): End Function
Private Function X_Wb(D As Database, Tny$(), Optional Way As eLoAddTblWay) As Workbook
Dim B As Workbook: Set B = WbNw
Dim T: For Each T In Itr(Tny)
    WsTblWb B, D, T, Way
Next
DltWsIf B, "Sheet1"
Set X_Wb = B
End Function
