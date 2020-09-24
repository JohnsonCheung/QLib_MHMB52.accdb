VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHO_MHOMB52"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit
Const CMod$ = "MxMHO_MHOMB52."
Public FbPgm$, FbDta$, TpMB52$, TpSHld$, FxiSalTxt$, Pth$
Private WCnDta As ADODB.Connection, WCnPgm As ADODB.Connection
Private WDbDta As Database, WDbPgm As Database
Property Get DbPgm() As Database
If IsNothing(WDbPgm) Then Set WDbPgm = Db(FbPgm)
Set DbPgm = WDbPgm
End Property
Property Get DbDta() As Database:
If IsNothing(WDbDta) Then Set WDbDta = Db(FbDta)
Set DbDta = WDbDta
End Property
Private Sub Class_Initialize()
Dim P$: P = MHO.PthLgs
Dim Pi$: Pi = MHO.PthLgsi
Pth = P
FbPgm = P & "StockHolding 8.accdb"
FbDta = P & "StockHolding 8.accdb"
FxiSalTxt = Pi & "sales text.XLSX"
End Sub

Private Sub Class_Terminate()
ClsDbAp WDbDta, WDbPgm
ClsCnAp WCnPgm, WCnDta
End Sub
