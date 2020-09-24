VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHO_MHODuty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit
Const CMod$ = "MxMHO_MHODuty."
Public FbPgm$, FbDta$, Pth$
Private WDbPgm As Database, WDbDta As Database
Private WCnPgm As ADODB.Connection, WCnDta As ADODB.Connection
Property Get DbPgm() As Database:         Set DbPgm = DbIf(WDbPgm, FbPgm): End Property
Property Get DbDta() As Database:         Set DbDta = DbIf(WDbDta, FbDta): End Property
Property Get CnPgm() As ADODB.Connection: Set CnPgm = CnIf(WCnPgm, FbPgm): End Property
Property Get CnDta() As ADODB.Connection: Set CnDta = CnIf(WCnDta, FbDta): End Property

Private Sub Class_Initialize()
Pth = MHO.PthLgs & "DutyPrepay7\"
FbDta = Pth & "DutyPrepay7_Data.accdb"
FbPgm = Pth & "DutyPrepay7.accdb"
End Sub

Private Sub Class_Terminate()
ClsDbAp WDbPgm, WDbDta
ClsCnAp WCnPgm, WCnDta
End Sub
