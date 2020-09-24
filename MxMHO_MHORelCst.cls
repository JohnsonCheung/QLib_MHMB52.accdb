VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHO_MHORelCst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit
Const CMod$ = "MxMHO_MHORelCst."
Public FbPgm$, Pth$
Private WDbPgm As Database
Property Get DbPgm() As Database: Set DbPgm = DbIf(WDbPgm, FbPgm): End Property

Private Sub Class_Initialize()
Dim P$: P = MHO.PthLgs
Pth = P & "RelCst\"
FbPgm = P & "RelCst.accdb"
End Sub
