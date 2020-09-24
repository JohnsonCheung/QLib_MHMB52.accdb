VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHO_MHO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit
Const CMod$ = "MxMHO_MHO."
Public PthLgs$, PthLgsi$, Pth$, PthAr$
Public MHOMB52 As New MxMHO_MHOMB52
Public MHORelCst As New MxMHO_MHORelCst
Public MHOTaxCpr As New MxMHO_MHOTaxCpr
Public MHOTaxAlert As New MxMHO_MHOTaxAlert
Public MHOStmtE As New MxMHO_MHOStmtE
Public MHODuty As New MxMHO_MHODuty
Dim A() As Access.Application
Function PjMB52() As VBProject: Set PjMB52 = PjMainAcs(AcsMHMB52): End Function
Function PjDuty() As VBProject: Set PjDuty = PjMainAcs(AcsMHDuty): End Function

Function AcsMHMB52() As Access.Application:         Set AcsMHMB52 = WAcs(MHO.MHOMB52.FbPgm):     End Function
Function AcsMHDuty() As Access.Application:         Set AcsMHDuty = WAcs(MHO.MHODuty.FbPgm):     End Function
Function AcsMHStmtE() As Access.Application:       Set AcsMHStmtE = WAcs(MHO.MHOStmtE.FbPgm):    End Function
Function AcsMHTaxAlert() As Access.Application: Set AcsMHTaxAlert = WAcs(MHO.MHOTaxAlert.FbPgm): End Function
Function AcsMHTaxCpr() As Access.Application:     Set AcsMHTaxCpr = WAcs(MHO.MHOTaxCpr.FbPgm):   End Function

Private Function WAcs(Fb$) As Access.Application
With WIxOp(A, Fb)
    If .Som Then
        Set WAcs = A(.I)
        Exit Function
    End If
End With
PushObj A, AcsFb(Fb)
Set WAcs = EleLas(A)
End Function

Private Function WIxOp(A() As Access.Application, Fb$) As Intoptt
Dim Ix%, I: For Each I In Itr(A)
    If HasFb(CvAcs(I), Fb) Then WIxOp = SomInt(Ix): Exit Function
    Ix = Ix + 1
Next
End Function

Private Sub Class_Initialize()
Pth = "C:\Users\Public\"
PthLgs = Pth & "Logistic\"
PthLgsi = PthLgs & "SapData\"
PthAr = Pth & "DebtorAging4 and ARStmt\"
End Sub
