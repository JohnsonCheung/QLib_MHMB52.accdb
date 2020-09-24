Attribute VB_Name = "MxAcs_AcFrm_AcFrm"
Option Compare Text
Option Explicit
Const CMod$ = "MxAcs_AcFrm_AcFrm."

Function CvFrm(A) As Access.Form: Set CvFrm = A: End Function
Function FrmFrmn(Frmn) As Access.Form
If Not IsFrmnOpn(Frmn) Then DoCmd.OpenForm Frmn, acDesign
Set FrmFrmn = Forms(Frmn)
End Function
Sub RequeryFrm(Frmn)
Dim F As Access.Form: For Each F In Forms
    If F.Name = Frmn Then
        If F.CurrentView = acCurViewFormBrowse Then
            F.Requery
            Exit Sub
        End If
    End If
Next
End Sub
Function IsFrmnOpn(Frmn) As Boolean
Dim F As Access.Form: For Each F In Forms
    If F.Name = Frmn Then
        If F.CurrentView = acCurViewFormBrowse Then
            IsFrmnOpn = True
            Exit Function
        End If
    End If
Next
End Function
Function CFrm() As Access.Form ' #Current-Opened-Form#
On Error Resume Next
Set CFrm = Access.Application.Screen.ActiveForm
End Function

Sub SetAcCtlnnVis(F As Access.Form, Ctlnn$, Vis As Boolean)
Dim N: For Each N In Split(Ctlnn)
    F.Controls(N).Visible = Vis
Next
End Sub

Sub OpnFrm(Frmn$): DoCmd.OpenForm Frmn: End Sub

Function HasCtl(F As Access.Form, Nm$) As Boolean
Dim C As Access.Control: For Each C In F.Controls
    If C.Name = Nm Then HasCtl = True: Exit Function
Next
End Function
