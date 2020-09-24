Attribute VB_Name = "MxAcs_AcCtl_TAcCtln"
Option Compare Text
Const CMod$ = "MxAcs_AcCtl_TAcCtln."
Option Explicit
Type TAcCtln: Btnny() As String: TglBtnny() As String: CBoxny() As String: TBoxny() As String: End Type
Function TAcCtlnFrmn(Frmn$) As TAcCtln: TAcCtlnFrmn = TAcCtlnFrm(FrmFrmn(Frmn)): End Function
Function TAcCtlnFrm(F As Access.Form) As TAcCtln
With TAcCtlnFrm
    .Btnny = AcBtnny(F)
    .TglBtnny = AcTglBtnny(F)
    .CBoxny = AcCBoxny(F)
    .TBoxny = AcTBoxny(F)
End With
End Function
Function AcBtnny(F As Access.Form) As String():       AcBtnny = SyAdd(AcCmdBtnny(F), AcTglBtnny(F)): End Function
Function AcCmdBtnny(F As Access.Form) As String(): AcCmdBtnny = ItnTyn(F.Controls, "CommandButton"): End Function
Function AcTglBtnny(F As Access.Form) As String(): AcTglBtnny = ItnTyn(F.Controls, "ToggleButton"):  End Function
Function AcCBoxny(F As Access.Form) As String():     AcCBoxny = ItnTyn(F.Controls, "CheckBox"):      End Function
Function AcTBoxny(F As Access.Form) As String():     AcTBoxny = ItnTyn(F.Controls, "TextBox"):       End Function
Function FmtTAcCtln(F As TAcCtln) As String()
Dim O$()
PushI O, "CommandButton": PushIAy O, LyTab4Spc(F.Btnny)
PushI O, "ToggleButton":  PushIAy O, LyTab4Spc(F.TglBtnny)
PushI O, "CheckBox":      PushIAy O, LyTab4Spc(F.CBoxny)
PushI O, "TextBox":       PushIAy O, LyTab4Spc(F.TBoxny)
FmtTAcCtln = O
End Function
Sub BrwAcCtln(Frmn$):                              Brw FmtAcCtln(Frmn):           End Sub
Sub VcAcCtln(Frmn$):                               Vc FmtAcCtln(Frmn):            End Sub
Function FmtAcCtln(Frmn$) As String(): FmtAcCtln = FmtTAcCtln(TAcCtlnFrmn(Frmn)): End Function
