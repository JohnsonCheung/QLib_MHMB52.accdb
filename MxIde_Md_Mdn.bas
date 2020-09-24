Attribute VB_Name = "MxIde_Md_Mdn"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Md_Mdn."

Function MdnyWhEIntl(Mdny$(), EI As eEIntl) As String()
Select Case True
Case EI = eEIntlBoth: MdnyWhEIntl = Mdny
Case EI = eEIntlI: MdnyWhEIntl = AwSsubssOr(Mdny, "_Intl_ _Tool_")
Case EI = eEItnlE: MdnyWhEIntl = AeSsubssOr(Mdny, "_Intl_ _Tool_")
Case Else: ThwEnm CSub, EI, EnmqssEIntl
End Select
End Function
Function MdnyPC(Optional PatnssAnd$) As String():                  MdnyPC = MdnyP(CPj, PatnssAnd):                End Function
Function ModnyPC(Optional PatnssAnd$) As String():                ModnyPC = ModnyP(CPj, PatnssAnd):               End Function
Function ClsnyPC(Optional PatnssAnd$) As String():                ClsnyPC = ClsnyP(CPj, PatnssAnd):               End Function
Function MdnyP(P As VBProject, Optional PatnssAnd$) As String():    MdnyP = AwPatnssAnd(WMdnyP(P), PatnssAnd):    End Function
Function ModnyP(P As VBProject, Optional PatnssAnd$) As String():  ModnyP = AwPatnssAnd(WModnyP(P), PatnssAnd):   End Function
Function ClsnyP(P As VBProject, Optional PatnssAnd$) As String():  ClsnyP = AwPatnssAnd(WClsnyP(CPj), PatnssAnd): End Function

Private Function WMdnyP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    If IsMd(C) Then PushI WMdnyP, C.Name
Next
End Function
Private Function WModnyP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    If IsMod(C) Then PushI WModnyP, C.Name
Next
End Function
Private Function WClsnyP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    If IsCls(C) Then PushI WClsnyP, C.Name
Next
End Function
