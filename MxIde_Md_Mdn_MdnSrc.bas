Attribute VB_Name = "MxIde_Md_Mdn_MdnSrc"
Option Compare Text
Option Explicit

Function MdnySrcPC(Optional Ssubss$, Optional C As eCas) As String():                MdnySrcPC = MdnySrcP(CPj, Ssubss, C):        End Function
Function MdnySrcP(P As VBProject, Optional Ssubss$, Optional C As eCas) As String():  MdnySrcP = AwSsub(WMdnySrcP(P), Ssubss, C): End Function
Private Function WMdnySrcP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    If IsCmpSrc(C) Then PushI WMdnySrcP, C.Name
Next
End Function
