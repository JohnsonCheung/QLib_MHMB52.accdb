Attribute VB_Name = "MxDta_Enm_Upd"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Enm_Upd."
Enum eRUpd: eRUpdRpt: eRUpdBoth: eRUpdUpd: End Enum
Public Const EnmqssRUpd$ = "eRUpd? Rpt Both Upd"
Enum eHdr: eHdrYes: eHdrNo: End Enum
Function EnmsyUpd() As String()
Static X$(): If Si(X) = 0 Then X = SySs(EnmqssRUpd)
EnmsyUpd = X
End Function
Function EnmsUpd$(Upd As eRUpd): EnmsUpd = EnmsyUpd()(Upd): End Function

Function IsRpt(Upd As eRUpd) As Boolean
Select Case Upd
Case eRUpdRpt, eRUpdBoth: IsRpt = True
End Select
End Function

Function IsUpd(Upd As eRUpd) As Boolean
Select Case True
Case Upd = eRUpdBoth, Upd = eRUpdUpd: IsUpd = True
End Select
End Function
