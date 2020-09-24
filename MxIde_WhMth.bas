Attribute VB_Name = "MxIde_WhMth"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_WhMth."
Type WhMd: HasDta As Boolean: SegyMd() As String: SsubyMd() As String: End Type
Type WhPm: HasDta As Boolean: End Type
Type WhMth
    HasDta As Boolean
    RyMthn() As RegExp
    WhPm As WhPm
    SsubRet As String
    HasAp As eTri
End Type
Type WhMdMth
    WhMd As WhMd
    WhMth As WhMd
End Type
Function AskWhMth() As WhMth

End Function
Function WhMthzPms(Pms$) As WhMth

End Function
