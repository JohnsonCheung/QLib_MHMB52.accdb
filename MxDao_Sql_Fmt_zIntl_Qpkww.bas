Attribute VB_Name = "MxDao_Sql_Fmt_zIntl_Qpkww"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Sql_Kww."
Function Qpkwwy() As String()
Static S$(): If Si(S) = 0 Then S = WS__KwwySrt(Tmy(EnmttmlQpt)) ' Qpkwwy need to sort by NKw in descending order by !WS__KwwySrt
Qpkwwy = S
End Function
Function Qpkwyy() As Variant()
Dim Kww: For Each Kww In Qpkwwy
    PushI Qpkwyy, SplitSpc(Kww)
Next
End Function
Private Function WS__KwwySrt(Kwwy$()) As String()
Dim NKwAy() As Byte
    Dim Kww: For Each Kww In Kwwy
        PushI NKwAy, Si(SplitSpc(Kww))
    Next
Dim NKwMax%: NKwMax = EleMax(NKwAy)
Dim O$()
Dim J%: For J = NKwMax To 1 Step -1
    PushIAy O, WS_KwwyWhIKw(Kwwy, NKwAy, J)
Next
WS__KwwySrt = O
End Function
Private Function WS_KwwyWhIKw(Kwwy$(), NKwAy() As Byte, NKw%) As String()
Dim O$()
Dim J%: For J = 0 To UB(NKwAy)
    If NKwAy(J) = NKw Then PushI O, Kwwy(J)
Next
WS_KwwyWhIKw = SySrtQ(O)
End Function

