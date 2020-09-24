Attribute VB_Name = "MxIde_Src_Op_EnsUdt"
Option Compare Text
Const CMod$ = "MxIde_Src_Ens_Udt."
Option Explicit

Function SrcEnsUdt(Src$(), CdlUdt$, UdtnDlt$, Optional UdtnAft$) As String()
'The Udtn of @CdlUdt should be @UdtnDlt
Const CSub$ = CMod & "SrcEnsUdt"
Dim Dcl$(): Dcl = DclSrc(Src)
Dim BeiDlt As Bei: BeiDlt = BeiUdt(Dcl, UdtnDlt)
If CdlUdt = "" Then
    SrcEnsUdt = AeBei(Src, BeiDlt) '<== No ins, just dlt
ElseIf HasSsub(JnCrLf(Src), CdlUdt) Then
    SrcEnsUdt = Src  '<= Has ins and same as in Src, just ret
Else
    '== Dlt & Ins
    Dim O$()
    O = AeBei(Src, BeiDlt)  '<===
    Dim IxIns%
    SrcEnsUdt = AyInsAy(Src, SplitCrLf(CdlUdt), IxIns) '<==
End If
End Function
Private Function IxInsUdt%(Dcl$(), UdtnInsAft$)
If UdtnInsAft = "" Then
    IxInsUdt = Si(Dcl)
Else
    Dim A As Bei: A = BeiUdt(Dcl, UdtnInsAft)
    If IsEmpBei(A) Then Thw CSub, "Given @UdtnInsAft not found in given @Dcl", "UdtnInsAft Dcl", UdtnInsAft, Dcl
    IxInsUdt = A.Eix + 1
End If
End Function

