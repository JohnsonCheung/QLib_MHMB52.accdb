Attribute VB_Name = "MxVb_Str_StrSS"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_SS."
Private Sub B_SssStrPfxx()
GoSub T1
Exit Sub
Dim Ss$, Exp$(), Pfx$
T1:
    Ss = "B C D"
    Pfx = "A"
    Ept = "AB AC AD"
    GoTo Tst
Tst:
    Act = SssStrPfxx(Pfx, Ss)
    C
    Return
End Sub

Function SssStrPfxx$(Pfx$, Ss): SssStrPfxx = JnSpc(AmAddPfx(SySs(Ss), Pfx)): End Function 'Add @Pfx to @SS

Function SySs(Ss) As String(): SySs = SplitSpc(RplDblSpc(Trim(Ss))): End Function

Function SyAddSS(Sy$(), Ss$) As String(): Stop 'SyAddSS = SyAp(Sy, Tml(Ss)): End Function

End Function
