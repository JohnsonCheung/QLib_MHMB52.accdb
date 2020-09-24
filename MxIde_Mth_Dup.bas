Attribute VB_Name = "MxIde_Mth_Dup"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mth_Dup."
Public Const FfoDupMth$ = FfTMthc
Sub DmpDupMthFtcac(): Dmp Mi3yDupNmtP(CPj): End Sub
Private Function Mi3yDupNmtP(P As VBProject, Optional InlPrv As Boolean) As String()
Dim Dy1(): Dy1 = DyMi8CmntfbelModFtcacP(P)
'                      12    7
'                     01234567
Dim Dy2(): Dy2 = DyePrv(Dy1, InlPrv)
If Si(Dy2) = 0 Then Exit Function
Dim Dy3Nml(): Dy3Nml = DySel(Dy2, IntySs("2 1 7")) ' Mthn Mdn Mthln
Dim Dy4(): Dy4 = DywDup(Dy3Nml, 0)
Dim Dy5(): Dy5 = DySrtCii(Dy4, "0")
Dim Dy6(): Dy6 = DyAdjMdn(Dy5)
Mi3yDupNmtP = FmtLndy(Dy6)
End Function
Private Function DyAdjMdn(Dy3Nml()) As Variant() ' Add Mthn after Mdn
Dim Dr: For Each Dr In Itr(Dy3Nml)
    Dim Mthn$, Mdn$
    Mthn = Dr(0)
    Mdn = Dr(1)
    Dr(1) = Mdn & "." & Mthn
    PushI DyAdjMdn, Dr
Next
End Function
Private Function DyePrv(DyMi8Cmntfbel(), InlPrv As Boolean) As Variant()
If InlPrv Then DyePrv = DyMi8Cmntfbel: Exit Function
Dim Dr: For Each Dr In Itr(DyMi8Cmntfbel)
    If Dr(4) <> "Prv" Then PushI DyePrv, Dr
Next
End Function
