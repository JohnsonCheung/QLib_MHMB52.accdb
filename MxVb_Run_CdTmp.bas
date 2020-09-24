Attribute VB_Name = "MxVb_Run_CdTmp"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Run_CdTmp."

Sub RunCd(Cd$()): RunCdl JnCrLf(Cd): End Sub

Sub RunCdl(Cdl$)
Dim N$: N = Tmpn("TmpMth_")
AddMthSub WMd, N, Cdl
Run N
End Sub

Private Function WMd() As CodeModule
Const Mdn$ = "ZTmpModForTmpCd"
Static X As Boolean: If Not X Then EnsMod CPj, Mdn
Set WMd = Md(Mdn)
End Function

Sub RunPPIx(FunPPIx$, P1, P2, Bix&, Eix&) ' Run @FunPPI (@Eix-@Bix)+1 times.  @FunPPI takes 3 parameters: @P1,@P2 and *Ix which is running from @Bix to @Eix
Dim I&: For I = Bix To Eix
    Run FunPPIx, P1, P2, I
Next
End Sub
