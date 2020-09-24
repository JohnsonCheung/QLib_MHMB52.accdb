Attribute VB_Name = "MxIde_Mth_Slm_zTool_DmpMdnyNeedSlmAli"
Option Compare Text
Option Explicit
Sub slmzDmpMdnyNeedAlI()
Debug.Print
Debug.Print "B_MdnyNeedAliSlm"
Dim Ny$(): Ny = AmAli(AmQuoDbl(MdnyNeedAliSlm))
Dim SfxOf$: SfxOf = " of " & Si(Ny)
Dim N: For Each N In Itr(AwFstN(SySrtQ(Ny), 50))
    Dim J%: J = J + 1
    Dim Sfx$: Sfx = "' " & J & " " & SfxOf
    Debug.Print "AliSlmzMdn" & N & Sfx
Next
End Sub
Private Function MdnyNeedAliSlm() As String(): MdnyNeedAliSlm = NyNlyy(MsrcySlmPC):  End Function
Sub slmzVcMsrcyPC():                                            BrwMsrcy MsrcySlmPC: End Sub
