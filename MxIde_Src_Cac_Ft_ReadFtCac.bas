Attribute VB_Name = "MxIde_Src_Cac_Ft_ReadFtCac"
Option Compare Text
Option Explicit
Function TsyMi8CmntfbelFtcacPC() As String(): TsyMi8CmntfbelFtcacPC = TsyMi8CmntfbelFtcacP(CPj): End Function
Function TsyMi8CmntfbelFtcacP(P As VBProject) As String()
EnsFtcacP P
Dim Ffn: For Each Ffn In Itr(FfnyFtcacMit8P(P))
    PushIAy TsyMi8CmntfbelFtcacP, LyFt(Ffn)
Next
End Function
Function TsyMi8CmntfbelFtcaczMdn(Mdn$) As String():         TsyMi8CmntfbelFtcaczMdn = TsyMi8CmntfbelFtcacM(MdMdn(Mdn)): End Function
Function TsyMi8CmntfbelFtcacM(M As CodeModule) As String():    TsyMi8CmntfbelFtcacM = LyFtIf(FtCacMt8M(M)):             End Function
Function TsyMi8CmntfbelFtcacMC() As String():                 TsyMi8CmntfbelFtcacMC = TsyMi8CmntfbelFtcacM(CMd):        End Function
Function SrclFtcacMC$():                                                SrclFtcacMC = SrclFtcacM(CMd):                  End Function
Function SrclFtcacM$(M As CodeModule):                                   SrclFtcacM = SrclFtcacMdn(Mdn(M)):             End Function
Function SrclFtcacMdn$(Mdn):                                           SrclFtcacMdn = LinesFtIf(FtCacM(Md(Mdn))):       End Function
Function SrcFtcacMdn(Mdn) As String():                                  SrcFtcacMdn = SplitCrLf(SrclFtcacMdn(Mdn)):     End Function
Function MthyyFtcacMdn(Mdn) As Variant() ' Mthyy from cur Ftcac from a @Mdn
MthyyFtcacMdn = MthyySrc(SrcFtcacMdn(Mdn))
End Function
Function MthyyFtcacP(P As VBProject) As Variant()  ' All Mthyy from cur Ftcac
EnsFtcacP P
Dim Dy(), Dr, MdnCur$, MdnLas$, S$()
Dim Pth$: Pth = PthFtcacP(P)
Dim Ft: For Each Ft In Itr(FfnyFtcacMit8P(P))
    Dy = DyTsy(LyFt(Ft)) ' Always has Mth in @Ft, otherwise the Ft will not exist
    MdnCur = Dy(0)(1)
    If MdnCur <> MdnLas Then
        MdnLas = MdnCur
        S = LyFt(Pth & MdnCur & FnsfxFtcac & ".txt")
    End If
    For Each Dr In Dy
        PushI MthyyFtcacP, AwBE(S, Dr(5), Dr(6)) 'TslMi8Cmntfbel
                                                 '         56    5=Bix; 6=Eix
    Next
Next
End Function
Private Sub B_DyMi8CmntfbelModFtcacP()
Brw DyJn(DyMi8CmntfbelModFtcacP(CPj), , vbTab)
End Sub
Function DyMi8CmntfbelFtcacPC() As Variant():              DyMi8CmntfbelFtcacPC = DyMi8CmntfbelFtcacP(CPj):       End Function
Function DyMi8CmntfbelFtcacP(P As VBProject) As Variant():  DyMi8CmntfbelFtcacP = DyTsy(TsyMi8CmntfbelFtcacP(P)): End Function
Function DyMi8CmntfbelModFtcacP(P As VBProject) As Variant()
EnsFtcacP P
Dim Ffny1$(): Ffny1 = FfnyFtcacMit8P(P)
Dim Ffny2$(): Ffny2 = FfnyWhStd(Ffny1)
Dim Tsy$(): Tsy = LyFty(Ffny2)
DyMi8CmntfbelModFtcacP = DyTsy(Tsy)
End Function
Private Function FfnyWhStd(FfnyFtcacMit8$()) As String()
Const P = "Std" & vbTab
Dim Ffn: For Each Ffn In Itr(FfnyFtcacMit8)
    If HasPfx(Ln1Ft(Ffn), P) Then
        PushI FfnyWhStd, Ffn
    End If
Next
End Function
