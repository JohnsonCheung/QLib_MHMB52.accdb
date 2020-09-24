Attribute VB_Name = "MxIde_Src_Cac_Ft_zIntl_FtcacFun"
Option Compare Text
Option Explicit
Public Const FfnspecFtcac$ = "*" & FnsfxFtcac & ".txt"
Public Const FfnspecFtcacMit8$ = "*" & FnsfxFtcacMit8 & ".txt"

Function PthFtcacP$(P As VBProject)
Static X As New Dictionary
Dim F$: F = P.FileName
If Not X.Exists(F) Then X.Add F, PthAss(F) & ".cache\"
PthFtcacP = X(F)
End Function

Function FtCacMdnP$(P As VBProject, Mdn): FtCacMdnP = PthFtcacP(P) & Mdn & FnsfxFtcac & ".txt":        End Function
Function FtCacM$(M As CodeModule):           FtCacM = PthFtcacM(M) & Mdn(M) & FnsfxFtcac & ".txt":     End Function
Function FtCacMD5P$(P As VBProject):      FtCacMD5P = PthFtcacP(P) & P.Name & "(MD5).txt":             End Function
Function FtCacMt8M$(M As CodeModule):     FtCacMt8M = PthFtcacM(M) & Mdn(M) & FnsfxFtcacMit8 & ".txt": End Function
Function PthFtcacM$(M As CodeModule):     PthFtcacM = PthFtcacP(PjM(M)):                               End Function

Function FfnyFtcacP(P As VBProject) As String():         FfnyFtcacP = Ffny(PthFtcacP(P), FfnspecFtcac):     End Function
Function FfnyFtcacMit8P(P As VBProject) As String(): FfnyFtcacMit8P = Ffny(PthFtcacP(P), FfnspecFtcacMit8): End Function
Function FnayFtcacP(P As VBProject) As String():         FnayFtcacP = Fnay(PthFtcacP(P), FfnspecFtcac):     End Function
Function FnayFtcacMit8P(P As VBProject) As String(): FnayFtcacMit8P = Fnay(PthFtcacP(P), FfnspecFtcacMit8): End Function

Function NFtcacP%(P As VBProject):         NFtcacP = Si(FfnyFtcacP(P)):     End Function
Function NFtcacMit8P%(P As VBProject): NFtcacMit8P = Si(FfnyFtcacMit8P(P)): End Function
Function MdnyFtcacOutP(P As VBProject) As String()
Dim C As VBComponent: For Each C In P.VBComponents
    If Not IsFtcacM(C.CodeModule) Then PushI MdnyFtcacOutP, C.Name
Next
End Function
Function MdnyFtcacP(P As VBProject) As String():             MdnyFtcacP = FfnyCutExtRmvFnsfx(FnayFtcacP(P), FnsfxFtcac):         End Function
Function MdnyFtcacMit8P(P As VBProject) As String():     MdnyFtcacMit8P = FfnyCutExtRmvFnsfx(FnayFtcacMit8P(P), FnsfxFtcacMit8): End Function
Function MdnyFtcacMthYesP(P As VBProject) As String(): MdnyFtcacMthYesP = MdnyFtcacMit8P(P):                                     End Function
Function MdnyFtcacMthNoP(P As VBProject) As String():   MdnyFtcacMthNoP = SyMinus(MdnyFtcacP(P), MdnyFtcacMit8P(P)):             End Function

Sub TimEnsFtcacPC():                                    TimFun "EnsFtcacPC":                                           End Sub
Function MD5FtcacMC$():                    MD5FtcacMC = MD5FtcacM(CMd):                                                End Function
Function MD5FtcacM$(M As CodeModule):       MD5FtcacM = MD5Ft(FtCacM(M)):                                              End Function
Function IsFtCacMdn(Mdn) As Boolean:       IsFtCacMdn = IsFtcacM(Md(Mdn)):                                             End Function
Function IsFtCacMC() As Boolean:            IsFtCacMC = IsFtcacM(CMd):                                                 End Function
Sub PmptIsFtCacMC():                                    PmptIsFtCacM CMd:                                              End Sub
Private Sub PmptIsFtCacM(M As CodeModule):              MsgBox "Is cached?" & vbCrLf & Mdn(M) & vbCrLf2 & IsFtcacM(M): End Sub
Sub PmptIsFtCacMdn(Mdn):                                MsgBox "Is cached?" & vbCrLf & CMdn & vbCrLf2 & IsFtCacMC:     End Sub

Function IsFtcacM(M As CodeModule) As Boolean
Dim F$
    F = FtCacM(M): If NoFfn(F) Then Exit Function
    If FileLen(F) <> Len(SrclM(M)) Then Exit Function
IsFtcacM = LinesFt(F) = SrclM(M)
End Function
