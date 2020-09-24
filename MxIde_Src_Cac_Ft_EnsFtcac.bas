Attribute VB_Name = "MxIde_Src_Cac_Ft_EnsFtcac"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_SrcFtcac_Pth."

Sub EnsFtcacPC(): EnsFtcacP CPj: End Sub
Sub EnsFtcacP(P As VBProject)
If LinesFtIf(FtCacMD5P(P)) = MD5P(P) Then Debug.Print "EnsFtcacP: Pj is already cached: " & P.Name: Exit Sub
ClrFtcacExcessP P
Dim Mdny$(): Mdny = MdnyFtcacOutP(P)
Dim M: For Each M In Itr(Mdny)
    CacM MdP(P, M)
Next
WrtStr MD5P(P), FtCacMD5P(P), OvrWrt:=True
End Sub

Sub EnsFtcacMC():      EnsFtcacM CMd:        End Sub
Sub EnsFtcacMdn(Mdn$): EnsFtcacM MdMdn(Mdn): End Sub
Private Sub EnsFtcacM(M As CodeModule)
If Not IsFtcacM(M) Then CacM M
End Sub
Private Sub CacM(M As CodeModule)
Dim S$(): S = SrcM(M)
Dim N$: N = Mdn(M)
WrtAy S, FtCacM(M), OvrWrt:=True
Dim A$(): A = TsyMi8CmntfbelS(S, N, ShtCmpTyM(M))
If Si(A) > 0 Then
    WrtAy A, FtCacMt8M(M), OvrWrt:=True
Else
    DltFfnIf FtCacMt8M(M)
End If
Debug.Print CSub; ": Cached: "; N; StrTrue(Si(A) = 0, " <==== no method")
End Sub
