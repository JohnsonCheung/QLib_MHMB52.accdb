Attribute VB_Name = "MxVb_Dta_Di_FmtDi"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_DicFmt."

Sub VcDi(D As Dictionary, _
Optional InlValTy As Boolean, _
Optional TmlH12$ = "Key Val", _
Optional Fmt As eTblFmt, _
Optional Zer As eZer, _
Optional Wdt%, _
Optional IsValSum As Boolean, _
Optional IsValAliR As Boolean, _
Optional NoIx As Boolean, _
Optional PfxFn$ = "Dic_")
VcAy FmtDi(D, InlValTy, TmlH12, Fmt, Zer, Wdt, IsValSum, IsValAliR, NoIx), PfxFn
End Sub
Sub BrwDi(D As Dictionary, _
Optional InlValTy As Boolean, _
Optional H12$ = "Key Val", _
Optional Fmt As eTblFmt, _
Optional Zer As eZer, _
Optional Wdt%, _
Optional IsValSum As Boolean, _
Optional IsValAliR As Boolean, _
Optional NoIx As Boolean, _
Optional PfxFn$ = "Dic_")
VcAy FmtDi(D, InlValTy, H12, Fmt, Zer, Wdt, IsValSum, IsValAliR, NoIx), PfxFn
End Sub
Sub DmpDi(D As Dictionary, _
Optional InlValTy As Boolean, _
Optional H12$ = "Key Val", _
Optional Fmt As eTblFmt, _
Optional Zer As eZer, _
Optional Wdt%, _
Optional IsValSum As Boolean, _
Optional IsValAliR As Boolean, _
Optional NoIx As Boolean)
Dmp FmtDi(D, InlValTy, H12, Fmt, Zer, Wdt, IsValSum, IsValAliR, NoIx)
End Sub

Private Sub B_FmtDi()
Dmp FmtDi(W2SampDi1, TmlH12:="[My Key] [My Val]", Fmt:=eTblFmtTb, InlValTy:=True, IsValAliR:=True, IsValSum:=True)
End Sub
Private Function W2SampDi1() As Dictionary
Set W2SampDi1 = New Dictionary
W2SampDi1.Add "D", 1
W2SampDi1.Add "B", Now
W2SampDi1.Add "C", 3&
W2SampDi1.Add "Lines", RplVbl("ksf|lksdf|lsdfjlsd flksdj fdf|  sldkfj")
W2SampDi1.Add "Lines1", RplVbl("ksf|lksdf|lsdfjlsd flksdj fdf|  sldkfj")
W2SampDi1.Add "B1", Now
W2SampDi1.Add "B2", Now
W2SampDi1.Add "Lines2", RplVbl("ksf|lksdf|lsdfjlsd flksdj fdf|  sldkfj")
W2SampDi1.Add "C1", 3&
End Function
Function FmtDi(D As Dictionary, Optional InlValTy As Boolean, Optional TmlH12$ = "Key Val", _
Optional Fmt As eTblFmt, _
Optional Zer As eZer, _
Optional Wdt%, _
Optional IsValSum As Boolean, _
Optional IsValAliR As Boolean, _
Optional NoIx As Boolean) As String()
Dim Drs As Drs: Drs = DrsDi(D, InlValTy, TmlH12)
Dim ColnnAliR$: If IsValAliR Then ColnnAliR = RmvA1T(TmlH12)
Dim ColnnSum$: If IsValSum Then ColnnSum = RmvA1T(TmlH12)
FmtDi = FmtDrs(Drs, eRdsNo, Fmt, Zer, Wdt, ColnnAliR, ColnnSum, NoIx)
End Function
Private Function WAddValTy(D As Dictionary, InlValTy As Boolean) As Dictionary ' ret nwDic with optionally K added with VarTy
If Not InlValTy Then Set WAddValTy = D: Exit Function
Dim K$(): K = WNwKy(D)
Dim V$(): V = DivyStr(D)
Set WAddValTy = DiAy12(K, V)
End Function
Private Function WNwKy(D As Dictionary) As String() 'ret DikyStr with VarTy of Dii added
Dim K$(): K = AmAli(DikyStr(D))
Dim T$(): T = TynyDii(D)
WNwKy = DyAddDcStr12(K, T)
End Function
