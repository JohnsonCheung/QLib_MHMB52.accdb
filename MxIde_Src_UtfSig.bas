Attribute VB_Name = "MxIde_Src_UtfSig"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_UtfSig."
Public Const Utf8Sig$ = "﻿"

Function RmvUtf8Sig$(S$)
RmvUtf8Sig = RmvPfx(S, Utf8Sig)
End Function

Private Sub B_HasUtfSig8()
Dim F$: F = LinesFt(resFfn("DrsTMthP\"))
Debug.Assert HasUtf8Sig(F)
End Sub

Function HasUtf8Sig(Ft$) As Boolean
HasUtf8Sig = HasPfx(FstNChrFfn(Ft, 3), Utf8Sig, vbBinaryCompare)
End Function
