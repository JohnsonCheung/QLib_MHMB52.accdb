Attribute VB_Name = "MxVb_Dta_Di_TfmDi"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Dta_Di_Tfm."

Function DrsDi(D As Dictionary, Optional InlValTy As Boolean, Optional H12$ = "Key Val") As Drs:  DrsDi = Drs(WFnyDi(H12, InlValTy), WDyDi(D, InlValTy)): End Function
Private Function WFnyDi(H12$, InlValTy As Boolean) As String():                                  WFnyDi = TmyAp(Tm1(H12), "TyVal", RmvTm1(H12)):          End Function
Private Function WDyDi(D As Dictionary, InlValTy As Boolean) As Variant()
If D.Count = 0 Then Exit Function
Dim K(): K = D.Keys
If Si(K) = 0 Then Exit Function
Dim I, Dr: For Each I In K
    If InlValTy Then
        Dr = Array(I, TypeName(D(I)), D(I))
    Else
        Dr = Array(I, D(I))
    End If
    Push WDyDi, Dr
Next
End Function
Function WsDi(Di As Dictionary, Optional InlValTy As Boolean, Optional Tit$ = "Key Val") As Worksheet
Set WsDi = WsDrs(DrsDi(Di, InlValTy))
End Function
