Attribute VB_Name = "MxDta_Fun_Cii"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Fun_Cii."

Function CiiColnn$(Fny$(), Colnn$) '#Cii:Cml:Column-Index-in-format-of-Ss-of-inter-index# Index is started from 0
CiiColnn = JnSpc(CiyColnn(Fny, Colnn))
End Function
Function CiyColnn(Fny$(), Colnn$) As Integer() '#Cii:Cml:Column-Index-in-format-of-Array-of-integer-index# Index is started from 0
If Colnn = "" Then Exit Function
Dim N: For Each N In Tmy(Colnn)
    PushI CiyColnn, IxEle(Fny, N)
Next
End Function
