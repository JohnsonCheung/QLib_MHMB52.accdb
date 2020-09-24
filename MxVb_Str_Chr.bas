Attribute VB_Name = "MxVb_Str_Chr"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Chr."

Function ChrLas$(S):    ChrLas = Right(S, 1):   End Function
Function ChrFst$(S):    ChrFst = Left(S, 1):    End Function
Function ChrAt$(S, At):  ChrAt = Mid(S, At, 1): End Function
Function ChrThd$(S):    ChrThd = Mid(S, 3, 1):  End Function
Function ChrSnd$(S):    ChrSnd = Mid(S, 2, 1):  End Function
