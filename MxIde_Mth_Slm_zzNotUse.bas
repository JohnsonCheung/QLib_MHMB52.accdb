Attribute VB_Name = "MxIde_Mth_Slm_zzNotUse"
Option Compare Text
Option Explicit

Function SlmbyAliSrc(Src$(), Optional InlAli As Boolean) As Variant()
Dim I, Slmb$(): For Each I In Itr(SlmbySrc(Src))
    Slmb = I
    With SlmboptAli(Slmb)
        If .Som Then
            PushI SlmbyAliSrc, .Ly
        Else
            If InlAli Then PushI SlmbyAliSrc, Slmb
        End If
    End With
Next
End Function
Function CpmdloptAliSlmzCmp(C As VBComponent) As Cpmdlopt
Dim Src$(): Src = SrcCmp(C)
With SrcoptSlm(Src)
    If .Som Then
        CpmdloptAliSlmzCmp = SomCpmdl(Cpmdl(C.Name, JnCrLf(Src), JnCrLf(.Ly)))
    End If
End With
End Function
