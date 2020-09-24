Attribute VB_Name = "MxDta_Dotly"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Dotly."

Function DyDotly(Dotly$()) As Variant()
Dim I: For Each I In Itr(Dotly)
    PushI DyDotly, SplitDot(I)
Next
End Function

Sub BrwDotly(Dotly$()): Brw FmtDotly(Dotly): End Sub
Sub VcDotly(Dotly$()):  Vc FmtDotly(Dotly):  End Sub
Sub DmpDotly(Dotly$()): Dmp FmtDotly(Dotly): End Sub

Function FmtDot(Dotly$()) As String():       FmtDot = FmtLndy(DyDotly(Dotly)):     End Function
Function FmtDotly(Dotly$()) As String():   FmtDotly = FmtLndy(DyDotly(Dotly)):     End Function
Function FmtDot2ly(Dotly$()) As String(): FmtDot2ly = FmtLndy(Dy2colDotly(Dotly)): End Function
Function Dy2colDotly(Dotly) As Variant()
Dim L: For Each L In Itr(Dotly)
    PushI Dy2colDotly, STup2S12(BrkDot(L))
Next
End Function
