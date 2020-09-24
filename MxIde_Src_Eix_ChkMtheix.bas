Attribute VB_Name = "MxIde_Src_Eix_ChkMtheix"
Option Compare Text
Option Explicit

Sub ChkMtheix()
Dim O$()
Dim C As VBComponent: For Each C In CPj.VBComponents
    PushIAy O, WMsgyCmp(C)
Next
Dmp O
End Sub
Private Function WMsgyCmp(C As VBComponent) As String()
Dim S$(): S = SrcCmp(C)
Dim N$: N = C.Name
Dim I: For Each I In ItrMthix(S)
    If Not HasMtheix(S, I) Then PushIAy WMsgyCmp, Jrcy(N, I + 1, P12(1, 1), S(I))
Next
End Function
