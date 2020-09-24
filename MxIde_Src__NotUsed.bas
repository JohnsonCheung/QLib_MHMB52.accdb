Attribute VB_Name = "MxIde_Src__NotUsed"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_SrcFtcac_Md."


Function SrcRmvLnAtr(Src$()) As String()
Dim N&: N = AtrVbLnCnt(Src$())
SrcRmvLnAtr = AeFstN(Src, N)
End Function

Function AtrVbLnCnt%(Src$())
Dim O%:
    Dim L: For Each L In Itr(Src)
        If Not HasPfx(L, "Attribute VB") Then Exit For
        O = O + 1
    Next
AtrVbLnCnt = O
End Function

Function RmvCls4Sigln(Src$()) As String()
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
If HasCls4Sigln(Src) Then
    RmvCls4Sigln = AeFstN(Src, 4)
Else
    RmvCls4Sigln = Src
End If
End Function

Function HasCls4Sigln(Src$()) As Boolean
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
If Si(Src) < 4 Then Exit Function
If Src(0) <> "VERSION 1.0 CLASS" Then Exit Function
If Src(1) <> "BEGIN" Then Exit Function
If HasPfx(Src(2), "  MultiUse =") Then Exit Function
If Src(3) = "End" Then Exit Function
HasCls4Sigln = True
End Function
