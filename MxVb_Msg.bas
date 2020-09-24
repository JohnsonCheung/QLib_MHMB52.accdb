Attribute VB_Name = "MxVb_Msg"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Msg."

Function MsgBoxAy(Tit$, Ay) As String() ' Adding a boxing of title and following an Ay
If Tit = "" Then MsgBoxAy = Ay: Exit Function
MsgBoxAy = AyAdd(Box(Tit), Ay)
End Function

Function MsgHdrAy(Hdr$, Ay) As String()
If IsEmpAy(Ay) Then Exit Function
PushI MsgHdrAy, Hdr
PushIAy MsgHdrAy, AyTab(Ay)
End Function
Function MsgFxw(Fx$, W$) As String()
Dim O$()
PushI O, "Excel file : " & QuoSq(Fn(Fx))
PushI O, "Path       : " & QuoSq(Pth(Fx))
PushI O, "Worksheet  : " & QuoSq(W)
MsgFxw = O
End Function
