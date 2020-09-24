Attribute VB_Name = "MxVb_Msg_VbEr"
Option Compare Text
Option Explicit


Function EryHdrInd(Hdr$, MsgyInd$()) As String()
If IsEmpAy(MsgyInd) Then Exit Function
PushI EryHdrInd, Hdr
PushIAy EryHdrInd, AyTab(MsgyInd)
End Function
