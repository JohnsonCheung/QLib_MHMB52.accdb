Attribute VB_Name = "MxVb_UI_Ask"
Option Compare Text
Option Explicit

Function AskDltFfn(Ffn) As Boolean
If Not HasFfn(Ffn) Then AskDltFfn = True: Exit Function
If Not AskYes(WMsgAskDlt(Ffn)) Then Exit Function
DltFfn Ffn
AskDltFfn = True
End Function
Private Function WMsgAskDlt$(Ffn)
Dim L1$: L1 = "File: [" & Fn(Ffn) & "] exist"
Dim L2$: L2 = "In folder: [" & Pth(Ffn) & "]"
Dim L3$: L3 = ""
Dim L4$: L4 = "Input [Yes] to delete it?"
WMsgAskDlt = LinesSap(L1, L2, L3, L4)
End Function
Function AskYes(Msg$, Optional Tit$ = "Input [Yes] to confirm") As Boolean: AskYes = UCase(InputBox(Msg, Tit)) = "YES": End Function
