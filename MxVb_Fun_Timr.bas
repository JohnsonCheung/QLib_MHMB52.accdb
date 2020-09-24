Attribute VB_Name = "MxVb_Fun_Timr"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fun_Timer."
Private WTimrBeg!, WTimrLas!, WMsg$
Private Sub B_BegTimr()
BegTimr
ShwTimr "sdflk"
Wait 1
ShwTimr "sdlfkj"
End Sub
Sub BegTimr(Optional Msg = ""): WTimrBeg = Timer: WTimrLas = WTimrBeg: WSetMsg Msg: End Sub
Sub ShwTimr(Optional Msg = ""): ShwTimrNoLf Msg: Debug.Print:                       End Sub
Sub ShwTimrNoLf(Optional Msg = "")
WSetMsg Msg
Dim SCumm$, SIntv$, T!
    T = Timer
    SCumm = AliR(Format(T - WTimrBeg, "#,##0.000"), 9)
    SIntv = AliR(Format(T - WTimrLas, "#,##0.000"), 9)
Debug.Print SCumm & SIntv & "(s) " & WMsg;
WTimrLas = T
End Sub
Private Sub WSetMsg(Msg)
If Msg <> "" Then WMsg = Msg
End Sub
Private Sub B_ShwTimr(): TimFun "SampFunA SampFunB": End Sub
Sub TimFun(FunNN)
Dim B!, E!, F
For Each F In SySs(FunNN)
    B = Timer
    Run F
    E = Timer
    Debug.Print F, "<-- Run"; E - B
Next
End Sub
