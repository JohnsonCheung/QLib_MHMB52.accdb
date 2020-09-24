Attribute VB_Name = "MxVb_Run_Thw_Raise"
Option Compare Text
Const CMod$ = "MxVb_Run_Thw_Raise."
Option Explicit
Sub RaiseNotePad(Optional Fun$)
Dim A$: If Fun <> "" Then A = " (source is in subroutine-[" & Fun & "])"
Err.Raise 0, , "Please see the message in the NotePad." & A
End Sub
Sub RaiseQQ(MsgQQ$, ParamArray Ap()): Dim Av(): Av = Ap: Raise FmtQQAv(MsgQQ, Av): End Sub
Sub Raise(Msg$):                      Err.Raise 1, , Msg:                          End Sub
Sub RaiseTrue(IfTrue As Boolean, Msg$)
If IfTrue Then Raise Msg
End Sub
Sub RaisePgmEr():       Raise "PgmEr":            End Sub
Sub RaiseMsg(Msg$):     BrwStr Msg: RaiseNotePad: End Sub
Sub RaiseMsgy(Msgy$()): BrwAy Msgy: RaiseNotePad: End Sub

Sub ChkEry(Er$(), Optional Fun$, Optional Tit$): ChkEr JnCrLf(Er), Fun, Tit: End Sub
Sub ChkEr(Er$, Optional Fun$, Optional Tit$)
If Trim(Er) = "" Then Exit Sub
BrwStr Boxl(Tit) & Er, "Er_"
RaiseNotePad Fun
End Sub
