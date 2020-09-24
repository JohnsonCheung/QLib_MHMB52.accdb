Attribute VB_Name = "MxVb_Str_Fmt_ParChd"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Fmt_ParChd."

Function ParChdrenTmly_ParChdLy(ParChdLy$()) As String()
Dim Pary$()
Dim Chdreny()
Dim P$, C$
Dim ParChd: For Each ParChd In Itr(ParChdLy)
    AsgT1r ParChd, P, C
    WPushParChd P, C, Pary, Chdreny
Next
Dim J&: For J = 0 To UB(Chdreny)
    Chdreny(J) = AySrtQ(Chdreny(J))
Next
Dim Ixy&: Stop 'Ixy = IxySrtAy(Pary)
For J = 0 To UB(Pary)
    Dim Ix&: Stop 'Ix = Ixy(J)
    PushI ParChdrenTmly_ParChdLy, Pary(Ix) & " " & JnSpc(Chdreny(Ix)) '<===
Next
End Function
Private Function WPushParChd(P$, C$, OPary$(), OChdreny())
Dim Ix&: Ix = IxEle(OPary, P)
If Ix = -1 Then
    PushI OPary, P
    PushI OChdreny, Sy(C)
Else
    PushI OChdreny(Ix), C
End If
End Function
