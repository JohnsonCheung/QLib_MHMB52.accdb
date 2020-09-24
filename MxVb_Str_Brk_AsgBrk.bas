Attribute VB_Name = "MxVb_Str_Brk_AsgBrk"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_Brk_AsgBrk."
Sub AsgBrkSpc(S, OA$, OB$):                                                 AsgS12 BrkSpc(S), OA, OB:                       End Sub
Sub AsgBrk1Dot(S, OA$, OB$, Optional NoTrim As Boolean):                    AsgS12 Brk1Dot(S, NoTrim), OA, OB:              End Sub
Sub AsgBrkDot(S, OA$, OB$, Optional NoTrim As Boolean):                     AsgS12 BrkDot(S, NoTrim), OA, OB:               End Sub
Sub AsgBrk1(S, Sep$, Optional O1, Optional O2, Optional NoTrim As Boolean): AsgS12 Brk1(S, Sep, NoTrim), O1, O2:            End Sub
Sub AsgBrk(S, Sep$, O1, O2, Optional NoTrim As Boolean):                    AsgBrkAt S, InStr(S, Sep), Sep, O1, O2, NoTrim: End Sub



Sub AsgBrkAt(S, At&, Sep$, O1, O2, Optional NoTrim As Boolean)
Const CSub$ = CMod & "AsgBrkAt"
If At = 0 Then
    Thw CSub, "String does not have Sep", "Str Sep At NoTrim", S, Sep, At, NoTrim
    Exit Sub
End If
O1 = Left(S, At - 1)
O2 = Mid(S, At + Len(Sep))
If Not NoTrim Then
    O1 = Trim(O1)
    O2 = Trim(O2)
End If
End Sub

Sub AsgBrk2(S, Sep$, O1, O2, Optional NoTrim As Boolean)
AsgS12 Brk2(S, Sep, NoTrim), O1, O2
End Sub
