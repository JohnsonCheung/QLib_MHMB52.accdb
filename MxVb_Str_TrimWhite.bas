Attribute VB_Name = "MxVb_Str_TrimWhite"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_TrimWhite."

Function TrimWhite$(S):   TrimWhite = TrimWhiteL(TrimWhiteR(S)): End Function
Function TrimWhiteL$(S): TrimWhiteL = RmvFstN(S, WLenPfx(S)):    End Function
Function TrimWhiteR$(S): TrimWhiteR = RmvLasN(S, WLenSfx(S)):    End Function
Private Function WLenPfx&(S)
Dim J&: For J = 1 To Len(S)
    If Not IsWhiteChr(Mid(S, J, 1)) Then
        WLenPfx = J - 1
        Exit Function
    End If
Next
WLenPfx = Len(S)
Stop
End Function
Private Function WLenSfx&(S)
Dim L&: L = Len(S)
Dim J&: For J = L To 1 Step -1
    If Not IsWhiteChr(Mid(S, J, 1)) Then
        WLenSfx = L - J
        Exit Function
    End If
Next
WLenSfx = L
Stop
End Function
