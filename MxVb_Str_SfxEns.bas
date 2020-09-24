Attribute VB_Name = "MxVb_Str_SfxEns"
Option Compare Text
Option Explicit
Const TimMdy As Date = #9/18/2020 9:38:02 PM#
Const CMod$ = "MxVb_Str_SfxEns."

Function EnsAyDotSfx(Ay) As String()
EnsAyDotSfx = EnsAySfx(Ay, ".")
End Function

Function EnsAySfx(Ay, Sfx$) As String()
Dim I: For Each I In Itr(Ay)
    PushI EnsAySfx, EnsSfx(I, Sfx)
Next
End Function

Function EnsSfx(S, Sfx)
If HasSfx(S, Sfx) Then
    EnsSfx = S
Else
    EnsSfx = S & Sfx
End If
End Function

Function EnsSfxDot$(S)
EnsSfxDot = EnsSfx(S, ".")
End Function

Function EnsSfxSemi$(S)
EnsSfxSemi = EnsSfx(S, ";")
End Function
