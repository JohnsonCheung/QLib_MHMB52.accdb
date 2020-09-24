Attribute VB_Name = "MxVb_Str_ExpandPfx"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Str_ExpandPfx."

Function SPfxx$(Pfx$, Ss$) '#SPfxx:S-Add-Pfx-To-Ss#
Dim O$()
Dim I: For Each I In SplitSpc(Ss)
    PushS O, Pfx & I
Next
SPfxx = Join(O, " ")
End Function
Function Psnosss$(Pfx$, Fst%, Las%, Optional Fmt$, Optional Sep$ = " ") '#Psnosss:Pfx-Seq-no-Ss#
Psnosss = Join(Psnoy(Pfx, Fst, Las, Fmt), Sep)
End Function
Function Psnoy(Pfx$, Fst%, Las%, Optional Fmt$) As String() '#Psnoy:Pfx-Seq-No-Array#
Dim I%: For I = Fst To Las
    PushS Psnoy, Pfx & Format(I, Fmt)
Next
End Function
