Attribute VB_Name = "MxVb_Fs_Ass"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_Ass."

Function PthAssEns$(Ffn)
Dim O$: O = PthAss(Ffn)
PthEns O
PthAssEns = O
End Function

Function PthAss$(Ffn): PthAss = Pth(Ffn) & "." & Fn(Ffn) & "\": End Function
