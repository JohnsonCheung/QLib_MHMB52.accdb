Attribute VB_Name = "MxVb_Fs_Ffn_SubStr"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_Ffn_SubStr."
Function CutPth$(Ffn)
Dim P%: P = InStrRev(Ffn, SepPth)
If P = 0 Then CutPth = Ffn: Exit Function
CutPth = Mid(Ffn, P + 1)
End Function
Function FnFfn$(Ffn):   FnFfn = CutPth(Ffn):   End Function
Function Fn$(Ffn):         Fn = CutPth(Ffn):   End Function
Function Fnn$(Ffn):       Fnn = Ffnn(Fn(Ffn)): End Function
Function FnnFfn$(Ffn): FnnFfn = Ffnn(Fn(Ffn)): End Function
Function CutExt$(Ffn): CutExt = Ffnn(Ffn):     End Function
Function Ffnn$(Ffn)
Dim B$, C$, P%
B = Fn(Ffn)
P = InStrRev(B, ".")
If P = 0 Then
    C = B
Else
    C = Left(B, P - 1)
End If
Ffnn = Pth(Ffn) & C
End Function

Function ExtFfn$(Ffn): ExtFfn = Ext(Ffn): End Function
Function Ext$(Ffn)
Dim B$, P%
B = Fn(Ffn)
P = InStrRev(B, ".")
If P = 0 Then Exit Function
Ext = Mid(B, P)
End Function
Function PthFfn$(Ffn): PthFfn = Pth(Ffn): End Function
Function Pth$(Ffn)
Dim P%: P = InStrRev(Ffn, "\")
If P = 0 Then Exit Function
Pth = Left(Ffn, P)
End Function
