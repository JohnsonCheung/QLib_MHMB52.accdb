Attribute VB_Name = "MxVb_Fs_S32Atr"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_S32Atr."

Type S32Atr
    Atr As String
    Val As String
End Type

Function S32AtrSI&(A() As S32Atr)
On Error Resume Next
S32AtrSI = UBound(A) + 1
End Function

Function S32AtrUB&(A() As S32Atr)
S32AtrUB = S32AtrSI(A) - 1
End Function

Sub PushS32Atr(O() As S32Atr, M As S32Atr)
Dim N&: N = S32AtrSI(O)
ReDim Preserve O(N)
O(N) = M
End Sub
