Attribute VB_Name = "MxVb_Fs_Pth_PthR"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_Pth_PthR."
Private O$(), A_Spec$, A_Atr As VbFileAttribute

Function PthyEmpR(Pth) As String()
Dim I: For Each I In Itr(PthyR(Pth))
    If IsEmpPth(I) Then PushI PthyEmpR, I
Next
End Function

Function EntyR(Pth, Optional FilSpec$ = "*.*") As String()
Erase O
A_Spec = FilSpec
WSetVarO Pth
EntyR = O
Erase O
End Function
Private Sub WSetVarO(Pth)
PushI O, Pth
PushIAy O, Ffny(Pth, A_Spec)
Dim P$(): P = Pthy(Pth, A_Spec)
Dim I: For Each I In Itr(P)
    WSetVarO I
Next
End Sub

Private Sub B_FfnyR()
Dim Pth, Spec$, Atr As FileAttribute
GoSub T0
GoSub T1
Exit Sub
T0:
    Pth = "C:\Users\User\Documents\Projects\Vba"
    GoTo Tst
T1:
    Pth = "C:\Users\User\Documents\WindowsPowershell\"
    GoTo Tst
Tst:
    Act = FfnyR(Pth, Spec)
    Brw Act
    Stop
    Return
End Sub
Function FfnyR(Pth, Optional Spec$ = "*.*", Optional Atr As VbFileAttribute) As String()
Erase O
A_Spec = Spec
A_Atr = Atr
FfnyR1 Pth
FfnyR = O
End Function

Sub FfnyR1(Pth)
Const CSub$ = CMod & "FfnyR1"
PushIAy O, Ffny(Pth, A_Spec, A_Atr)
If Si(O) Mod 1000 = 0 Then Debug.Print CSub, "...Reading", "#Ffn-read", Si(O)
Dim P$(): P = Pthy(Pth)
If Si(P) = 0 Then Exit Sub
Dim I: For Each I In P
    FfnyR1 I
Next
End Sub

Private Sub B_EntyR()
Dim A$(): A = EntyR("C:\users\user\documents\")
Debug.Print Si(A)
Stop
DmpAy A
End Sub

Private Sub B_PthyEmpR()
Brw PthyEmpR(PthTmpRoot)
End Sub

Private Sub B_Enty()
BrwPth Enty(PthTmpRoot)
End Sub

Private Sub B_DltEmpPthR()
Z:
    DltEmpPthR PthTmpRoot
    Return
Z1:
    Debug.Print "Before-----"
    D PthyEmpR(PthTmpRoot)
    DltEmpPthR PthTmpRoot
    Debug.Print "After-----"
    D PthyEmpR(PthTmpRoot)
    Return
End Sub

Function PthyR(Pth, Optional Spec$ = "*.*", Optional Atr As VbFileAttribute) As String()
Erase O
A_Spec = Spec
A_Atr = Atr
W2SetVarO Pth
PthyR = O
Erase O
End Function
Private Sub W2SetVarO(Pth)
PushI O, Pth
Dim P$(): P = Pthy(Pth, A_Spec, A_Atr)
Dim I: For Each I In Itr(P)
    W2SetVarO I
Next
End Sub
