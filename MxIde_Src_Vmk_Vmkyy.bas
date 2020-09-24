Attribute VB_Name = "MxIde_Src_Vmk_Vmkyy"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_Vmk_Vmkyy."

Function VmklsyPC() As String():                        VmklsyPC = Vmklsy(SrcPC):            End Function
Function VmkyyPC() As Variant():                         VmkyyPC = Vmkyy(SrcPC):             End Function
Function Vmklsy(Src$(), Optional Fmix = 0) As String():   Vmklsy = LsyLyy(Vmkyy(Src, Fmix)): End Function
Function Vmkyy(Src$(), Optional Fmix = 0) As Variant()
Dim U&: U = UB(Src)
Dim Fm&: Fm = Fmix
Do
    If Fm > U Then Exit Function
    Dim Vmky$(): Vmky = VmkyFm(Src, Fm)
    Dim NLn%: NLn = Si(Vmky)
    If NLn = 0 Then Exit Function
    PushI Vmkyy, Vmky
    Fm = Fm + NLn
    DoEvents
Loop
End Function
Private Sub B_LsyVmk(): VcLsy LsyVmk(SrcPC): End Sub
Function LsyVmk(Src$()) As String()
Dim O$(), InBlk As Boolean, IsRmk As Boolean
Dim L: For Each L In Itr(Src)
    IsRmk = IsLnVmk(L)
    Select Case True
    Case InBlk And IsRmk
        PushI O, L
    Case InBlk
        If Not IsEmpAy(O) Then PushI LsyVmk, JnCrLf(O)
        InBlk = False
    Case IsRmk
        If Not IsEmpAy(O) Then PushI LsyVmk, JnCrLf(O)
        InBlk = True
        O = Sy(L)
    End Select
Next
If Not IsEmpAy(O) Then PushI LsyVmk, JnCrLf(O)
End Function
