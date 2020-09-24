Attribute VB_Name = "MxIde_Mth_CSub_zIntl_EnsCSubSrc"
Option Compare Database
Option Explicit
Private Sub B_SrcoptEnsCSub()
GoSub ZZ1
Stop
Exit Sub
Exit Sub
Dim M As CodeModule, Act As Lyopt, Cnt%

ZZ1:
    Cnt = 0
    InfNCmp CSub
    Dim C As VBComponent: For Each C In CPj.VBComponents
        'InfCnt Cnt
        DoEvents
        If HasPfx(C.Name, "MxIde_Mth_Slm_zIntl_AliSlmSrc") Then
            Debug.Print C.Name
            SrcoptEnsCSub C.CodeModule
        End If
    Next
    Exit Sub
    Return '<-- it will crash', Exit Sub it is used
End Sub
Function SrcoptEnsCSub(M As CodeModule) As Lyopt
Dim S$(): S = SrcM(M): If Si(S) = 0 Then Exit Function
SrcoptEnsCSub = LyoptOldNew(S, SrcEnsCSub(S))
End Function
Private Function SrcEnsCSub(Src$()) As String()
Dim O$(): O = Src
Dim Mthix: For Each Mthix In Itr(AyRev(Mthixy(O)))
    Dim Eix%: Eix = Mtheix(O, Mthix)
    Dim Mthy$(): Mthy = AwBE(O, Mthix, Eix)
    With CSubboptEns(Mthy)
        If .Som Then
            O = AyRplBE(O, .Ly, Mthix, Eix)
        End If
    End With
Next
SrcEnsCSub = O
End Function
