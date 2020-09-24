Attribute VB_Name = "MxVb_Fs_Pth_PthOp"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fs_Pth_Op."
Function SiPth&(Pth): SiPth = SiFfny(Ffny(Pth)): End Function
Sub VcPth(Pth)
If NoPth(Pth) Then Exit Sub
Shell FmtQQ("Code.cmd ""?""", Pth), vbMaximizedFocus
'ShellHid FmtQQ("""?"" --add ""?""", vc_cmdFfn, Pth)
End Sub
Sub BrwPth(Pth)
If NoPth(Pth) Then Exit Sub
ShellMax FmtQQ("Explorer ""?""", Pth)
End Sub

Sub DltPth(Pth): RmDir Pth: End Sub
Sub DltPthIf(Pth)
If HasPth(Pth) Then DltPth Pth
End Sub

Private Sub B_DltAllPthFil()
DltAllPthFil PthTmpRoot
End Sub

Sub DltAllPthFil(Pth)
If NoPth(Pth) Then Exit Sub
Dim F
For Each F In Itr(Ffny(Pth))
   DltFfn F
Next
End Sub

Sub DltEmpPthR(Pth)
Dim Ay$(), I, J%
Lp:
    J = J + 1: If J > 10000 Then Stop
    Dim Dlt As Boolean: Dlt = False
    For Each I In Itr(PthyEmpR(Pth))
        DltPthSilent I
        Dlt = True
    Next
    If Dlt Then GoTo Lp
End Sub
Sub DltPthSilent(Pth)
On Error Resume Next
RmDir Pth
End Sub

Sub DltAllEmpFdr(Pth)
Dim S: For Each S In Itr(Pthy(Pth))
   DltPthIfEmp S
Next
End Sub

Sub DltPthIfEmp(Pth)
If IsEmpPth(Pth) Then DltPth Pth
End Sub

Sub RenPthAddFdrPfx(Pth, Pfx)
RenPth Pth, PthAddFdrPfx(Pth, Pfx)
End Sub

Sub RenPth(Pth, NewPth)
Fso.GetFolder(Pth).Name = NewPth
End Sub

Private Sub B_DltEmpSubDir()
DltAllEmpFdr PthTmp
End Sub
