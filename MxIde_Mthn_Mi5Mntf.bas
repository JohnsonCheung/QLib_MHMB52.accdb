Attribute VB_Name = "MxIde_Mthn_Mi5Mntf"
Option Compare Text
Option Explicit
Private Sub Mi5yMntflModPC__Tst():                      Vc Mi5yMntflModPC:  End Sub
Function Mi5yMntflModPC() As String(): Mi5yMntflModPC = Mi5yMntflModP(CPj): End Function 'Mntfl = module-mthn-shtty-modifier
Function Mi5yMntflModP(P As VBProject) As String()  'Method-module-key-name-array
Dim C As VBComponent: For Each C In P.VBComponents
    If IsMod(C) Then
        PushIAy Mi5yMntflModP, Mi5yMntflM(C.CodeModule)
    End If
Next
End Function
Private Sub Mi5yMntflPC__Tst():                   Vc Mi5yMntflPC:  End Sub
Function Mi5yMntflPC() As String(): Mi5yMntflPC = Mi5yMntflP(CPj): End Function 'Mntfl = module-mthn-shtty-modifier
Function Mi5yMntflP(P As VBProject) As String()  'Method-module-key-name-array
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy Mi5yMntflP, Mi5yMntflM(C.CodeModule)
Next
End Function
Function Mi5yMntflM(M As CodeModule) As String(): Mi5yMntflM = Mi5yMntflS(SrcM(M), Mdn(M)): End Function
Function Mi5yMntflS(Src$(), Mdn$) As String()
Dim L: For Each L In Itr(Mthlny(Src))
    PushI Mi5yMntflS, Mi5MntflL(L, Mdn)
Next
End Function
Function Mi5MntflL$(Mthln, Mdn$): Mi5MntflL = TslMi4Mntf(Mthln, Mdn) & " " & Mthln: End Function
