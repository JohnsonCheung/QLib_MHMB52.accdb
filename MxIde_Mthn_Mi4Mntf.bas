Attribute VB_Name = "MxIde_Mthn_Mi4Mntf"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Mthn_Mi4Mntf."
Private Sub DyMi4MntfModP__Tst(): Vc TsyDy(DyMi4MntfModP(CPj)): End Sub
Function DyMi4MntfModP(P As VBProject) As Variant()  'Method-module-key-name-array
Dim C As VBComponent: For Each C In P.VBComponents
    If IsMod(C) Then
        PushIAy DyMi4MntfModP, DyMi4MntfM(C.CodeModule)
    End If
Next
End Function
Function Mit4yMntfP(P As VBProject) As String(): Mit4yMntfP = TsyDy(DyMi4MntfP(P)): End Function
Function DyMi4MntfP(P As VBProject) As Variant()  'Method-module-key-name-array
Dim C As VBComponent: For Each C In P.VBComponents
    PushIAy DyMi4MntfP, DyMi4MntfM(C.CodeModule)
Next
End Function
Private Function DyMi4MntfM(M As CodeModule) As Variant(): DyMi4MntfM = DyMi4MntfS(SrcM(M), Mdn(M)): End Function
Private Function DyMi4MntfS(Src$(), Mdn$) As Variant()
Dim L: For Each L In Itr(Src)
    PushSomSi DyMi4MntfS, DrMi4MntfL(L, Mdn)
Next
End Function
Function TslMi4Mntf$(Ln, Mdn$): TslMi4Mntf = JnTab(DrMi4MntfL(Ln, Mdn)): End Function
Function SsMi4Mntf$(Ln, Mdn$):   SsMi4Mntf = JnSpc(DrMi4MntfL(Ln, Mdn)): End Function
Private Function DrMi4MntfL(Ln, Mdn$) As String()
Dim M As TMth: M = TMthL(Ln): If M.Mthn = "" Then Exit Function
DrMi4MntfL = DrMi4MntfzTMth(TMthmdn(Mdn, M))
End Function

Private Function DrMi4MntfzTMth(N As TMthmdn) As String()
With N
    Dim O$(): ReDim O(3)
    O(0) = .Mdn
    With .TMth
    O(1) = .Mthn
    O(2) = .ShtTy
    O(3) = StrDft(.ShtMdy, "Pub")
    End With
    DrMi4MntfzTMth = O
End With
End Function
