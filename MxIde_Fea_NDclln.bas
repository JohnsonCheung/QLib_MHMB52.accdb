Attribute VB_Name = "MxIde_Fea_NDclln"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Fea_NDclln."
Private Sub B_NDclln()
'GoSub T1
'GoSub Z1
GoSub ZZ
Exit Sub
Dim Src$()
T1:
    Src = SrcMC
    Ept = 3
    GoTo Tst
Tst:
    Act = NDclln(Src)
    C
    Return
ZZ:
    Dim Ix%: Ix = 0
    Dim Cmp As VBComponent: For Each Cmp In CPj.VBComponents
        Dim N1%: N1 = Cmp.CodeModule.CountOfDeclarationLines
        Dim N2%: N2 = NDclln(SrcCmp(Cmp))
        If N1 <> N2 Then
            Ix = Ix + 1
            Debug.Print Ix; " "; AliR(N1, 4); AliR(N2, 4); " "; Cmp.Name
        End If
    Next
    Return
Z1:
    Dim M As CodeModule: Set M = CMd
    N1 = M.CountOfDeclarationLines
    N2 = NDclln(SrcM(M))
    If N1 <> N2 Then MsgBox N1 & " " & N2 & " " & M.Name
    Return
End Sub
Function NDclln%(Src$())
Dim MIxFst&: MIxFst = MthixFst(Src): If MIxFst = -1 Then NDclln = EixNB(Src) + 1: Exit Function
Dim O&: For O = MIxFst - 1 To 0 Step -1
    If LTrim(Src(O)) <> "" Then NDclln = O + 1: Exit Function ' Ix of fst non-blank line above fst mthln, return that ix + 1
Next
NDclln = 0 ' All are blank lines above fst Mthix, return 0
End Function
