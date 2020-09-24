Attribute VB_Name = "MxIde_Pj_Prp_PjPrp"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Pj_Prp."

Function CPjf$(): CPjf = Pjf(CPj): End Function
Function Pjf$(P As VBProject)
On Error GoTo X
Pjf = P.FileName
Exit Function
X: Debug.Print FmtQQ("Cannot get Pjf for Pj(?). Err[?]", P.Name, Err.Description)
End Function
Function IsPjProt(P As VBProject) As Boolean
Const CSub$ = CMod & "IsPjProt"
If Not IsPjProt(P) Then Exit Function
Debug.Print CSub, FmtQQ("Skip protected Pj{?)", P.Name)
IsPjProt = True
End Function
Function IsPjFba(P As VBProject) As Boolean: IsPjFba = IsFba(Pjf(P)): End Function
Function MdFst(P As VBProject) As CodeModule
Dim Cmp As VBComponent
For Each Cmp In P.VBComponents
    If IsMd(CvCmp(Cmp)) Then
        Set MdFst = Cmp.CodeModule
        Exit Function
    End If
Next
End Function

Function ModFst(P As VBProject) As CodeModule
Dim Cmp As VBComponent
For Each Cmp In P.VBComponents
    If IsMod(Cmp) Then
        Set ModFst = Cmp.CodeModule
        Exit Function
    End If
Next
End Function

Function PthPC$():              PthPC = PthP(CPj):   End Function
Function PthP$(P As VBProject):  PthP = Pth(Pjf(P)): End Function '#Pj-Pth#
Function Pjfn$(P As VBProject):  Pjfn = Fn(Pjf(P)):  End Function

Function MdP(P As VBProject, Mdn) As CodeModule:     Set MdP = P.VBComponents(Mdn).CodeModule: End Function
Function CmpP(P As VBProject, Cmpn) As VBComponent: Set CmpP = P.VBComponents(Cmpn):           End Function
Function LinesPC$():                                 LinesPC = LinesFt(CPjf):                  End Function
Function Pjn$(P As VBProject):                           Pjn = P.Name:                         End Function
