Attribute VB_Name = "MxIde_Src_MthBE"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Src_MthBE."
Function BeiMthn(Src$(), Mthn, Optional ShtMthTy$) As Bei
Dim B&: B = Mthix(Src, Mthn, ShtMthTy)
If B = -1 Then BeiMthn = BeiEmp: Exit Function
BeiMthn = Bei(B, Mtheix(Src, B))
End Function
Function BeiyMth(Src$()) As Bei()
Dim Ix: For Each Ix In ItrMthix(Src)
    PushBei BeiyMth, Bei(Ix, Mtheix(Src, Ix))
Next
End Function
Function BeiyMthn(Src$(), Mthn, Optional ShtMthTy$) As Bei()
Dim Ix: For Each Ix In Itr(MthixyMthn(Src, Mthn, ShtMthTy))
   PushBei BeiyMthn, Bei(Ix, Mtheix(Src, Ix))
Next
End Function
Function Mtheix&(Src$(), Bix)
If Bix < 0 Then Mtheix = -1: Exit Function
Mtheix = EixSrcItm(Src, Bix, MthkdL(Src(Bix)))
End Function
Function HasMtheix(Src$(), Bix) As Boolean
On Error GoTo X
Mtheix Src, Bix
HasMtheix = True
Exit Function
X:
End Function

Function Mtheno&(M As CodeModule, Mthlno&)
Dim MLn$: MLn = M.Lines(Mthlno, 1)
Dim ELn$: ELn = "End " & MthkdL(MLn)
Dim O&: For O = Mthlno + 1 To M.CountOfLines
    If HasPfx(M.Lines(O, 1), ELn) Then Mtheno = O: Exit Function
Next
End Function
