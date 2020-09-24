Attribute VB_Name = "MxIde_Md_Op_SrtMd_zIntl_DiMthl"
'#DiMth:Di-Mthl# Its Key is Mi2MdyNm.  It is used in SrtMd.  If it is Dcl, the Key is * *Dcl, which will be sorted a begining
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_MthDic."
Private Sub B_DiMthlMC():                                       VcDi DiMthlMC:                  End Sub
Private Sub B_DiMthlPC():                                       VcDi DiMthlPC:                  End Sub
Function DiMthlM(M As CodeModule) As Dictionary:  Set DiMthlM = DiMthlSrc(SrcM(M), MdnDotM(M)): End Function
Function DiMthlMC() As Dictionary:               Set DiMthlMC = DiMthlM(CMd):                   End Function
Function DiMthlP(P As VBProject) As Dictionary
Set DiMthlP = New Dictionary
Dim C As VBComponent: For Each C In P.VBComponents
    PushDi DiMthlP, DiMthlM(C.CodeModule)
Next
End Function
Function DiMthlPC() As Dictionary: Set DiMthlPC = DiMthlP(CPj): End Function

Function DiMthlSrc(Src$(), Optional Mdn$) As Dictionary
Dim P$: If Mdn <> "" Then P = Mdn & "."
Dim O As New Dictionary
With O
    .CompareMode = BinaryCompare
    .Add P & "*Dcl", DcllSrc(Src)
    Dim Ix: For Each Ix In ItrMthix(Src)
        .Add P & Mth3nLn(Src(Ix)), MthlIx(Src, Ix)
    Next
End With
Set DiMthlSrc = O
End Function
