Attribute VB_Name = "MxIde_Dcl_DimDrs"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Dcl_DimDrs."
Public Const DimFf$ = "Varn Tyc Tyn IsAy"
Private Sub B_DrsTDimPC()
GoSub ZZ
Exit Sub
ZZ:  BrwDrs DrsTDimPC
    Return
End Sub
Function DrsTDimPC() As Drs:                  DrsTDimPC = DrsTDimP(CPj):                                                End Function
Function DrsTDimP(P As VBProject) As Drs:      DrsTDimP = DrsTDimItmy(SySrtQ(SywDis(ItmyDimStmty(StmtyDim(SrcP(P)))))): End Function
Function DrsTDimItmy(ItmyDim$()) As Drs:    DrsTDimItmy = DrsFf(DimFf, WDyItmy(ItmyDim)):                               End Function
Function DrsTDimStmty(StmtyDim$()) As Drs: DrsTDimStmty = DrsFf(DimFf, WDyItmy(ItmyDimStmty(StmtyDim))):                End Function

Private Function WDyItmy(ItmyDim$()) As Variant()
Dim ItmDim: For Each ItmDim In Itr(ItmyDim)
    PushI WDyItmy, WDr(ItmDim)
Next
End Function
Private Function WDr(ItmDim) As Variant()
Const CSub$ = CMod & "WDr"
Dim S$: S = ItmDim
Dim Varn$: Varn = ShfNm(S): If Varn = "" Then Thw CSub, "Given ItmDim does not have a name", "ItmDim", ItmDim
With TVtVsfx(S)
    WDr = Array(Varn, .Tyc, .Tyn, .IsAy)
End With
End Function
