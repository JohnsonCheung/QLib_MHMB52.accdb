Attribute VB_Name = "MxVb_Ay_Op_AgrAy"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Ay_Agr."

Sub BrwAgrMdLnCntPC():                                        BrwDi DiAgrMdLnCntPC: End Sub
Function DiAgrMdLnCntPC() As Dictionary: Set DiAgrMdLnCntPC = DiAgr(CntyMdLnPC):    End Function

Function CntyMdLnPC() As Long(): CntyMdLnPC = CntyMdLnP(CPj): End Function

Function CntyMdLnP(P As VBProject) As Long()
Dim C As VBComponent: For Each C In P.VBComponents
    PushI CntyMdLnP, C.CodeModule.CountOfLines
Next
End Function

Function CntNo0&(Nbry)
Dim O&
Dim V: For Each V In Itr(Nbry)
    If V <> 0 Then O = O + 1
Next
CntNo0 = O
End Function

Function DiAgr(Nbry) As Dictionary
'Ret : Agr Val ! where *Arg has Cnt Avg Max Min Sum
Dim O As New Dictionary
Dim Sum#: Sum = AySum(Nbry)
Dim NNo0&: NNo0 = CntNo0(Nbry)
Dim N&: N = Si(Nbry)
Dim AvgAll#, AvgNo0#
If N <> 0 Then AvgAll = Sum / N
If NNo0 <> 0 Then AvgNo0 = Sum / NNo0

O.Add "CntNo0", NNo0
O.Add "CntAll", N
O.Add "AvgNo0", AvgNo0
O.Add "AvgAll", AvgAll
O.Add "Sum", Sum
O.Add "Max", EleMax(Nbry)
O.Add "Min", EleMin(Nbry)
O.Add "MinGT0", MinEleGT0(Nbry)
Set DiAgr = O
End Function
