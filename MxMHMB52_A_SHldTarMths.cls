VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MxMHMB52_A_SHldTarMths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private Type TarMthDs
    StmTMthDy() As Variant
    BusTMthDy() As Variant
    L1TMthDy() As Variant
    L2TMthDy() As Variant
    L3TMthDy() As Variant
    L4TMthDy() As Variant
    SkuTMthDy() As Variant
    Msg() As String
End Type
Sub LoadTarMth()
Const CSub$ = CMod & "TarMthLoad"
Dim A$: A = InputBox("Stocking Report file name:")
If A = "" Then Exit Sub
If Dir(A) = "" Then MsgBox "File not found", vbInformation: Exit Sub
Dim X As New Excel.Application: X.Visible = False
Dim Wb As Workbook: Set Wb = X.Workbooks.Open(A)
If Not IsShWb(Wb) Then
    Dim M$: M = MisShWsMsg(Wb)
    Wb.Close
    MsgBox M, vbCritical
    Wb.Close: Exit Sub
End If
Dim D As TarMthDs: D = TarMthDs(Wb): Wb.Close
ChkEry D.Msg, CSub

RplTarMth "Stm", D.StmTMthDy
RplTarMth "Bus", D.BusTMthDy
RplTarMth "L1", D.L1TMthDy
RplTarMth "L2", D.L2TMthDy
RplTarMth "L3", D.L3TMthDy
RplTarMth "L4", D.L4TMthDy
RplTarMth "Sku", D.L4TMthDy
End Sub
Private Function TarMthDs(Wb As Workbook) As TarMthDs

End Function

Private Sub RplTarMth(PhItm$, PHItmTMthDy())

End Sub
Private Function IsShWb(Wb As Workbook) As Boolean
Dim N$(): N = Wny(Wb)
Dim I: For Each I In MH.PH.Itmy7
    If Not HasEle(N, "StkHld " & I) Then Exit Function
Next
IsShWb = True
End Function
Private Function MisShWsMsg$(Wb As Workbook)
Dim A$(): A = Wny(Wb)
Dim E$(): Stop 'E = Split(ExpandPfxx("StkHld ", Ph7Ss), " ")
Dim M$(): M = SyMinus(E, A)
If Si(M) = 0 Then Exit Function
MisShWsMsg = "This StkHld Ws is missing:" & vbCrLf & Join(M, vbCrLf)
End Function
