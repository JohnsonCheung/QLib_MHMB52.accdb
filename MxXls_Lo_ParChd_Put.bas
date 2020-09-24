Attribute VB_Name = "MxXls_Lo_ParChd_Put"
':Pcat: :Cml #Par-chd-At# ! It is a feature to allow a Drs to be shown in a ws
Option Compare Text
Option Explicit
Const CMod$ = "MxXls_Lo_ParChd_Put."
Private A_LonyPar$()
Private A_RnoyPar&()
Private A_IxParCur%
Sub PutLoParChd(Target As Range)
'Each non-associative Lo of name Xxx have 2 associative Lo of name, Xxx_Par and Xxx_Chd in Same ws.
'Functionals:
'   When at Hdr line, open the Par-Lo, locate cell2, highlight cell2
'   When at data non cell1, change the cell1 back color
'   When at data cell1,     hide the rows, filter the Xxx_Chd
Const CSub$ = CMod & "PutLoParChd"
Dim L As ListObject: Set L = LoCellIn(Target): If IsNothing(L) Then Exit Sub
If Not HasSfx(L.Name, "_Par") Then Exit Sub
WSetA L
Select Case True
Case HasCell(L.HeaderRowRange, Target): WAtHdr Target
Case HasCell(L.DataBodyRange, Target)
    If Target.Row = L.ListColumns(1).Range.Row Then
        WAtCell1 Target
    Else
        WAtNonCell1 Target
    End If
Case Else: ThwLgc CSub, "Current Cell in in Non-Empty-Lo, but not in Hdr nor DataRange", "Cell-Adr Lo-Range-Addr", Target.Address, L.Range.Address
End Select
End Sub
Private Sub WSetA(L As ListObject)

End Sub
Private Sub WAtHdr(R As Range)

End Sub
Private Sub WAtCell1(R As Range)

End Sub
Private Sub WAtNonCell1(R As Range)

End Sub
