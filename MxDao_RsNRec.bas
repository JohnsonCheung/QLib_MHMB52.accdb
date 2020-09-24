Attribute VB_Name = "MxDao_RsNRec"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_RsNRec."
Function NoRec(R As Dao.Recordset) As Boolean
Select Case True
Case Not R.EOF, Not R.BOF
Case Else:: NoRec = True
End Select
End Function
Function NRecRs&(R As Dao.Recordset): NRecRs = NRec(R): End Function
Function NRec&(R As Dao.Recordset)
If NoRec(R) Then Exit Function
R.MoveLast
NRec = R.RecordCount
R.MoveFirst
End Function
