Attribute VB_Name = "MxDao_Ado_Dta"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Ado_Dta."
Function DrsCnq(Cn As ADODB.Connection, Q) As Drs:    DrsCnq = DrsArs(ArsCnq(Cn, Q)):    End Function
Function DrsFbqAdo(Fb, Q) As Drs:                  DrsFbqAdo = DrsArs(ArsFbq(Fb, Q)):    End Function
Function DrsArs(R As ADODB.Recordset) As Drs:         DrsArs = Drs(FnyArs(R), DyArs(R)): End Function

Private Sub B_DyArs()
GoSub ZZ
Exit Sub
Dim Cn As ADODB.Connection
Dim Q$
'S = "GRANT SELECT ON MSysObjects TO Admin;"
'CurrentProject.Connection.Execute S
ZZ:
    Set Cn = MHO.MHODuty.CnDta
    Q = "Select * From KE24"
    GoTo Tst
Tst:
    BrwDy DyArs(ArsCnq(Cn, Q))
    Return
End Sub
Function DyArs(R As ADODB.Recordset) As Variant()
While Not R.EOF
    PushI DyArs, W3Dr(R.Fields)
    R.MoveNext
Wend
End Function
Private Function W3Dr(R As ADODB.Fields, Optional N%) As Variant()
Dim F As ADODB.Field
For Each F In R
   PushI W3Dr, Nz(F.Value)
Next
End Function
