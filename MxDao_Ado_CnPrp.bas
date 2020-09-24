Attribute VB_Name = "MxDao_Ado_CnPrp"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Ado_CnPrp."
Private Sub B_TnyOupCn()
Dim T$(): T = TnyOupCn(MHO.MHODuty.CnDta)
Stop
End Sub
Function TnyOupCn(C As ADODB.Connection) As String(): TnyOupCn = AwPfx(TnyCn(C), "@"): End Function
Function TnyCn(C As ADODB.Connection) As String()
Dim R As ADODB.Recordset: Set R = C.OpenSchema(adSchemaTables)
'BrwDrs DrsArs(R)
With R
    While Not .EOF
        If !TABLE_TYPE = "TABLE" Then
            PushI TnyCn, !TABLE_NAME
        End If
        .MoveNext
    Wend
End With
End Function

Function DcStrArs(R As ADODB.Recordset, Optional F = 0) As String():  DcStrArs = DcIntoArs(SyEmp, R, F):   End Function ' return string column
Function DcArs(R As ADODB.Recordset, Optional F = 0) As Variant():       DcArs = DcIntoArs(AvEmp, R, F):   End Function ' return column
Function DcIntArs(R As ADODB.Recordset, Optional F = 0) As Integer(): DcIntArs = DcIntoArs(IntyEmp, R, F): End Function ' return integer column
Function DcIntoArs(IntoyAy, A As ADODB.Recordset, Optional F = 0)
Dim O: O = IntoyAy: Erase O
With A
    While Not .EOF
        PushI O, Nz(.Fields(F).Value, Empty)
        .MoveNext
    Wend
    .Close
End With
DcIntoArs = O
End Function

Function Ars(Cn As ADODB.Connection, Q) As ADODB.Recordset:       Set Ars = ArsCnq(Cn, Q): End Function
Function ArsCnq(Cn As ADODB.Connection, Q) As ADODB.Recordset: Set ArsCnq = Cn.Execute(Q): End Function
