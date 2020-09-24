Attribute VB_Name = "MxDao_Sql_QpTy"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Sql_QpTy."
Function SqlTyDyC$(Dy(), C&)
SqlTyDyC = SqlTyzAv(DcDy(Dy, C))
End Function
Function SqlTyzAv$(Av())
Dim O As VbVarType, V, T As VbVarType
For Each V In Av
    T = VarType(V)
    If T = vbString Then
        If Len(V) > 255 Then SqlTyzAv = "Memo": Exit Function
    End If
'    O = MaxVbt(O, T)
Next
End Function
Function SqlTyzVbt$(Dy As VbVarType)
Dim O$
Select Case Dy
Case vbEmpty:   O = "Text(255)"
Case vbBoolean: O = "YesNo"
Case vbByte:    O = "Byte"
Case vbInteger: O = "Short"
Case vbLong:    O = "Long"
Case vbDouble:  O = "Double"
Case vbSingle:  O = "Single"
Case vbCurrency: O = "Currency"
Case vbDate:    O = "Date"
Case vbString:  O = "Text(255)"
Case Else: Stop
End Select
SqlTyzVbt = O
End Function
