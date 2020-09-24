Attribute VB_Name = "MxDao_Rs_NoAtt"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Rs_NoAtt."

Function CmaFldNoAtt$(D As Database, T) 'Return '' if no att fields else return CmaFld after rmv the att fld
Dim F$(): F = WFnyAtt(D, T): If Si(F) = 0 Then CmaFldNoAtt = "*": Exit Function
CmaFldNoAtt = JnCma(AmQuoSq(SyMinus(Fny(D, T), F)))
End Function
Private Function WFnyAtt(D As Database, T) As String()
Dim F As Dao.Field: For Each F In Td(D, T).Fields
    If F.Type = Dao.DataTypeEnum.dbAttachment Then PushI WFnyAtt, F.Name
Next
End Function

Function RsTblNoAtt(D As Database, T) As Dao.Recordset
Dim F$: F = CmaFldNoAtt(D, T)
Set RsTblNoAtt = Rs(D, SqlSelX(T, F))
End Function
