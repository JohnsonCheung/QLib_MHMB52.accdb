Attribute VB_Name = "MxVb_Fun_Ty"
Option Compare Text
Option Explicit
Const CMod$ = "MxVb_Fun_Ty."

Function Vbt(TypeName) As VbVarType
Dim O As VbVarType
Select Case RmvSfx(TypeName, "()")
Case "Integer":  O = vbInteger
Case "Boolean":  O = vbBoolean
Case "Byte":     O = vbByte
Case "Currency": O = vbCurrency
Case "Date":     O = vbDate
Case "Decimal": O = vbDecimal
Case "Double":  O = vbDouble
Case "Empty":   O = vbEmpty
Case "Error":   O = vbError
Case "Integer": O = vbInteger
Case "Long":    O = vbLong
Case "Null":    O = vbNull
Case "Object":  O = vbObject
Case "Single":  O = vbSingle
Case "String":  O = vbString
Case "Variant": O = vbVariant
End Select
If HasSfx(TypeName, "()") Then O = O + vbArray
Vbt = O
End Function

Function TynyVy(Vy) As String()
Dim V: For Each V In Itr(Vy)
    PushI TynyVy, TypeName(V)
Next
End Function

Function VbtyVy(Vy) As VbVarType()
Dim V: For Each V In Itr(Vy)
    PushI VbtyVy, VarType(V)
Next
End Function

Function VbtyTyny(Vbtny$()) As VbVarType()
Dim T: For Each T In Itr(Vbtny)
    PushI VbtyTyny, Vbt(T)
Next
End Function

Function ValS(S$, T As VbVarType)
Dim O
Select Case T
Case VbVarType.vbBoolean:   O = CBool(S)
Case VbVarType.vbByte:      O = CByte(S)
Case VbVarType.vbCurrency:  O = CCur(S)
Case VbVarType.vbDate:      O = CDate(S)
Case VbVarType.vbDecimal:   O = CDec(S)
Case VbVarType.vbDouble:    O = CDbl(S)
Case VbVarType.vbEmpty:     O = Empty
Case VbVarType.vbError:     O = CVErr(Val(S))
Case VbVarType.vbInteger:   O = CInt(S)
Case VbVarType.vbLong:      O = CLng(S)
Case VbVarType.vbNull:      O = Null
Case VbVarType.vbSingle:    O = CSng(S)
Case VbVarType.vbString:    O = S
Case VbVarType.vbVariant:   O = S
Case Else:
End Select
End Function
