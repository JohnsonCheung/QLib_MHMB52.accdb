Attribute VB_Name = "MxDao_Ty_DaoTy"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Ty."
Public Const SsShtTyDao$ = "Att Boo Byt Chr Cur Dbl Dte Dec Int Lng Mem Tim Txt"
Function CvDaoty(A) As Dao.DataTypeEnum
CvDaoty = A
End Function

Function TyDao(StrTyDao$) As Dao.DataTypeEnum
Const CSub$ = CMod & "TyDao"
Dim O
Select Case StrTyDao
Case "Attachment": O = Dao.DataTypeEnum.dbAttachment
Case "Boolean":    O = Dao.DataTypeEnum.dbBoolean
Case "Byte":       O = Dao.DataTypeEnum.dbByte
Case "Currency":   O = Dao.DataTypeEnum.dbCurrency
Case "Date":       O = Dao.DataTypeEnum.dbDate
Case "Decimal":    O = Dao.DataTypeEnum.dbDecimal
Case "Double":     O = Dao.DataTypeEnum.dbDouble
Case "Integer":    O = Dao.DataTypeEnum.dbInteger
Case "Long":       O = Dao.DataTypeEnum.dbLong
Case "Memo":       O = Dao.DataTypeEnum.dbMemo
Case "Single":     O = Dao.DataTypeEnum.dbSingle
Case "Text":       O = Dao.DataTypeEnum.dbText
Case Else: Thw CSub, "Invalid ShtDaoty", "ShtDaoty Valid", StrTyDao, _
    SySs("Attachment Boolean Byte Currency Date Decimal Double Integer Long Memo Signle Text")
End Select
TyDao = O
End Function

Function DaotyShtTy(ShtTyDao$) As Dao.DataTypeEnum
Const CSub$ = CMod & "DaotyShtTy"
Dim O As Dao.DataTypeEnum
Select Case ShtTyDao
Case "Att":  O = dbAttachment
Case "Bool": O = dbBoolean
Case "Byt": O = dbByte
Case "Cur": O = dbCurrency
Case "Chr": O = dbChar
Case "Dte": O = dbDate
Case "Dec": O = dbDecimal
Case "Dbl": O = dbDouble
Case "Int": O = dbInteger
Case "Lng": O = dbLong
Case "Mem": O = dbMemo
Case "Sng": O = dbSingle
Case "Txt": O = dbText
Case "Tim": O = dbTime
Case Else: Thw CSub, "Invalid ShtTy", "The-Invalid-ShtTy Valid-ShtTy", ShtTyDao, SsShtTyDao
End Select
DaotyShtTy = O
End Function

Function Daoty(V) As Dao.DataTypeEnum
Dim T As VbVarType: T = VarType(V)
If T = vbString Then
    If Len(V) > 255 Then
        Daoty = dbMemo
    Else
        Daoty = dbText
    End If
    Exit Function
End If
Daoty = DaotyVb(T)
End Function

Function DaotyVb(A As VbVarType) As Dao.DataTypeEnum
Const CSub$ = CMod & "DaotyVb"
Dim O As Dao.DataTypeEnum
Select Case A
Case vbBoolean: O = dbBoolean
Case vbByte: O = dbByte
Case VbVarType.vbCurrency: O = dbCurrency
Case VbVarType.vbDate: O = dbDate
Case VbVarType.vbDecimal: O = dbDecimal
Case VbVarType.vbDouble: O = dbDouble
Case VbVarType.vbInteger: O = dbInteger
Case VbVarType.vbLong: O = dbLong
Case VbVarType.vbSingle: O = dbSingle
Case VbVarType.vbString: O = dbText
Case Else: Thw CSub, "Vbt cannot convert to Daoty", "Vbt", A
End Select
DaotyVb = O
End Function

Function DicntRs(A As Dao.Recordset, Optional Fld = 0) As Dictionary: Set DicntRs = DiCnt(AvRsF(A)): End Function

Function DtaTy$(T As Dao.DataTypeEnum)
Dim O$
Select Case T
Case dbAttachment: O = "Attachment"
Case Dao.DataTypeEnum.dbBoolean:    O = "Boolean"
Case Dao.DataTypeEnum.dbByte:       O = "Byte"
Case Dao.DataTypeEnum.dbCurrency:   O = "Currency"
Case Dao.DataTypeEnum.dbDate:       O = "Date"
Case Dao.DataTypeEnum.dbDecimal:    O = "Decimal"
Case Dao.DataTypeEnum.dbDouble:     O = "Double"
Case Dao.DataTypeEnum.dbInteger:    O = "Integer"
Case Dao.DataTypeEnum.dbLong:       O = "Long"
Case Dao.DataTypeEnum.dbMemo:       O = "Memo"
Case Dao.DataTypeEnum.dbSingle:     O = "Single"
Case Dao.DataTypeEnum.dbText:       O = "Text"
Case Dao.DataTypeEnum.dbChar:       O = "Char"
Case Dao.DataTypeEnum.dbTime:       O = "Time"
Case Dao.DataTypeEnum.dbLongBinary: O = "LongBinary"
Case Else: Stop
End Select
DtaTy = O
End Function

Function DtaTyy() As String(): DtaTyy = DtaTyySht(ShtTyDaoy): End Function
Function DtaTyySht(ShtTyDaoy$()) As String()
Dim S: For Each S In Itr(ShtTyDaoy)
    PushS DtaTyySht, DtaTySht(S)
Next
End Function
Function DtaTySht$(ShtTyDao)
Const CSub$ = CMod & "DtaTySht"
Select Case ShtTyDao
Case ""
Case ""
Case ""
Case Else: Thw CSub, "Invalid ShtTyDao", "ShtTyDao SsShtTyDao", ShtTyDao, SsShtTyDao
End Select
End Function
Function DtaTyFld$(D As Database, T, F$)
DtaTyFld = DtaTy(Fd(D, T, F).Type)
End Function

Function IsShtTyDao(S) As Boolean
Select Case Len(S)
Case 1, 3
    If Not IsAscUCas(Asc(S)) Then Exit Function
    IsShtTyDao = HasSsub(SsShtTyDao, " " & S & " ")
End Select
End Function

Function JnStrDicRsKeyJn(A As Dao.Recordset, KeyFld, JnStrFld, Optional Sep$ = " ") As Dictionary
Dim O As New Dictionary
Dim K, V$
While Not A.EOF
    K = A.Fields(KeyFld).Value
    V = Nz(A.Fields(JnStrFld).Value, "")
    If O.Exists(K) Then
        O(K) = O(K) & Sep & V
    Else
        O.Add K, CStr(Nz(V))
    End If
    A.MoveNext
Wend
Set JnStrDicRsKeyJn = O
End Function

Function ShtAdoTyAy(A() As ADODB.DataTypeEnum) As String()
Dim I
For Each I In Itr(A)
    PushI ShtAdoTyAy, ShtAdoTy(CLng(I))
Next
End Function

Function ShtTyDaoy() As String()
ShtTyDaoy = SySs(SsShtTyDao)
End Function

Function LyShtTyDao() As String()
Dim S: For Each S In ShtTyDaoy
    PushI LyShtTyDao, S & " " & DtaTySht(S)
Next
End Function

Function ShtTyLisDaotyAy$(A() As DataTypeEnum)
Dim O$, I
For Each I In A
    O = O & ShtDaoty(CvDaoty(I))
Next
ShtTyLisDaotyAy = O
End Function

Function ShtAdoTy$(A As ADODB.DataTypeEnum)
Dim O$
Select Case A
Case ADODB.DataTypeEnum.adTinyInt: O = "Byt"
Case ADODB.DataTypeEnum.adInteger: O = "Lng"
Case ADODB.DataTypeEnum.adSmallInt: O = "Int"
Case ADODB.DataTypeEnum.adDate: O = "Dte"
Case ADODB.DataTypeEnum.adVarChar: O = "Txt"
Case ADODB.DataTypeEnum.adBoolean: O = "Yes"
Case ADODB.DataTypeEnum.adDouble: O = "Dbl"
Case ADODB.DataTypeEnum.adCurrency: O = "Cur"
Case ADODB.DataTypeEnum.adSingle: O = "Sng"
Case ADODB.DataTypeEnum.adDecimal: O = "Dec"
Case ADODB.DataTypeEnum.adVarWChar: O = "Mem"
Case Else: O = "?" & A & "?"
End Select
ShtAdoTy = O
End Function
Function StrAdoTy$(A As ADODB.DataTypeEnum)
Const CSub$ = CMod & "StrAdoTy"
Dim O$
Select Case A
Case ADODB.DataTypeEnum.adTinyInt:  O = "TinyInt"
Case ADODB.DataTypeEnum.adCurrency: O = "Currency"
Case ADODB.DataTypeEnum.adDecimal:  O = "Decimal"
Case ADODB.DataTypeEnum.adDouble:   O = "Double"
Case ADODB.DataTypeEnum.adSmallInt: O = "SmallInt"
Case ADODB.DataTypeEnum.adInteger:  O = "Integer"
Case ADODB.DataTypeEnum.adSingle:   O = "Single"
Case ADODB.DataTypeEnum.adChar:     O = "Char"
Case ADODB.DataTypeEnum.adGUID:     O = "GUID"
Case ADODB.DataTypeEnum.adVarChar:  O = "VarChar"
Case ADODB.DataTypeEnum.adVarWChar: O = "VarWChar"
Case ADODB.DataTypeEnum.adLongVarChar: O = "LongVarChar"
Case ADODB.DataTypeEnum.adBoolean:  O = "Boolean"
Case ADODB.DataTypeEnum.adDate:     O = "Date"
Case Else
   Thw CSub, "Not supported Case ADODB type", "AdoTy", A
End Select
StrAdoTy = O
End Function

Function StrTyDao$(A As Dao.DataTypeEnum)
Const CSub$ = CMod & "StrTyDao"
Dim O$
Select Case A
Case Dao.DataTypeEnum.dbAttachment: O = "Attachment"
Case Dao.DataTypeEnum.dbBoolean:    O = "Boolean"
Case Dao.DataTypeEnum.dbByte:       O = "Byte"
Case Dao.DataTypeEnum.dbCurrency:   O = "Currency"
Case Dao.DataTypeEnum.dbChar:       O = "Char"
Case Dao.DataTypeEnum.dbDate:       O = "Date"
Case Dao.DataTypeEnum.dbDecimal:    O = "Decimal"
Case Dao.DataTypeEnum.dbDouble:     O = "Double"
Case Dao.DataTypeEnum.dbInteger:    O = "Integer"
Case Dao.DataTypeEnum.dbLong:       O = "Long"
Case Dao.DataTypeEnum.dbLongBinary: O = "LongBinary"
Case Dao.DataTypeEnum.dbMemo:       O = "Memo"
Case Dao.DataTypeEnum.dbSingle:     O = "Single"
Case Dao.DataTypeEnum.dbText:       O = "Text"
Case Dao.DataTypeEnum.dbTime:       O = "Time"
Case Dao.DataTypeEnum.dbTimeStamp:  O = "TimeStamp"
Case Else: Thw CSub, "Unsupported Daoty, cannot covert to ShtTy", "Daoty", A
End Select
StrTyDao = O
End Function

Function ShtDaoty$(A As Dao.DataTypeEnum)
Const CSub$ = CMod & "ShtDaoty"
Dim O$
Select Case A
Case Dao.DataTypeEnum.dbAttachment: O = "Att"
Case Dao.DataTypeEnum.dbBoolean:    O = "Bln"
Case Dao.DataTypeEnum.dbByte:       O = "Byt"
Case Dao.DataTypeEnum.dbCurrency:   O = "Cur"
Case Dao.DataTypeEnum.dbChar:       O = "Chr"
Case Dao.DataTypeEnum.dbDate:       O = "Dte"
Case Dao.DataTypeEnum.dbDecimal:    O = "Dec"
Case Dao.DataTypeEnum.dbDouble:     O = "Dbl"
Case Dao.DataTypeEnum.dbInteger:    O = "Int"
Case Dao.DataTypeEnum.dbLong:       O = "Lgn"
Case Dao.DataTypeEnum.dbMemo:       O = "Mem"
Case Dao.DataTypeEnum.dbSingle:     O = "Sgn"
Case Dao.DataTypeEnum.dbText:       O = "Txt"
Case Dao.DataTypeEnum.dbTime:       O = "Tim"
Case Else: Thw CSub, "Unsupported Daoty, cannot covert to ShtTy", "Daoty", A
End Select
ShtDaoty = O
End Function
