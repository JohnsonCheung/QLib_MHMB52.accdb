Attribute VB_Name = "MxDao_Def_Fdzz"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Def_Fdzz."
Public Const EleLblss$ = "*Fld *Ty ?Req ?AlZZLen Dft VTxt VRul TxtSz Epr"

Function FdNwNN(F, Optional Ty As Dao.DataTypeEnum = dbText, Optional TxtSi As Byte = 255, Optional ZLen As Boolean, Optional Epr$, Optional Dft$, Optional VRul$, Optional VTxt$) As Dao.Field2
Set FdNwNN = FdNw(F, Ty, True, TxtSi, ZLen, Epr, Dft, VRul, VTxt)
End Function
Function FdNw(F, Optional Ty As Dao.DataTypeEnum = dbText, Optional Req As Boolean, Optional TxtSi As Byte = 255, Optional ZLen As Boolean, Optional Epr$, Optional Dft$, Optional VRul$, Optional VTxt$) As Dao.Field2
Dim O As New Dao.Field
With O
    .Name = F
    .Required = Req
    .Type = Ty
    If Ty = dbText Then
        .Size = TxtSi
        .AllowZeroLength = ZLen
    End If
    If Epr <> "" Then
        CvFd2(O).Epression = Epr
    End If
    If Dft <> "" Then O.DefaultValue = Dft
End With
Set FdNw = O
End Function

Function FdAtt(F) As Dao.Field2:   Set FdAtt = FdNw(F, dbAttachment):    End Function
Function FdBool(F) As Dao.Field2: Set FdBool = FdNw(F, dbBoolean, True): End Function
Function FdByt(F) As Dao.Field2:   Set FdByt = FdNw(F, dbByte):          End Function
Function FdChr(F) As Dao.Field2:   Set FdChr = FdNw(F, dbChar):          End Function
Function FdCur(F) As Dao.Field2:   Set FdCur = FdNw(F, dbCurrency):      End Function
Function FdDbl(F) As Dao.Field2:   Set FdDbl = FdNw(F, dbDouble):        End Function
Function FdDec(F) As Dao.Field2:   Set FdDec = FdNw(F, dbDecimal):       End Function
Function FdDte(F) As Dao.Field2:   Set FdDte = FdNw(F, dbDate):          End Function

Function FdNNByt(F, Optional Dft$) As Dao.Field2: Set FdNNByt = FdNwNN(F, dbByte, Dft:=Dft):   End Function
Function FdNNDte(F, Optional Dft$) As Dao.Field2: Set FdNNDte = FdNwNN(F, dbDate, Dft:=Dft):   End Function
Function FdNNLng(F, Optional Dft$) As Dao.Field2: Set FdNNLng = FdNwNN(F, dbLong, Dft:=Dft):   End Function
Function FdNNDbl(F, Optional Dft$) As Dao.Field2: Set FdNNDbl = FdNwNN(F, dbDouble, Dft:=Dft): End Function

Function FdTxt(F, Optional TxtSi As Byte = 255, Optional ZLen As Boolean) As Dao.Field2: Set FdTxt = FdNw(F, dbText, True, TxtSi, ZLen): End Function
Function FdNNTxt(F, Optional TxtSi As Byte, Optional ZLen As Boolean, Optional Dft$) As Dao.Field2: Set FdNNTxt = FdNwNN(F, dbText, TxtSi, ZLen, Dft:=Dft): End Function

Function FdFk(F) As Dao.Field2
Set FdFk = New Dao.Field
With FdFk
    .Name = F
    .Type = dbLong
End With
End Function

Function FdId(Tbn) As Dao.Field2
Const CSub$ = CMod & "FdId"
If HasSfx(Tbn, "Id") Then Thw CSub, "Tbn must has Sfx-Id", "Tbn", Tbn
Dim O As New Dao.Field
With O
    .Name = Tbn & "Id"
    .Type = dbLong
    .Attributes = Dao.FieldAttributeEnum.dbAutoIncrField
    .Required = True
End With
Set FdId = O
End Function

Function FdInt(F) As Dao.Field2
Set FdInt = FdNw(F, dbInteger, True, Dft:="0")
End Function

Function FdLng(F) As Dao.Field2
Set FdLng = FdNw(F, dbLong, True, Dft:="0")
End Function

Function FdMem(F) As Dao.Field2
Set FdMem = FdNw(F, dbMemo, True, Dft:="""""")
End Function

Function FdNm(F) As Dao.Field2
If Right(F, 2) <> "Nm" Then Stop
Set FdNm = FdNw(F, dbText, True, 50, False)
End Function

Function FdPk(F) As Dao.Field2
If Right(F, 2) <> "Id" Then Stop
Set FdPk = FdNw(F, dbLong, True)
FdPk.Attributes = Dao.FieldAttributeEnum.dbAutoIncrField
End Function

Function FdShtTys(F, ShtTys) As Dao.Field2
Const CSub$ = CMod & "FdShtTys"
'Public Const ShtTyLis$ = "ABBytCChrDDteDecILMSTTimTxt"
Dim O As Dao.Field2
Select Case ShtTys
Case "Att", "A":  Set O = FdAtt(F)
Case "Bool", "B": Set O = FdBool(F)
Case "Byt":       Set O = FdByt(F)
Case "Chr", "C":  Set O = FdCur(F)
Case "Dte":       Set O = FdDte(F)
Case "Dec":       Set O = FdDec(F)
Case "Dbl", "D":  Set O = FdDbl(F)
Case "Int", "I":  Set O = FdInt(F)
Case "Lng", "L":  Set O = FdLng(F)
Case "Mem", "M":  Set O = FdMem(F)
Case "Sng", "S":  Set O = FdSng(F)
Case "Txt", "T":  Set O = FdTxt(F)
Case "Tim":       Set O = FdTim(F)
Case Else:
    If ChrFst(ShtTys) = "T" Then
        Dim Si As Byte
        Si = CByte(RmvFst(ShtTys))
        Set O = FdTxt(F, Si)
        Exit Function
    End If
    Thw CSub, "ShtTys Err", "ShtTys", ShtTys
End Select
Set FdShtTys = O
End Function

Function FdTy(F, T As Dao.DataTypeEnum) As Dao.Field

End Function
Function FdSng(F) As Dao.Field2
Set FdSng = FdNw(F, dbSingle, True, Dft:="0")
End Function


Function FdTim(F) As Dao.Field2
Set FdTim = FdNw(F, dbTime, True, Dft:="0")
End Function

Function FdTnnn(F, EleTnnn) As Dao.Field2
If Left(EleTnnn, 1) <> "T" Then Exit Function
Dim A$
A = Mid(EleTnnn, 2)
If CStr(Val(A)) <> A Then Exit Function
Set FdTnnn = FdNw(F, dbText, True)
With FdTnnn
    .Size = A
    .DefaultValue = """"""
    .AllowZeroLength = True
End With
End Function
