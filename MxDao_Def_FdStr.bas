Attribute VB_Name = "MxDao_Def_FdStr"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Def_FdStr."
Public Const SSoStdEle$ = "CrtDte Pk Fk Ty Nm Dte Amt Att"
Public Const Daotynn$ = "Boolean Byte Integer Int Long Single Double Char Text Memo Attachment" ' used in TzPFld
Function FdStdFld(StrFd, Optional T) As Dao.Field2
Dim E$: E = EleFmFdStr(StrFd, T)
Set FdStdFld = FdStdEle(StrFd, E)
End Function

Function IsStdFld(F) As Boolean
IsStdFld = EleFmFdStr(F) <> ""
End Function

Function FdStdEle(F, E) As Dao.Field2
Set FdStdEle = F & " " & StdEleStr(E)
End Function

Function StdEleStr$(E)
Const CSub$ = CMod & "StdEleStr"
Dim O$
Select Case E
Case "CrtDte"
Case "Dte"
Case "Pk"
Case "Fk"
Case "Ty"
Case "Nm"
Case "Dte"
Case "Amt"
Case "Att"
Case Else: Thw CSub, "Given Ele is not std", "E", E
End Select
StdEleStr = O
End Function

Function EleFmFdStr$(StrFd, Optional T)
Dim R2$, R3$
R2 = Right(StrFd, 2)
R3 = Right(StrFd, 3)
Dim O$
Select Case True
Case StrFd = "CrtDte":  O = "CrtDte"
Case T & "Id" = StrFd:  O = "Pk"
Case R2 = "Id":     O = "Fk"
Case R2 = "Ty":     O = "Ty"
Case R2 = "Nm":     O = "Nm"
Case R3 = "Dte":    O = "Dte"
Case R3 = "Amt":    O = "Amt"
Case R3 = "Att":    O = "Att"
End Select
EleFmFdStr = O
End Function

Function StdEley() As String()
ClrBfr
BrwBfr
BfrLn
BfrV "Id "
BfrV "*Id"
BfrV "*Id"
BfrV "*Id"
BfrV "*Id"
BfrV "*Id"
BfrV "*Id"
StdEley = LyBfr
End Function

Function StdEleDi() As Dictionary
Static X As Boolean, Y As Dictionary
If Not X Then
    X = True
    Set Y = New Dictionary
End If
Set StdEleDi = Y
End Function

Function StrFd$(F As Dao.Field2)
Dim D$, R$, Z$, VTxt$, VRul, E$, S$
If F.Type = Dao.DataTypeEnum.dbText Then S = " TxtSz=" & F.Size
If F.DefaultValue <> "" Then D = "Dft=" & F.DefaultValue
If F.Required Then R = "Req"
If F.AllowZeroLength Then Z = "AlZZLen"
If F.Expression <> "" Then E = "Epr=" & F.Expression
If F.ValidationRule <> "" Then VRul = "VRul=" & F.ValidationRule
If F.ValidationText <> "" Then VTxt = "VTxt=" & F.ValidationText
StrFd = TmlAp(F.Name, ShtDaoty(F.Type), R, Z, VTxt, VRul, D, E, IIf((F.Attributes And Dao.FieldAttributeEnum.dbAutoIncrField) <> 0, "Auto", ""))
End Function

Function FdFmStr(StrFd) As Dao.Field2
Const CSub$ = CMod & "FdFmStr"
Dim N$, S$ ' #Fldn and #EleStr
Dim O As Dao.Field2
AsgBrkSpc StrFd, N, S
Select Case True
Case S = "Boolean":  Set O = FdBool(N)
Case S = "Byte":     Set O = FdByt(N)
Case S = "Integer", S = "Int": Set O = FdInt(N)
Case S = "Long":     Set O = FdLng(N)
Case S = "Single":   Set O = FdSng(N)
Case S = "Double":   Set O = FdDbl(N)
Case S = "Currency": Set O = FdCur(N)
Case S = "Char":     Set O = FdChr(N)
Case HasPfx(S, "Text"): Set O = FdTxt(N, BetBkt(S))
Case S = "Memo":     Set O = FdMem(N)
Case S = "Attachment": Set O = FdAtt(N)
Case S = "Time":     Set O = FdTim(N)
Case S = "Date":     Set O = FdDte(N)
Case Else: Thw CSub, "Invalid StrFd", "Nm Spec vdt-Daotynn, N, S, Daotynn"
End Select
Set FdFmStr = O
End Function

Function FdStr(StrFd$) As Dao.Field2
Dim F$, ShtTy$, Req As Boolean, AlZZLen As Boolean, Dft$, VTxt$, VRul$, TxtSz As Byte, Epr$
Dim L$: L = StrFd
Dim Vy(): Vy = ShfVy(L, EleLblss)
AsgAy Vy, _
    F, ShtTy, Req, AlZZLen, Dft, VTxt, VRul, TxtSz, Epr
Set FdStr = FdNw( _
    F, DaotyShtTy(ShtTy), Req, TxtSz, AlZZLen, Epr, Dft, VRul, VTxt)
End Function

Function SyFd(A As Dao.Fields) As String()
Dim F As Dao.Field: For Each F In A
    PushI SyFd, StrFd(F)
Next
End Function


Private Sub B_FdStr()
GoSub T2
Exit Sub
Dim Act As Dao.Field2, Ept As Dao.Field2, StrFd_$
T2:
    StrFd_ = "AA Int Req AlZZLen Dft=ABC TxtSz=10"
    Set Ept = New Dao.Field
    With Ept
        .Type = Dao.DataTypeEnum.dbInteger
        .Name = "AA"
        '.AllowZeroLength = False
        .DefaultValue = "ABC"
        .Required = True
        .Size = 2
    End With
    GoTo Tst
T1:
    StrFd_ = "Txt Req Dft=ABC [VTxt=Loc must cannot be blank] [VRul=IsNull([Loc]) or Trim(Loc)='']"
    GoTo Tst
Tst:
    Set Act = FdStr(StrFd_)
    If Not IsEqFd(Act, Ept) Then
        D MsgyNNAp("Act", "StrFd", StrFd(Act))
        D MsgyNNAp("Ept", "StrFd", StrFd(Ept))
    End If
    Return
End Sub
