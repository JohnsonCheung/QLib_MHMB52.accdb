Attribute VB_Name = "MxDta_SimTy"
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_SimTy."
Enum eSimTy: eSimTyUnk: eSimTyBool: eSimTyDte: eSimTyNbr: eSimTyTxt: End Enum: Public Const EnmmSimTy$ = "eSimTyUnk eSimTyBool eSimTyDte eSimTyNbr eSimTyTxt"
Public Const NnEnmSimTy$ = "U B D N T"
Function NmEnmSimTy$(E As eSimTy): NmEnmSimTy = NyEnmSimTy()(E):   End Function
Function EnmsSimTy$(E As eSimTy):   EnmsSimTy = EnmsySimTy()(E):   End Function
Function EnmsySimTy() As String(): EnmsySimTy = NyQtp(EnmmSimTy):  End Function
Function SimTyVal(V) As eSimTy:      SimTyVal = SimTy(VarType(V)): End Function
Function EnmSimTyNm(Nm$) As eSimTy
Const CSub$ = CMod & "EnmSimTyNm"
Dim I%: I = IxEle(NyEnmSimTy, Nm)
If I = -1 Then Thw CSub, "Given NmEnmSimTy is invalid", "[Invalid @NmEnmSimTy] [Valid NmEnmSimTy]", Nm, NnEnmSimTy

End Function
Function NyEnmSimTy() As String()
Static Ny$(): If Si(Ny) = 0 Then Ny = SplitSpc(NnEnmSimTy)
NyEnmSimTy = Ny
End Function
Function SimTy(T As VbVarType)
Const CSub$ = CMod & "SimTy"
Dim O As eSimTy
Select Case True
Case T = T = vbByte, T = vbInteger, T = vbLong, T = vbCurrency, T = vbDecimal, T = vbDouble, T = vbSingle: O = eSimTyNbr
Case T = vbBoolean: O = eSimTyBool
Case T = vbDate: O = eSimTyDte
Case T = vbString: O = eSimTyTxt
Case Else
    Thw CSub, "Given @VbTy is not valid", "@VbTy", T
End Select
SimTy = O
End Function
Function SimTyDc(Dc()) As eSimTy
Dim V: For Each V In Itr(Dc)
    Dim O As eSimTy: O = SimTyMax(O, SimTyVal(V))
    If O = eSimTyTxt Then SimTyDc = O: Exit Function
Next
End Function
Function SimTyMax(A As eSimTy, B As eSimTy) As eSimTy: SimTyMax = Max(A, B): End Function

Function SimTyyLo(L As ListObject) As eSimTy()
Dim Sq(): Sq = SqLo(L)
Dim C%: For C = 1 To UBound(Sq, 2)
    PushI SimTyyLo, SimTyDc(DcSq(Sq, C))
Next
End Function

Function SimTyDao(T As Dao.DataTypeEnum) As eSimTy
Const CSub$ = CMod & "SimTyDao"
Dim O As eSimTy
Select Case True
Case T = dbBigInt, T = dbByte, T = dbInteger, T = dbNumeric, T = dbLong, T = dbCurrency, T = dbDecimal, T = dbDouble, T = dbFloat, T = dbSingle: O = eSimTyNbr
Case T = vbBoolean: O = eSimTyBool
Case T = dbDate, T = dbTime: O = eSimTyDte
Case T = dbText, T = dbMemo, T = dbChar: O = eSimTyTxt
Case Else
    Thw CSub, "Given @Daoty is not valid", "@Daoty", T
End Select
SimTyDao = O
End Function
