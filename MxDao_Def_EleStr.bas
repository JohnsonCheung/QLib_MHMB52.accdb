Attribute VB_Name = "MxDao_Def_EleStr"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Def_EleStr."
Function ErEleStr$(EleStr$)

End Function
Function FdEleStr(F, EleStr$) As Dao.Field2
Stop '
End Function
Function FdE(F, StdEle$) As Dao.Field2
Dim O As Dao.Field2
Set O = FdTnnn(F, StdEle): If Not IsNothing(O) Then Set FdE = O: Exit Function
Select Case StdEle
Case "Nm":  Set FdE = FdNm(F)
Case "Amt": Set FdE = FdCur(F): FdE.DefaultValue = 0
Case "Txt": Set FdE = FdTxt(F, dbText, True): FdE.DefaultValue = """""": FdE.AllowZeroLength = True
Case "Dte": Set FdE = FdDte(F)
Case "Int": Set FdE = FdInt(F)
Case "Lng": Set FdE = FdLng(F)
Case "Dbl": Set FdE = FdDbl(F)
Case "Sng": Set FdE = FdSng(F)
Case "Lgc": Set FdE = FdBool(F)
Case "Mem": Set FdE = FdMem(F)
End Select
End Function
