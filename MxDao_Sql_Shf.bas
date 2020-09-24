Attribute VB_Name = "MxDao_Sql_Shf"
Option Compare Text
Option Explicit
Function IsShfAs(OLn$) As Boolean: IsShfAs = IsShfPfxSpc(OLn, "As"): End Function
Function IsShfOn(OLn$) As Boolean: IsShfOn = IsShfPfxSpc(OLn, "On"): End Function
Function ShfSqlTm$(OLn$)
If ChrFst(OLn) = "[" Then
    Dim P%: P = PosBktCls(OLn, 1, "[")
    ShfSqlTm = Left(OLn, P)
    OLn = LTrim(Mid(OLn, P + 1))
    Exit Function
End If
ShfSqlTm = ShfNm(OLn)
End Function
