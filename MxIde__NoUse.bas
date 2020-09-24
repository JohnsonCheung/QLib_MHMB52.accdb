Attribute VB_Name = "MxIde__NoUse"
Option Compare Text
Option Explicit



Function ShfRmk$(OLn$)
Dim L$
L = LTrim(OLn)
If ChrFst(L) = "'" Then
    ShfRmk = Mid(L, 2)
    OLn = ""
End If
End Function
