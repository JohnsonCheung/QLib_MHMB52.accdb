Attribute VB_Name = "MxIde_Tyn"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Tyn."
Function IsEnmn(Nm$) As Boolean
Static X$(): If Si(X) = 0 Then X = EnmnyPC
IsEnmn = HasEle(X, Nm)
End Function

Function IsEnmnKnown(Nm$) As Boolean
Static X$(): If Si(X) = 0 Then X = Sy("AdoDb.DataTypeEnum")
IsEnmnKnown = HasEle(X, Nm)
End Function

Function IsTynObj(Tyn$) As Boolean ' return true if @Tyn (isBlnk | IsTynPrim | IsUdtn | IsEnmn)
If Tyn = "" Then Exit Function
If IsTynPrim(Tyn) Then Exit Function
If IsUdtn(Tyn) Then Exit Function
If IsEnmn(Tyn) Then Exit Function
If IsEnmnKnown(Tyn) Then Exit Function
IsTynObj = True
End Function

Function Shttyn$(Tyn$) ' return ShtTyn for some known class.  Used in Msig
Dim O$
Select Case Tyn
Case "VbProject": O = "Pj"
Case "Access.Application": O = "Acs"
Case "Access.Control": O = "AcsCtl"
Case "Access.CommandButton": O = "ABtn"
Case "Access.ToggleButton": O = "ATgl"
Case "Excel.Application": O = "Xls"
Case "Excel.Addin": O = "XlsAddin"
Case "Range": O = "Rg"
Case "ListObject": O = "Lo"
Case "ListObject()": O = "LoAy"
Case "Excel.Worksheet", "Worksheet": O = "Ws"
Case "Excel.Workbook", "Workbook": O = "Wb"
Case "ADODB.Recordset": O = "Ars"
Case "ADODB.Connection": O = "Cn"
Case "ADOX.Table": O = "AdoTd"
Case "ADODB.DataTypeEnum": O = "AdoTy"
Case "VBA.Collection", "Collection": O = "Coll"
Case "Dictionary": O = "Dic"
Case "CodeModule": O = "Md"
Case "VBComponent": O = "Cmp"
Case "vbext_ComponentType": O = "eCmpTy"
Case "Database": O = "Db"
Case "Variant": O = "Var"
Case Else: O = Tyn
End Select
Shttyn = O
End Function
