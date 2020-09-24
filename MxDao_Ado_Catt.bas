Attribute VB_Name = "MxDao_Ado_Catt"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Ado_Catt."
Type TCatTbl: C As Catalog: T As ADOX.Table: End Type
Function TCatTbl(C As Catalog, T As ADOX.Table) As TCatTbl
With TCatTbl
    Set .C = C
    Set .T = T
End With
End Function
Function TCatTblFxw(Fx$, Optional W$) As TCatTbl
Dim C As Catalog, T As ADOX.Table
Set C = CatFx(Fx)
Set T = C.Tables(Axtn(W))
TCatTblFxw = TCatTbl(C, T)
End Function
Function TCatTblFbt(Fb$, T$) As TCatTbl
Dim C As Catalog, Td As ADOX.Table
Set C = CatFb(Fb)
Set Td = C.Tables(T)
TCatTblFbt = TCatTbl(C, Td)
End Function
