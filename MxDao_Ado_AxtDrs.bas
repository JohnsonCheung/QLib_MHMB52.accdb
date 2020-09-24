Attribute VB_Name = "MxDao_Ado_AxtDrs"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Ado_AxtDrs."
Public Const AxtFf$ = "Tbn Name Type DefinedSize NumericScale Precision ZRelatedColumn SortOrder"
Private Sub B_DrsTAxDcFxw()
BrwDrs DrsTAxDcFxw(MH.MB52Las.Fxi, MH.MB52IO.WsnFxi)
End Sub
Function DrsTAxDcFxw(Fx, W) As Drs
Dim C As Catalog: Set C = CatFx(Fx)
Dim T As Table: Set T = C.Tables(Axtn(W))
DrsTAxDcFxw = DrsTAxCol(T)
End Function

Function DrsTAxCol(T As ADOX.Table) As Drs
DrsTAxCol = DrsFf(AxtFf, ZAxColDy(T))
End Function

Private Function ZAxColDy(T As ADOX.Table) As Variant()
Dim C As ADOX.Column: For Each C In T.Columns
    PushI ZAxColDy, ZAxColDr(T.Name, C)
Next
End Function

Private Function ZAxColDr(Tbn$, C As ADOX.Column) As Variant()
With C
ZAxColDr = Array(Tbn, .Name, .Type, .DefinedSize, .NumericScale, .Precision, ZRelatedColumn(C), ZSortOrder(C))
End With
End Function

Private Function ZRelatedColumn$(C As ADOX.Column)
On Error Resume Next
ZRelatedColumn = C.RelatedColumn
End Function

Private Function ZSortOrder(C As ADOX.Column) As SortOrderEnum
On Error Resume Next
ZSortOrder = C.SortOrder
End Function
