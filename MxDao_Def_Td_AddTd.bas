Attribute VB_Name = "MxDao_Def_Td_AddTd"
Option Compare Text
Const CMod$ = "MxDao_Def_Td_AddTd."
Option Explicit
Sub AppTdy(D As Database, TdAy() As Dao.TableDef)
Dim T: For Each T In Itr(TdAy)
    D.TableDefs.Append T
Next
End Sub

Sub AddFdTxt(T As Dao.TableDef, FF$, Optional Req As Boolean, Optional Si As Byte = 255)
Dim F: For Each F In FnyFF(FF)
    T.Fields.Append FdNw(F, dbText, Req, Si)
Next
End Sub
Sub AddFdMem(T As Dao.TableDef, FF$): AddFdy T, W_Fdy(FF, dbMemo): End Sub
Sub AddFdLng(T As Dao.TableDef, FF$): AddFdy T, W_Fdy(FF, dbLong): End Sub
Function TdFdy(T, Fdy() As Dao.Field) As Dao.TableDef
Dim O As New TableDef
O.Name = T
Dim F: For Each F In Fdy
    O.Fields.Append F
Next
Set TdFdy = O
End Function

Sub AddFdEpr(T As Dao.TableDef, F$, Epr$, Ty As Dao.DataTypeEnum): T.Fields.Append FdNw(F, Ty, Epr:=Epr):           End Sub
Sub AddFdTimstmp(T As Dao.TableDef, F$):                           T.Fields.Append FdNw(F, Dao.dbDate, Dft:="Now"): End Sub
Sub AddFdy(T As Dao.TableDef, Fdy() As Dao.Field)
Dim F: For Each F In Fdy
    T.Fields.Append F
Next
End Sub
Sub AddFdId(T As Dao.TableDef): T.Fields.Append FdId(T.Name): End Sub

Private Function W_Fdy(FF$, T As Dao.DataTypeEnum) As Dao.Field()
Dim F: For Each F In FnyFF(FF)
    PushObj W_Fdy, FdNw(F, T)
Next
End Function
