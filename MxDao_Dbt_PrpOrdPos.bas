Attribute VB_Name = "MxDao_Dbt_PrpOrdPos"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Dbt_PrpOrdPos."
Function DrsTOrdPos(D As Database, T) As Drs: DrsTOrdPos = Drs(SySs("OrdPos Fldn"), OrdPosDy(D, T)): End Function
Private Function OrdPosDy(D As Database, T) As Variant()
Dim J%: Dim F As Dao.Field: For Each F In D.TableDefs(T).Fields
    PushI OrdPosDy, Array(F.OrdinalPosition, F.Name)
    J = J + 1
Next
End Function

Sub DmpOrdPos(D As Database, T):                              DmpAy FmtOrdPos(D, T):                    End Sub
Sub DmpCOrdPos(T):                                            DmpOrdPos CDb, T:                         End Sub
Function FmtOrdPos(D As Database, T) As String(): FmtOrdPos = FmtDrs(DrsTOrdPos(D, T)):                 End Function
Function OrdPos%(D As Database, T, F):               OrdPos = D.TableDefs(T).Fields(F).OrdinalPosition: End Function
Function COrdPos%(T, F):                            COrdPos = OrdPos(CDb, T, F):                        End Function

Function MaxOrdPosTd%(T As Dao.TableDef): MaxOrdPosTd = MaxItp(T.Fields, "OrdinalPosition"): End Function
Function MaxOrdPos%(D As Database, T):      MaxOrdPos = MaxOrdPosTd(Td(D, T)):               End Function
Function CMaxOrdPos%(T):                   CMaxOrdPos = MaxOrdPos(CDb, T):                   End Function
