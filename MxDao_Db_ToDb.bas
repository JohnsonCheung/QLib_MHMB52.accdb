Attribute VB_Name = "MxDao_Db_ToDb"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Db_ToDb."

Function CvDb(A) As Database: Set CvDb = A:                             End Function
Function Db(Fb) As Database:    Set Db = Dao.DBEngine.OpenDatabase(Fb): End Function
Function DbIf(ODb As Database, Fb)
If IsNothing(ODb) Then Set ODb = Db(Fb)
Set DbIf = ODb
End Function
