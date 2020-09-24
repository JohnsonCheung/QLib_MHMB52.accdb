Attribute VB_Name = "MxDao_Def_Qd"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Def_Qd."
Function Qd(D As Database, Qn) As Dao.QueryDef:  Set Qd = D.QueryDefs(Qn):   End Function
Function QdC(Qn) As Dao.QueryDef:               Set QdC = Qd(CDb, Qn):       End Function
Sub BrwQdC():                                             BrwQd CDb:         End Sub
Sub BrwQd(D As Database):                                 BrwS12y S12yQd(D): End Sub

Function S12yQdC() As S12(): S12yQdC = S12yQd(CDb): End Function
Function S12yQd(D As Database) As S12()
Dim Q: For Each Q In Itr(Qny(D))
    PushS12 S12yQd, S12(Q, FmtSql(SqlQn(D, Q)))
Next
End Function

Function QnyC() As String():             QnyC = Qny(CDb):                                                                                                 End Function
Function Qny(D As Database) As String():  Qny = DcStrQ(D, "Select Name from MSysObjects where Type=5 and Left(Name,4)<>'MSYS' and Left(Name,4)<>'~sq_'"): End Function
