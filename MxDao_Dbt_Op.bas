Attribute VB_Name = "MxDao_Dbt_Op"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Dbt_Op."
Public Const C_Des$ = "Description"
Public Const SqlTbnMSysObj$ = "Select Name from MSysObjects where Type in (1,6) and Name not Like 'MSys*' and Name not Like 'f_*_Data'"
Sub BrwDb(D As Database): BrwFb D.Name: End Sub
Sub ClsDbAp(ParamArray ApDb())
Dim Av(): Av = ApDb
Dim Db: For Each Db In Av
    ClsDb CvDb(Db)
Next
End Sub
Sub ClsCnAp(ParamArray ApCn())
Dim Av(): Av = ApCn
Dim Cn: For Each Cn In Av
    ClsCn CvCn(Cn)
Next
End Sub
Sub ClsCn(C As ADODB.Connection)
On Error Resume Next
C.Close
End Sub
Sub ClsDb(D As Database):
On Error Resume Next
D.Close
End Sub

Sub Crtt(D As Database, T, SpecSqlFldLn$): D.Execute FmtQQ("Create Table [?] (?)", T, SpecSqlFldLn): End Sub
Sub CrttTbHshAC():                         CrttTbHshA CDb:                                           End Sub
Sub CrttTbHshA(D As Database)
Drp D, "#A"
D.TableDefs.Append WTd
End Sub
Private Function WTd() As Dao.TableDef
Dim Fdy() As Dao.Field
PushObj Fdy, FdTxt("F1")
Set WTd = TdFdy("#A", Fdy)
End Function

Sub EnsTmpTbl(D As Database)
If HasT(D, "#Tmp") Then Exit Sub
D.Execute "Create Table [#Tmp] (AA Int, BB Text 10)"
End Sub

Sub SetFldDesByDi(D As Database, TFDes As Dictionary)
Dim T$, F$, Des$, TDotF$, I, J
For Each I In TFDes.Keys
    TDotF = I
    Des = TFDes(TDotF)
    If HasDot(TDotF) Then
        AsgBrkDot TDotF, T, F
        SetPvFDes D, T, F, Des
    Else
        For Each J In Tny(D)
            T = J
            If HasFld(D, T, F) Then
                SetPvFDes D, T, F, Des
            End If
        Next
    End If
Next
End Sub
