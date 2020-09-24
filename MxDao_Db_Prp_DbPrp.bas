Attribute VB_Name = "MxDao_Db_Prp_DbPrp"
Option Compare Text
Option Explicit
Const CMod$ = "MxDao_Db_Prp."

Function IsDbOk(D As Database) As Boolean
On Error GoTo X
IsDbOk = D.Name = D.Name
Exit Function
X:
End Function

Function FrmnyC() As String():                     FrmnyC = Frmny(CDb):                End Function
Function Frmny(D As Database) As String():          Frmny = Itn(CntrFrm(D).Documents): End Function
Function CntrFrm(D As Database) As Container: Set CntrFrm = D.Containers("Forms"):     End Function
Function CntrMd(D As Database) As Container:   Set CntrMd = D.Containers("Modules"):   End Function
Function CntrMdC() As Container:              Set CntrMdC = CntrMd(CDb):               End Function

Function RptnyC() As String():                     RptnyC = Rptny(CDb):                End Function
Function Rptny(D As Database) As String():          Rptny = Itn(CntrRpt(D).Documents): End Function
Function CntrRpt(D As Database) As Container: Set CntrRpt = D.Containers("Reports"):   End Function

Function FrmnyOpn() As String()
Dim F As Access.Form: For Each F In Acs.Forms
    PushI FrmnyOpn, F.Name
Next
End Function
