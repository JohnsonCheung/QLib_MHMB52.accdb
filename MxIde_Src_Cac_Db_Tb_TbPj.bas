Attribute VB_Name = "MxIde_Src_Cac_Db_Tb_TbPj"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_SrcFtcac_Db_Tb_Pj."
Sub RfhTbPjP(): RfhTbPj CPj: End Sub
Sub RfhTbPj(P As VBProject)
If Not WHasRec(P.Name) Then WSetSrcIns P.Name
End Sub
Private Function WHasRec(Pjn$) As Boolean: WHasRec = HasRecQ(CDb, "Select * from Pj where Pjn='" & Pjn & "'"):   End Function
Private Sub WSetSrcIns(Pjn$):                        DoCmd.RunSQL "Insert into Pj (Pjn) values ('" & Pjn & "')": End Sub
