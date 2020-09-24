Attribute VB_Name = "MxDao_Att_AttOp_DltAtt"
Option Compare Text
Const CMod$ = "MxDao_Att_Do_DltAtt."
Option Explicit
Sub DltAtt(D As Database, Attn$, Attf$)
D.Execute FmtQQ("Delete * from Attd where AttFn_='?' and AttId in (Select AttId from Att where Attn = '?')", Attf, Attn)
Dim R As Dao.Recordset: Set R = RsAtt(D, Attn)
Dim Rs2 As Dao.Recordset2: Set Rs2 = R.Fields("Att").Value
With Rs2
    While .EOF
        If Rs2!FileName = Attf Then
            R.Edit
            .Delete
            R.Update
            Exit Sub
        End If
    Wend
End With
End Sub
Sub DltAttC(Attn$, Attfn$): DltAtt CDb, Attn, Attfn: End Sub
