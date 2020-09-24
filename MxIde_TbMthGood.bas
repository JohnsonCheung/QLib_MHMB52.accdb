Attribute VB_Name = "MxIde_TbMthGood"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_TbMthGood."
Sub RfhTbMthGood()
WCrtTbTmpMthn4New TMthmdnyPC
Dim JnOn$: Stop 'JnOn = QpJnInr("Mdn Mthn ShtTy ShtMdy")
RunqC "Insert into MthGood (Mdn,Mthn,ShtTy,ShtMdy,IsGood) Select x.Mdn,x.Mthn,x.ShtTy,x.ShtMdy,False from [#TmpNew] x inner join MthGood a " & JnOn
RunqC "Delete * from MthGood x inner join [#Mthn4Dlt] a " & JnOn
End Sub
Private Sub WCrtTbTmpMthn4New(N() As TMthmdn)
DrpC "#Mthn4New"
RunqC "Create Table [#Mthn4New] (Mdn Text(255),Mthn Text(255), ShtTy Text(3),ShtMdy Text(3))"
Dim R As Dao.Recordset: Set R = RsTblC("#Mthn4New")
Dim J&: For J = 0 To UbTMthmd(N)
    With N(J)
        R.AddNew
        R!Mdn = .Mdn
        R!Mthn = .TMth.Mthn
        R!ShtMdy = .TMth.ShtMdy
        R!ShtTy = .TMth.ShtTy
        R.Update
    End With
Next
End Sub
