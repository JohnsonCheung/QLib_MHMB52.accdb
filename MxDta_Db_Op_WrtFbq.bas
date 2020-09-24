Attribute VB_Name = "MxDta_Db_Op_WrtFbq"
':Fbka: :Ft #Fil-name-of-BacK-Apostrophe#
' ! It is a fil of ext *.bka.txt.  There 1-Sgn-Ln, 0-to-N-Rmk-Lines and 0-To-N-Tbl-Lines.
' ! Sgn-Ln is  **BackApostropheSeparatedFile**<Dsn>**, where <Dsn> is a dta-set-nm.
' ! Rmk-Lines are lines between Sgn-Ln and (fst-T-Ln or eof)
' ! 1-Tbl-Lines is  1-T-Ln, 0-to-N-TblRmk-Lines, 1-Fld-Ln and 0-to-N-Dta-Lines.
' ! T-Ln           is
' ! Rmk-Lines are lines before fst *-Ln.  Rmk are for all tbl in the :Fbka:  Each individual tbl does not have it own rmk
' ! Lines before fst *-Ln are Rmk.  Each gp of one-*-Ln & N-`-Ln is one tbl.
' ! *-Ln is a Ln wi fst chr is *, :Starl: #Star-Line#.  `-Ln is a Ln wi fst chr is `, :Bkal:, #BacK-Apostrophe-Ln#.
' ! The *-Ln is *<T>
' ! The fst `-Ln is :Scff
' ! The rst `-Ln is :dta
':T:   :S  #Table-Name#
':Scff: :Ff #ShtTyc-Colon-Ff#  ! It is spc sep of :Scfld:.  It desc ty and fldn of the tbl.
'!It has first line as ShtTyscfQBLin.
'!It rest of lines are records."
Option Compare Text
Option Explicit
Const CMod$ = "MxDta_Db_Op_WrtFbq."

Sub InsRsBql(R As Dao.Recordset, Bql$)
R.AddNew
Dim Ay$(): Ay = Split(Bql, "`")
Dim F As Dao.Field, J%
For Each F In R.Fields
    If Ay(J) <> "" Then
        F.Value = Ay(J)
    End If
    J = J + 1
Next
R.Update
End Sub

Private Sub B_WrtFbkquoDb()
Dim P$: P = PthTmpInst
WrtFbkquoDb P, MHO.MHODuty.DbDta
BrwPth P
Stop
End Sub

Private Sub B_WrtFbkquoDbC()
Dim P$: P = PthTmpInst("Fbkquo")
WrtFbkquoDb P, MHO.MHODuty.DbDta
BrwPth P
End Sub
Sub WrtFbkquoDbC(Pth): WrtFbkquoDb Pth, CDb: End Sub

Sub WrtFbkquoDb(Pth, D As Database): WrtFbkquoTny Pth, D, Tny(D): End Sub

Sub WrtFbkquoTny(Pth, D As Database, Tny$())
Dim T: For Each T In Tny
    WrtFbkquoT Pth, D, T
Next
End Sub
Sub WrtFbkquoT(Pth, D As Database, T)
Dim F%: F = FnoO(PthEnsSfx(Pth) & T & ".bkquo.txt")
Dim R As Dao.Recordset
Set R = RsTblNoAtt(D, T)
Print #F, JnBkquo(Fny(D, T))
Print #F, WLnTy(Td(D, T).Fields)
With R
    While Not .EOF
        Print #F, WLnRs(R)
        .MoveNext
    Wend
    .Close
End With
Close #F
End Sub
Private Sub B_WLnTy()
Debug.Print WLnTy(RsTbl(CDb, "PErmit").Fields)
End Sub
Private Function WLnTy$(F As Dao.Fields)
Dim O$()
Dim Fd As Dao.Field: For Each Fd In F
    PushI O, NmEnmSimTy(SimTyDao(Fd.Type))
Next
WLnTy = JnBkquo(O)
End Function
Private Function WLnRs$(A As Dao.Recordset)
Dim O$(), F As Dao.Field
For Each F In A.Fields
    If IsNull(F.Value) Then
        PushI O, ""
    Else
        PushI O, Replace(Replace(F.Value, vbCr, ""), vbLf, " ")
    End If
Next
Dim L$: L = JnBkquo(O)
If L = "401`HD0V4FOF00C9ZT" Then Stop
WLnRs = L

End Function

Function JnBkquo$(Ay): JnBkquo = Jn(Ay, "`"): End Function
