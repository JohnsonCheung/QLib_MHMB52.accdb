Attribute VB_Name = "MxAcs_Acs_AcsObj_AcsObjOp"
Option Compare Text
Option Explicit
Const CMod$ = "MxAcs_Acs_OpnAcsObj."
Private Sub B_FbOpnAcs()
Dim A As Database: Set A = AcsNw.CurrentDb
End Sub

Sub OpnAcsFb(Fb, A As Access.Application, Optional IsExl As Boolean)
If FbAcs(A) = Fb Then Exit Sub
ClsAcsDb A
A.OpenCurrentDatabase Fb, IsExl
End Sub
Sub ChkInFbTbn(Fb$, Tbn, Optional Fun$ = CMod & "ChkInFbTbn")
If Not InFbTbn(Fb, Tbn) Then Thw Fun, "In @Fb should have @Tbn", "@Fb @Tbn", Fb, Tbn
End Sub
Function InFbTbn(Fb$, Tbn) As Boolean: InFbTbn = HasEle(TnyFb(Fb), Tbn): End Function
Sub BrwFb(Fb)
Static Acs As New Access.Application
OpnAcsFb Fb, Acs
If Not Acs.Visible Then Acs.Visible = True
End Sub

Sub BrwQC(Q$, Optional QnPfx$ = "#BrwQry_"): BrwQ CDb, Q, QnPfx: End Sub
Sub BrwQ(D As Database, Q$, Optional QnPfx$ = "#BrwQry_")
Dim Qn$: Qn = QnTmpCrt(D, Q, QnPfx)
Dim A As Access.Application: Set A = AcsDb(D): If Not A.Visible Then A.Visible = True
A.DoCmd.OpenQuery Qn, acViewNormal
End Sub

Sub BrwTbPfxC(Pfx$):               BrwTbPfx CDb, Pfx:        End Sub
Sub BrwTbPfx(D As Database, Pfx$): BrwTny D, TnyPfx(D, Pfx): End Sub

Sub BrwTbQtp2C(Qtp1$, Rst$):               BrwTbQtp2 CDb, Qtp1, Rst:    End Sub
Sub BrwTbQtp2(D As Database, Qtp1$, Rst$): BrwTny D, NyQtp2(Qtp1, Rst): End Sub
Sub BrwTbQtpC(Qtp$):                       BrwTbQtp CDb, Qtp:           End Sub
Sub BrwTbQtp(D As Database, Qtp$):         BrwTny D, NyQtp(Qtp):        End Sub
Sub BrwTbRo(D As Database, T):
Stop '
End Sub
Sub BrwTbRoC(T):                   BrwTbRo CDb, T:                                 End Sub
Sub BrwT(D As Database, T):        AcsFb(D.Name).DoCmd.OpenTable T, , acReadOnly:  End Sub
Sub BrwTC(T):                      BrwT CDb, T:                                    End Sub
Sub BrwTTC(TT$):                   BrwTT CDb, TT:                                  End Sub
Sub BrwTnyC(Tny$()):               BrwTny CDb, Tny:                                End Sub
Sub BrwTny(D As Database, Tny$()): Dim T: For Each T In Itr(Tny): BrwT D, T: Next: End Sub
Sub BrwTT(D As Database, TT$):     BrwTny D, Tmy(TT):                              End Sub

Sub ClsTblAllC(): ClsTblAll Acs: End Sub
Sub ClsTblAll(A As Access.Application)
Dim T: For Each T In Itr(TnyAcs(A))
    A.DoCmd.Close acTable, T, acSaveYes
Next
End Sub

Sub ClsTblSav(A As Access.Application, T): A.DoCmd.Close acTable, T, acSaveYes: End Sub
Sub ClsTbl(A As Access.Application, T):    A.DoCmd.Close acTable, T:            End Sub
Sub ClsFrm(Frmn$):                         DoCmd.Close acForm, Frmn, acSaveYes: End Sub

Sub ClsFrmAll(): ClsFrmAllA Acs: End Sub
Sub ClsFrmAllA(A As Access.Application)
While A.Forms.Count > 0
A.DoCmd.Close acForm, A.Forms(0).Name, acSaveYes
Wend
End Sub

Sub ClsRptAll(): ClsRptAllA Acs: End Sub
Sub ClsRptAllA(A As Access.Application)
While A.Reports.Count > 0
A.DoCmd.Close acReport, A.Reports(0).Name, acSaveYes
Wend
End Sub
Sub ClsAcsDb(A As Access.Application): On Error Resume Next: A.CurrentDb.Close: End Sub
