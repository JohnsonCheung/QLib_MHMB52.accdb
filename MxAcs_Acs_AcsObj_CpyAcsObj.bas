Attribute VB_Name = "MxAcs_Acs_AcsObj_CpyAcsObj"
Option Compare Text
Option Explicit
Const CMod$ = "MxAcs_Acs_CpyAcsObj."

Private Sub B_CpyAcsObj()
Dim FbFm$: FbFm = MHO.MHOMB52.FbPgm
Dim FbTo$: FbTo = FbTmp
CrtFb FbTo
CpyAcsObj FbFm, FbTo
Debug.Print FbTo
End Sub
Private Sub B_CpyAcsFrm(): CpyAcsFrm FbCorrudFstC, FbRescudFstC: End Sub
Sub CpyAcsFrm(FbFm, FbTo$)
Dim A As Access.Application: Set A = AAcsFb(FbFm, IsExl:=True)
WCpyFrm A, FbTo
ClsAcsDb A
End Sub
Sub CpyAcsObjNoMd(FbFm, FbTo$)
ChkFfnExi FbFm, CSub, "FbFm"
ChkFfnExi FbTo, CSub, "FbTo"
Dim A As Access.Application: Set A = AAcsFb(FbFm, IsExl:=True)
WCpyAcsObjNoMd A, FbTo
ClsAcsDb A
End Sub
Private Sub WCpyAcsObjNoMd(AcsFm As Access.Application, FbTo$)
BegTimr "CpyAcsObj"
Dim A As Access.Application: Set A = AcsFm
WCpyTbl A, FbTo
WCpyFrm A, FbTo
WCpyRpt A, FbTo
WCpyQry A, FbTo
WCpyRf A, FbTo
End Sub

Sub CpyAcsObj(FbFm, FbTo$)
ChkFfnExi FbFm, CSub, "FbFm"
ChkFfnExi FbTo, CSub, "FbTo"
Dim A As Access.Application: Set A = AAcsFb(FbFm, IsExl:=True)
WCpyAcsObjNoMd A, FbTo
WCpyMd A, FbTo
ClsAcsDb A
End Sub

Private Sub WCpyRf(A As Access.Application, FbTo)
Dim PjFm As VBProject: Set PjFm = A.VBE.ActiveVBProject
Dim PjTo As VBProject
    Dim AcsTo As New Access.Application: AcsTo.OpenCurrentDatabase FbTo
    Set PjTo = AcsTo.VBE.ActiveVBProject
Dim R As VbIde.Reference: For Each R In PjFm.References
    ShwTimrNoLf "Reference " & R.Name
    If Not WHasRf(PjTo, R.Guid) Then
        PjTo.References.AddFromGuid R.Guid, R.Major, R.Minor
        Debug.Print "Copied <================="
    Else
        Debug.Print "Skipped"
    End If
Next
End Sub
Private Function WHasRf(PjTar As VBProject, Guid$) As Boolean
Dim R As VbIde.Reference: For Each R In PjTar.References
    If R.Guid = Guid Then WHasRf = True: Exit Function
Next
End Function

Private Sub WCpyFrm(A As Access.Application, FbTo$)
Dim F: For Each F In Itr(Frmny(A.CurrentDb))
    WShwMsg F, acForm
    A.DoCmd.CopyObject FbTo, , acForm, F
Next
End Sub

Private Sub WCpyRpt(A As Access.Application, FbTo$)
Dim FbFm$: FbFm = A.Name
Dim R: For Each R In Itr(Rptny(A.CurrentDb))
    WShwMsg R, acReport
    A.DoCmd.CopyObject FbTo, , acReport, R
Next
End Sub

Private Sub WCpyMd(A As Access.Application, FbTo$)
Dim FbFm$: FbFm = A.Name
Dim Ny$(): Ny = MdnyP(PjMainAcs(A))
Dim N%: N = Si(Ny)
Dim M, J%: For Each M In Itr(Ny)
    WShwMsg M, acModule, J, N
    A.DoCmd.CopyObject FbTo, , acModule, M
Next
End Sub

Private Sub WCpyTbl(A As Access.Application, FbTo$)
Dim DbTo As Database: Set DbTo = Db(FbTo)
Dim Ny$(), N%
Ny = Tny(A.CurrentDb): N = Si(Ny)
Dim T, J%: For Each T In Itr(Ny)
    WShwMsg T, acTable, J, N
    CpyAcsTbl A, T, FbTo
Next
End Sub
Private Sub WShwMsg(Nm, T As AcObjectType, Optional OIx%, Optional N%)
DoEvents
Dim P$
If N > 0 Then
    P = " " & OIx & " of " & N & " "
    OIx = OIx + 1
End If
ShwTimr EnmsAcObjTy(T) & P & Nm
End Sub
Private Sub WCpyQry(A As Access.Application, FbTo$)
Dim I As QueryDef: For Each I In A.CurrentDb.QueryDefs
    WShwMsg I.Name, acQuery
    A.DoCmd.CopyObject FbTo, , acQuery, I.Name
Next
End Sub

Sub CpyAcsTbl(A As Access.Application, T, FbTo$)
ChkFbtNExi FbTo, T
A.DoCmd.CopyObject FbTo, , acTable, T
End Sub
Sub CpyAcsTblC(T, FbTo$): CpyAcsTbl Acs, T, FbTo: End Sub
