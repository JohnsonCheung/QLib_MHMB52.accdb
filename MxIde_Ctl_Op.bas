Attribute VB_Name = "MxIde_Ctl_Op"
Option Compare Text
Option Explicit
Const CMod$ = "MxIde_Ctl_Op."
Sub NxtStmt()
Const CSub$ = CMod & "NxtStmt"
With IBtnNxtStmt
    If Not .Enabled Then
        'Msg CSub, "BoJmpNxtStmt is disabled"
        Exit Sub
    End If
    .Execute
End With
End Sub

Sub TileH():   IBtnTileH.Execute:                       End Sub
Sub MaxiImm(): IWinImm.WindowState = vbext_ws_Maximize: End Sub
Sub TileV():   IBtnTileV.Execute:                       End Sub

Sub CompilePC(): CompileP CPj: End Sub
Sub Compile(Pjn$)
JmpPj Pj(Pjn)
IBtnCompile.Execute
End Sub


Sub CompileP(P As VBProject)
JmpPj P
With IBtnCompile
    If .Enabled Then
        .Execute
        Debug.Print P.Name, "<--- Compiled"
    Else
        Debug.Print P.Name, "already Compiled"
    End If
End With
IBtnTileV.Execute
IBtnSav.Execute
End Sub

Sub CompileV(V As VBE): ItrDo V.VBProjects, "CompileP": End Sub

Sub ChkCompileBtnGood(Pjn$, Fun$)
Const CSub$ = CMod & "ChkCompileBtnGood"
Dim Act$, Ept$
Act = IBtnCompile.Caption
Ept = "Compi&le " & Pjn
If Act <> Ept Then Thw CSub, "Cur CompileBtn.Caption <> Compi&le {Pjn}", "Compile-Btn-Caption Pjn Ept-Btn-Caption", Act, Pjn, Ept
End Sub

Private Sub B_CompileP()
CompileP CPj
End Sub

Sub DltClr(A As CommandBar)
Dim I: For Each I In Itr(OyItr(A.Controls))
    CvIdeCtl(I).Delete
Next
End Sub

Sub DltIBar(IBarn$): IBarsC(IBarn).Delete: End Sub

Sub EnsIBtnSpy(SpyBtn$())
Dim I: For Each I In Itr(SpyBtn)
    EnsIBtnSp I
Next
End Sub

Sub EnsIBtnSp(SpIBtn)
Dim L$: L = SpIBtn
Dim Barn$: Barn = ShfTm(L)
EnsIBar Barn
EnsIBtn IBar(Barn), L
End Sub

Sub RmvIBarny(Barny$())
Dim Barn: For Each Barn In Barny
    If HasItn(CVbe.CommandBars, Barn) Then
        IBar(Barn).Delete
    End If
Next
End Sub
Sub EnsIBar(IBarn$)
If Not HasIBar(IBarn) Then Exit Sub
VisIBar IBarsC.Add(IBarn)
End Sub
Sub VisIBar(IBar As Office.CommandBar): IBar.Visible = True: End Sub
Sub EnsIBtnTml(Bar As CommandBar, TmlCapBtn$)
Dim CapBtn
For Each CapBtn In Tmy(TmlCapBtn)
    EnsIBtn Bar, CapBtn
Next
End Sub


Sub EnsIBtn(Bar As CommandBar, CapBtn)
If InIBarBtn(Bar, CapBtn) Then Exit Sub
Dim B As CommandBarButton
Set B = Bar.Controls.Add(MsoControlType.msoControlButton)
B.Caption = CapBtn
B.Style = msoButtonCaption
End Sub

Sub AddBtn(Bar As CommandBar, BtnCap)
Dim B As CommandBarButton
Set B = Bar.Controls.Add(MsoControlType.msoControlButton)
B.Caption = BtnCap
B.Style = msoButtonCaption
End Sub

Private Function SpSamp() As String()
'Cml:Sp::Sy#Specification:
Erase XX
X "Bars"
X " AA A1 A2 A3"
X " BB B1 B2 B3"
X "Btns"
X " A1"
SpSamp = XX  '*Spec
Erase XX
End Function

Function Btnny(SplnyBar$()) As String()
Stop 'Btnny = SrcInd(SpIBarc, "Bars")
End Function

Sub InstallIdeTools(ToolBarSpec$())
Stop 'EnsBtns BtnSpec(ToolBarSpec)
'EnsMdl Md("IdeTool"), ToolClsCd
End Sub

Function IBarnySp(SpITool$()) As String()
Stop ': IBarnySp = Tm1y(SpITool(SpITool)):      End Function
End Function
Sub RmvIBarSp(SpITool$()): RmvIBarny IBarnySp(SpITool): End Sub
