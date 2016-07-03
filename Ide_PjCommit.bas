Attribute VB_Name = "Ide_PjCommit"
Option Explicit

Sub CurPj_BrwSrcFdr()
Pj_BrwSrcFdr CurPj
End Sub

Sub CurPj_Commit()
Pj_Commit CurPj
End Sub

Sub Pj_BrwSrcFdr(Pj As VBProject)
Pth_Opn Pj_SrcPth(Pj)
End Sub


Sub Pj_Commit(Pj As VBProject, Optional Msg$ = "Commit")
Dim Cmp As VBComponent
Dim PjExt$
Dim CmpAy() As VBComponent
Dim FnAy$()
Dim SrcPth$
Dim BatFn$, BatAy$()
Dim ToFfn$
Dim J%

For Each Cmp In Pj.VBComponents
    If CmpTy_IsWithMd(Cmp.Type) Then
        Push CmpAy, Cmp
        Push FnAy, Cmp.Name & CmpTy_Ext(Cmp.Type)
    End If
Next
SrcPth = Pj_SrcPth(Pj)
If Pth_IsExist(SrcPth) Then
    Pth_DltAllFil SrcPth
Else
    Pth_CrtEachSeg SrcPth
End If

ToFfn = Ffn_ReplPth(Pj.Filename, SrcPth) '<--
Fso.CopyFile Pj.Filename, ToFfn, OverwriteFiles:=True  '<===

ChDir SrcPth    '<=== ChDir
For J = 0 To UB(FnAy)
    If CmpAy(J).CodeModule.CountOfLines > 1 Then
        CmpAy(J).Export FnAy(J)     '<=====
    End If
Next

If Not Pth_IsExist(SrcPth & ".git\") Then
    Stop
    Shell "git init"
End If

Push BatAy, "git add *.*"
Push BatAy, Fmt_QQ("git commit -m ""?""", Msg)
BatFn = SrcPth & "Commit.bat"
Ay_Wrt BatAy, BatFn
Shell BatFn, vbHide
End Sub

Sub Ay_Wrt(Ay, Ffn$)
Dim F%, J&
F = FreeFile(1)
Open Ffn For Output As F
For J = 0 To UB(Ay)
    Print #F, Ay(J)
Next
Close F
End Sub
