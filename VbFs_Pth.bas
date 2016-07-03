Attribute VB_Name = "VbFs_Pth"
Option Explicit
Public Fso As New FileSystemObject

Sub Pth_AssertExist(Pth)
If Not Pth_IsExist(Pth) Then Err.Raise 1, , "Folder[" & Pth & "] not exist"
End Sub

Sub Pth_CrtEachSeg(Pth$)
Dim A$(), iPth$, J%
A = Split(Pth, "\")
iPth = A(0) & "\"
For J = 1 To UB(A) - 1
    iPth = iPth & A(J) & "\"
    If Dir(iPth, vbDirectory) = "" Then MkDir iPth
Next
End Sub

Sub Pth_CrtIfNotExist(Pth$)
If Not Pth_IsExist(Pth) Then MkDir Pth
End Sub

Sub Pth_DltAllFil(Pth$)
Dim A$(), J%
A = Pth_FfnAy(Pth)
For J = 0 To UB(A)
    Kill A(J)
Next
End Sub

Function Pth_DirAy(Pth$) As String()
Dim A$(), O$(), J%
A = Pth_EntriesAy("*.*", vbDirectory)
For J = 0 To UB(A)
    If Fso.FolderExists(Pth & A(J)) Then
        Push O, A(J)
    End If
Next
Pth_DirAy = O
End Function

Function Pth_FfnAy(Pth$, Optional FSpec$ = "*.*") As String()
Pth_FfnAy = Ay_AddPfx(Pth_FnAy(Pth, FSpec), Pth)
End Function

Function Pth_FnAy(Pth$, Optional FSpec$ = "*.*") As String()
Pth_FnAy = Pth_EntriesAy(Pth, FSpec)
End Function

Function Pth_FtAy() As String()
Pth_FtAy = Pth_FfnAy("*.txt")
End Function

Function Pth_IsExist(Pth) As Boolean
Pth_IsExist = Fso.FolderExists(Pth)
End Function

Function Pth_Normalize$(Pth$)
Dim J%, Fnd As Boolean
Pth = Repl2X_To1(Pth, "\")
If ChrEnd(Pth) <> "\" Then Pth = Pth & "\"
If InStr(Pth, "..") > 0 Then
    Dim Ay$(), P&
    Ay = Split(Pth, "\")
    Do
        Fnd = False
        For J = 0 To UB(Ay)
            If Ay(J) = ".." Then
                P = J
                Fnd = True
                Exit For
            End If
        Next
        If Not Fnd Then
            Pth = Join(Ay, "\")
            Exit Do
        End If
        If P = 0 Or P = 1 Then Err.Raise 1
        Ay = Ay_RmvEleAt(Ay_RmvEleAt(Ay, P), P - 1)
    Loop
End If
Pth_Normalize = Pth
End Function

Sub Pth_Opn(Pth)
Pth_AssertExist Pth
Shell Fmt_QQ("Explorer ""?""", Pth), vbNormalFocus
End Sub

Function Pth_PthAy(Pth$) As String()
Pth_PthAy = Ay_AddSfx(Pth_DirAy(Pth), "\")
End Function

Private Function Pth_EntriesAy(Pth$, Spec$, Optional Attr As VbFileAttribute = vbNormal) As String()
Dim A$, O$()
If Not Pth_IsExist(Pth) Then Exit Function
A = Dir(Pth & Spec, Attr)
While A <> ""
    If A <> "." And A <> ".." Then
        Push O, A
    End If
    A = Dir
Wend
Pth_EntriesAy = O
End Function
