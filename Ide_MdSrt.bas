Attribute VB_Name = "Ide_MdSrt"
Option Explicit
Private Md As CodeModule
Dim Mth_Modifier$()
Dim Mth_Nm$()
Dim Mth_PrpTy$()
Dim Mth_BdyAy() As Variant
Sub AA()
CurMd_Srt
End Sub
Sub CurMd_Srt()
Md_Srt CurMd
End Sub

Sub CurPj_Srt()
Pj_Srt CurPj
End Sub

Sub Md_Srt(P As CodeModule)
Dim A$
If P.CountOfLines = 0 Then Exit Sub
Set Md = P
Mth_Ay
Md.DeleteLines Dlt_Beg, Dlt_Cnt
A = Mth_Sorted
If A = "" Then Exit Sub
Md.InsertLines Md.CountOfLines + 1, A
End Sub

Sub Pj_Srt(P As VBProject)
Dim I, Cmp As VBComponent, O$()
For Each I In P.VBComponents
    Set Cmp = I
    Select Case Cmp.Type
    Case _
        vbext_ComponentType.vbext_ct_ClassModule, _
        vbext_ComponentType.vbext_ct_StdModule, _
        vbext_ComponentType.vbext_ct_Document
        Debug.Print Cmp.Name
        Md_Srt Cmp.CodeModule
        Push O, Cmp.Name
    End Select
Next
Ay_Brw O
End Sub

Private Property Get Dlt_Beg&()
Dlt_Beg = Md.CountOfDeclarationLines + 1
End Property

Private Property Get Dlt_Cnt&()
Dlt_Cnt = Md.CountOfLines - Dlt_Beg + 1
End Property

Private Sub Md_Srt__Tst()
Md_Srt Md_ByNm("Ide_MdSrt")
End Sub

Private Sub Mth_Ay()
Dim IsMthLin As Boolean
Dim Lno_Nm$
Dim Lno_Modifier$
Dim Lno_Bdy$()
Dim Lno_PrpTy$
Dim FnTy$
Dim L$
Dim Lno&
Dim LinAy$()
Dim A$, P%
Erase Mth_Modifier
Erase Mth_Nm
Erase Mth_BdyAy
Erase Mth_PrpTy

LinAy = Split(Md.Lines(1, Md.CountOfLines), vbCrLf)
Lno = Md.CountOfDeclarationLines + 1
Do
    If Lno >= Md.CountOfLines Then Exit Do
    L = LinAy(Lno - 1)
    If IsPfx(L, "Private ") Then
        Lno_Modifier = "Private"
        L = RmvPfx(L, "Private ")
    ElseIf IsPfx(L, "Public ") Then
        Lno_Modifier = "Public"
        L = RmvPfx(L, "Public ")
    ElseIf IsPfx(L, "Friend ") Then
        Lno_Modifier = "Friend"
        L = RmvPfx(L, "Friend ")
    Else
        Lno_Modifier = "Public"
    End If
    
    Lno_PrpTy = ""
    If IsPfx(L, "Sub ") Then
        L = RmvPfx(L, "Sub ")
        IsMthLin = True
        FnTy = "Sub"
    ElseIf IsPfx(L, "Function ") Then
        L = RmvPfx(L, "Function ")
        IsMthLin = True
        FnTy = "Function"
    ElseIf IsPfx(L, "Property ") Then
        L = RmvPfx(L, "Property ")
        If IsPfx(L, "Get") Then
            L = RmvPfx(L, "Get ")
            Lno_PrpTy = "Get"
        ElseIf IsPfx(L, "Let") Then
            L = RmvPfx(L, "Let ")
            Lno_PrpTy = "Let"
        ElseIf IsPfx(L, "Set") Then
            L = RmvPfx(L, "Set ")
            Lno_PrpTy = "Set"
        Else
            Er "AA"
        End If
        IsMthLin = True
        FnTy = "Property"
    Else
        IsMthLin = False
    End If

    A = "End " & FnTy
    If IsMthLin Then
        P = InStr(L, "(")
        If P = 0 Then Er "No [(] of {MthLin}", LinAy(Lno - 1)
        Lno_Nm = Left(L, P - 1)
        Select Case Right(Lno_Nm, 1)
        Case "%", "#", "!", "@", "^", "&": Lno_Nm = Left(Lno_Nm, Len(Lno_Nm) - 1)
        End Select
        Push Mth_Modifier, Lno_Modifier '<===
        Push Mth_Nm, Lno_Nm             '<===
        Push Mth_PrpTy, Lno_PrpTy       '<===
        Erase Lno_Bdy
        For Lno = Lno To Md.CountOfLines
            L = LinAy(Lno - 1)
            Push Lno_Bdy, L
            If L = A Then
                Push Mth_BdyAy, Lno_Bdy '<===
                GoTo Nxt
            End If
        Next
        Er "AA"
    End If
Nxt:
    Lno = Lno + 1
Loop
End Sub

Private Function Mth_Key() As String()
Dim O$(), A$(), B$(), C$(), U&, J&, D As Byte
A = Mth_Nm
B = Mth_Modifier
C = Mth_PrpTy

U = UB(A)
If U = -1 Then Exit Function
ReDim O(U)
For J = 0 To U
    Select Case B(J)
    Case "Public": D = 0
    Case "Friend": D = 1
    Case "Private": D = 2
    Case Else
        Er "{J} {B(J)}", J, B(J)
    End Select
    O(J) = D & ":" & A(J) & ":" & C(J)
Next
Mth_Key = O
End Function

Private Function Mth_Sorted$()
Dim I&()
Dim J&
Dim A()
Dim O$()
A = Mth_BdyAy
I = Mth_SortedIdx
For J = 0 To UB(I)
    Push O, ""
    PushAy O, A(I(J))
Next
Mth_Sorted = Join(O, vbCrLf)
End Function

Private Function Mth_SortedIdx() As Long()
Mth_SortedIdx = Ay_Srt_IntoIdxAy(Mth_Key)
End Function

