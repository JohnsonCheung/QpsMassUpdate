Attribute VB_Name = "Vb"
Option Explicit

Function AddPfx$(S$, Pfx)
If IsPfx(S, Pfx) Then
    AddPfx = S
Else
    AddPfx = Pfx & S
End If
End Function

Function AddSfx$(S$, Sfx)
If IsSfx(S, Sfx) Then
    AddSfx = S
Else
    AddSfx = S & Sfx
End If
End Function

Sub Assert_AyIsEmpty(ErAy$())
If Sz(ErAy) = 0 Then Exit Sub
Ay_Brw ErAy, "ErAy-Should-Be-Empty"
Er "Error"
End Sub

Function Brk1(S, Brk$) As S1S2
Dim P&, O As S1S2
P = InStr(S, Brk)
If P = 0 Then
    O.S1 = Trim(S)
Else
    O.S1 = Trim(Left(S, P - 1))
    O.S2 = Trim(Mid(S, P + Len(Brk)))
End If
Brk1 = O
End Function

Function Brk2(S, Brk$) As S1S2
Dim P&, O As S1S2
P = InStr(S, Brk)
If P = 0 Then
    O.S2 = Trim(S)
Else
    O.S1 = Trim(Left(S, P - 1))
    O.S2 = Trim(Mid(S, P + Len(Brk)))
End If
Brk2 = O
End Function

Function Brk(S, BrkStr$) As S1S2
Dim P&, O As S1S2
P = InStr(S, BrkStr)
If P = 0 Then Err.Raise 1
O.S1 = Trim(Left(S, P - 1))
O.S2 = Trim(Mid(S, P + Len(BrkStr)))
Brk = O
End Function

Function BrkQuote(Q) As S1S2
Dim O As S1S2, P%
Select Case Len(Q)
Case 1: O.S1 = Q: O.S2 = Q
Case 2: O.S1 = Left(Q, 1): O.S2 = Right(Q, 1)
Case Is > 2:
    P = InStr(Q, "*")
    If P = 0 Then Err.Raise 1, , "Invalid Q[" & Q & "] to BrkQuote"
    O.S1 = Left(Q, P - 1)
    O.S2 = Mid(Q, P + 1)
End Select
BrkQuote = O
End Function

Function Camel(S) As String()
Dim A$, O$(), J%
A = S
While A <> ""
    For J = 2 To Len(A)
        If IsUCase(Mid(A, J, 1)) Then
            Push O, Left(A, J - 1)
            A = Mid(A, J)
            GoTo Nxt
        End If
    Next
    Push O, A
    A = ""
Nxt:
Wend
Camel = O
End Function

Function ChrBeg$(S$)
ChrBeg = Left(S, 1)
End Function

Function ChrEnd$(S$)
ChrEnd = Right(S, 1)
End Function

Function CutExt$(Ffn)
CutExt = Ffn_CutExt(Ffn)
End Function

Function DblUL_Pfx$(S$)
Dim P%
P = InStr(S, "__")
If P = 0 Then Exit Function
DblUL_Pfx = Left(S, P - 1)
End Function

Function DblUL_Sfx$(S$)
Dim P%
P = InStrRev(S, "__")
If P = 0 Then Exit Function
DblUL_Sfx = Mid(S, P + 2)
End Function

Function Dft(A, B)
If IsObject(A) Then
    If IsNothing(A) Then
        Set Dft = B
    Else
        Set Dft = A
    End If
    Exit Function
End If
If Trim(A) = "" Then
    Dft = B
Else
    Dft = A
End If
End Function

Function DotOrNm$(Bool As Boolean, NmForTrue$)
DotOrNm = IIf(Bool, NmForTrue, ".")
End Function

Sub Er(ErMsg$, ParamArray Ap())
Dim Av()
    Av = Ap
Dim A$
A = Fmt_ErDes(ErMsg, Av)
Str_Brw A
Err.Raise 1, , A
End Sub

Sub Er_ByAv(ErMsg$, Av())
Dim A$
A = Fmt_ErDes(ErMsg, Av)
Str_Brw A
Err.Raise 1, , A
End Sub

Function IfGivenStr$(S, UseThis_If_S_IsGiven$)
If IsMissing(S) Then Exit Function
If IsEmpty(S) Then Exit Function
If IsNothing(S) Then Exit Function

If S = "" Then Exit Function
IfGivenStr = UseThis_If_S_IsGiven
End Function

Function IsBackupNm(Nm$) As Boolean
Dim A$, B$
A = Right(Nm, 4)
If Left(A, 2) <> "__" Then Exit Function
B = Right(A, 2)
If Format(Val(B), "00") <> B Then Exit Function
IsBackupNm = True
End Function

Function IsBareName(Nm$) As Boolean
Dim J%
For J = 1 To Len(Nm)
    If Not IsNmChr(Mid(Nm, J, 1)) Then Exit Function
Next
IsBareName = True
End Function

Function IsBlank(S) As Boolean
IsBlank = Trim(S) = ""
End Function

Function IsDigit(Nm$) As Boolean
Dim A$
A = Left(Nm, 1)
If "0" <= A And A <= "9" Then IsDigit = True
End Function

Function IsLetter(S$) As Boolean
IsLetter = True
Dim A$
A = Left(S, 1)
If "a" <= A And A <= "z" Then Exit Function
If "A" <= A And A <= "Z" Then Exit Function
IsLetter = False
End Function

Function IsLikAy(S, Ay$()) As Boolean
'Aim: If p is like any one of the element of pAylik$()
Dim J&
For J = 0 To UB(Ay)
    If S Like Ay(J) Then IsLikAy = True: Exit Function
Next
End Function

Function IsNmChr(S$) As Boolean
IsNmChr = True
If IsDigit(S) Then Exit Function
If IsLetter(S) Then Exit Function
If IsUL(S) Then Exit Function
IsNmChr = False
End Function

Function IsNothing(V) As Boolean
IsNothing = Var_IsNothing(V)
End Function

Function IsPfx(S, Pfx, Optional IgnoreCase As Boolean) As Boolean
If IgnoreCase Then
    IsPfx = UCase(Left(S, Len(Pfx))) = UCase(Pfx)
Else
    IsPfx = Left(S, Len(Pfx)) = Pfx
End If
End Function

Function IsPfxAy(S$, PfxAy) As Boolean
'Return true if S has one the pfx in PfxAy
Dim Pfx
For Each Pfx In PfxAy
    If IsPfx(S, Pfx) Then IsPfxAy = True: Exit Function
Next
End Function

Function IsPfxInDict(S$, D As Dictionary) As Boolean
IsPfxInDict = IsPfxAy(S, D.Keys)
End Function

Function IsSfx(S, Sfx) As Boolean
IsSfx = Right(S, Len(Sfx)) = Sfx
End Function

Function IsUCase(A$) As Boolean
Dim B%
B = Asc(A)
If 65 <= B And B <= 90 Then IsUCase = True
End Function

Function IsUL(S$) As Boolean
IsUL = Left(S, 1) = "_"
End Function

Function Max(A, B)
Max = IIf(A > B, A, B)
End Function

Sub Msg(Msg$, ParamArray Ap())
Dim Av()
    Av = Ap
Str_Brw Fmt_ErDes(Msg, Av)
End Sub

Function PfxInAy$(S$, Ay)
'Return the pfx in S if S has such pfx in Ay
Dim J&
For J = 0 To UB(Ay)
    If IsPfx(S, Ay(J)) Then
        PfxInAy = Ay(J)
        Exit Function
    End If
Next
End Function

Function PfxInDict$(S$, D As Dictionary)
PfxInDict = PfxInAy(S, D.Keys)
End Function

Function Quote$(S, Q$)
Dim A As S1S2
A = BrkQuote(Q)
Quote = A.S1 & S & A.S2
End Function

Function Repl1DblQ_To2$(S)
Repl1DblQ_To2 = Repl1X_To2(S, vbDblQ)
End Function

Function Repl1X_To2$(S, X$)
Dim P&, XX$, O$
P = 1
XX = X & X
Do
    P = InStr(P, S, X)
    If P = 0 Then
        Repl1X_To2 = S
        Exit Function
    End If
    S = ReplOnce(S, X, XX, P)
    P = P + Len(XX)
Loop
End Function

Function Repl2DblQ_To1$(S$)
Repl2DblQ_To1 = Repl2X_To1(S, vbDblQ)
End Function

Function Repl2Spc_To1$(S$)
Repl2Spc_To1 = Repl2X_To1(S, " ")
End Function

Function Repl2X_To1$(S, X$)
Dim TwoX$
TwoX = X & X
While InStr(S, TwoX) > 0
    S = Replace(S, TwoX, X)
Wend
Repl2X_To1 = S
End Function

Function ReplExt$(Ffn, NewExt$)
ReplExt = Ffn_ReplExt(Ffn, NewExt)
End Function

Function ReplOnce$(S, Find$, By$, Optional StartPos = 1)
Dim P%
P = InStr(StartPos, S, Find)
If P = 0 Then
    ReplOnce = S
Else
    ReplOnce = Left(S, P - 1) & By & Mid(S, P + Len(Find))
End If
End Function

Function ReplPth$(Ffn$, RelativePth$)
ReplPth = Ffn_ReplPth(Ffn, RelativePth)
End Function

Function ReplQMrk$(S, C)
ReplQMrk = Replace(S, "?", C)
End Function

Function ReplVBar$(S$)
ReplVBar = Replace(S, "|", vbCrLf)
End Function

Function RmvChrBeg$(S$)
RmvChrBeg = Mid(S, 2)
End Function

Function RmvChrEnd$(S$)
RmvChrEnd = Left(S, Len(S) - 1)
End Function

Function RmvChrEnd_OfAnyInList$(S$, EndChrList$)
Dim A$
A = ChrEnd(S)
If InStr(EndChrList, A) > 0 Then
    RmvChrEnd_OfAnyInList = RmvChrEnd(S)
Else
    RmvChrEnd_OfAnyInList = S
End If
End Function

Function RmvLastNChr(S, N)
RmvLastNChr = Left(S, Len(S) - N)
End Function

Function RmvPfx$(S$, Pfx)
If IsPfx(S, Pfx) Then
    RmvPfx = Mid(S, Len(Pfx) + 1)
Else
    RmvPfx = S
End If
End Function

Function RmvPfxAy$(S$, PfxAy$())
Dim Pfx
For Each Pfx In PfxAy
    If IsPfx(S, Pfx) Then RmvPfxAy = RmvPfx(S, Pfx): Exit Function
Next
RmvPfxAy = S
End Function

Function RmvPfx_OfAnyDblUL$(S$)
RmvPfx_OfAnyDblUL = TakAftX(S, "__", ReturnS_IfNoX:=True)
End Function

Function RmvSfx$(S$, Sfx)
Dim L%
L = Len(Sfx)
If Right(S, L) = Sfx Then
    RmvSfx = Left(S, Len(S) - L)
Else
    RmvSfx = S
End If
End Function

Function RmvSfx_OfFirstDblUL$(S$)
RmvSfx_OfFirstDblUL = TakBefX(S, "__", ReturnS_IfNoX:=True)
End Function

Function RmvSfx_OfLastDblUL$(S$)
RmvSfx_OfLastDblUL = TakBefLastX(S, "__", ReturnS_IfNoX:=True)
End Function

Function TakAftLastX$(S, X$, Optional AlsoTakX As Boolean = False, Optional ReturnS_IfNoX As Boolean)
Dim P%, L%
P = InStrRev(S, X)
If P = 0 Then
    If ReturnS_IfNoX Then TakAftLastX = S
    Exit Function
End If
If Not AlsoTakX Then L = Len(X)
TakAftLastX = Mid(S, P + L)
End Function

Function TakAftX$(S, X$, Optional AlsoTakX As Boolean = False, Optional ReturnS_IfNoX As Boolean)
Dim P%, L%
P = InStr(S, X)
If P = 0 Then
    If ReturnS_IfNoX Then TakAftX = S
    Exit Function
End If
If Not AlsoTakX Then L = Len(X)
TakAftX = Mid(S, P + L)
End Function

Function TakBefLastX$(S, X$, Optional AlsoTakX As Boolean = False, Optional ReturnS_IfNoX As Boolean)
Dim P%, L%
P = InStrRev(S, X)
If P = 0 Then
    If ReturnS_IfNoX Then TakBefLastX = S
    Exit Function
End If
If AlsoTakX Then L = Len(X)
TakBefLastX = Left(S, P + L - 1)
End Function

Function TakBefX$(S, X$, Optional AlsoTakX As Boolean = False, Optional ReturnS_IfNoX As Boolean)
Dim P%, L%
P = InStr(S, X)
If P = 0 Then
    If ReturnS_IfNoX Then TakBefX = S
    Exit Function
End If
If AlsoTakX Then L = Len(X)
TakBefX$ = Left(S, P + L - 1)
End Function

Function TimStmp$()
Static Cnt%
TimStmp = Format(Now, "yyyy-mm-dd HHMMSS") & "-" & Cnt
Cnt = Cnt + 1
End Function
