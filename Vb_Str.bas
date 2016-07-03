Attribute VB_Name = "Vb_Str"
Option Explicit
Public Const vbSngQ$ = "'"
Public Const vbDblQ$ = """"
Public Const vbSemiColon$ = ";"
Public Const vbComma$ = ","
Public Const vbCommaSpc$ = ", "

Type S1S2
    S1 As String
    S2 As String
End Type

Function SplitLvs(Lvs$) As String()
Dim A$
    A = Trim(Lvs)
While InStr(A, "  ") > 0
    A = Replace(A, "  ", " ")
Wend
SplitLvs = Split(A)
End Function

Function Str_BrkQuote(S) As S1S2
Dim O As S1S2, L%
L = Len(S)
If L = 1 Then
    O.S1 = S
    O.S2 = S
    Str_BrkQuote = O
    Exit Function
End If
If L = 2 Then
    O.S1 = Left(S, 1)
    O.S2 = Right(S, 1)
    Str_BrkQuote = O
    Exit Function
End If
Dim P%
P = InStr(S, "*")
If P > 0 Then
    O.S1 = Left(S, P - 1)
    O.S2 = Mid(S, P + 1)
    Str_BrkQuote = O
End If
End Function

Sub Str_Brw(S, Optional Nm$)
Dim Ft$
Ft = Tmp_Ft(Nm)
Ft_WrtStr Ft, S
Ft_Brw Ft
Kill Ft
End Sub

Function Str_Esc$(S)
If IsEmpty(S) Then Exit Function
Str_Esc = Replace(Replace(Replace(Replace(Replace(S, _
    "\", "\\"), _
    vbTab, "\t"), _
    ",", "\c"), _
    vbCr, "\r"), _
    vbLf, "\n")
End Function

Function Str_FirstDigitPos%(S)
Dim J%
For J = 1 To Len(S)
    If Str_IsDigit(Mid(S, J, 1)) Then
        Str_FirstDigitPos = J: Exit Function
    End If
Next
End Function

Function Str_IsDigit(S$) As Boolean
If Len(S) <> 1 Then Stop
Str_IsDigit = "0" <= S And S <= "9"
End Function

Function Str_IsSfx(S, Sfx$) As Boolean
Str_IsSfx = Right(S, Len(Sfx)) = Sfx
End Function

Function Str_Quote$(S, Q$)
With Str_BrkQuote(Q)
    Str_Quote = .S1 & S & .S2
End With
End Function

Function Str_ReplPfx$(S, Pfx$, ToPfx$)
If IsPfx(S, Pfx) Then
    Str_ReplPfx = ToPfx & Str_RmvPfx(S, Pfx)
Else
    Str_ReplPfx = S
End If
End Function

Function Str_ReplVBar$(S)
Str_ReplVBar = Replace(S, "|", vbCrLf)
End Function

Function Str_RmvEnd_LF$(S)
Dim O$, J&
O = S
For J = Len(O) To 1 Step -1
    If Right(O, 1) <> vbLf Then Str_RmvEnd_LF = O: Exit Function
    O = Left(O, Len(O) - 1)
Next
Str_RmvEnd_LF = O
End Function

Function Str_RmvPfx$(S, Pfx$)
If IsPfx(S, Pfx) Then
    Str_RmvPfx = Mid(S, Len(Pfx) + 1)
Else
    Str_RmvPfx = S
End If
End Function

Function Str_RmvSfx$(S, Sfx$)
If Str_IsSfx(S, Sfx) Then
    Str_RmvSfx = Left(S, Len(S) - Len(Sfx))
Else
    Str_RmvSfx = S
End If
End Function

Function Str_UnEsc$(S)
Str_UnEsc = Replace(Replace(Replace(Replace(Replace(S, _
    "\t", vbTab), _
    "\c", ","), _
    "\r", vbCr), _
    "\n", vbLf), _
    "\\", "\")
End Function

