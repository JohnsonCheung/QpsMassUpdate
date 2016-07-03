Attribute VB_Name = "Vb_Ay"
Option Explicit
Option Compare Text

Function Ay_AddPfx(Ay, Pfx) As String()
Dim N&, J&
N = Sz(Ay)
If N = 0 Then Exit Function
ReDim O$(N - 1)
For J = 0 To N - 1
    O(J) = Pfx & Ay(J)
Next
Ay_AddPfx = O
End Function

Function Ay_AddSfx(Ay, Sfx) As String()
Dim N&, J&
N = Sz(Ay)
If N = 0 Then Exit Function
ReDim O$(N - 1)
For J = 0 To N - 1
    O(J) = Ay(J) & Sfx
Next
Ay_AddSfx = O
End Function

Function Ay_Distinct(Ay)
Dim O
    O = Ay
    Erase O
    Dim I&
    For I = 0 To UB(Ay)
        Push_NoDup O, Ay(I)
    Next
Ay_Distinct = O
End Function

Sub Ay_Dmp(Ay)
Dim J&
For J = LB(Ay) To UB(Ay)
    Debug.Print Ay(J)
Next
End Sub

Function Ay_HSqv(Ay)
Dim O, J&
ReDim O(1 To 1, 1 To Sz(Ay))
For J = 0 To UBound(Ay)
    O(1, J + 1) = Ay(J)
Next
Ay_HSqv = O
End Function

Function Ay_Has(Ay, Itm) As Boolean
Dim J&
For J = 0 To UB(Ay)
    If Ay(J) = Itm Then Ay_Has = True: Exit Function
Next
End Function

Function Ay_Idx&(Ay, Itm)
Dim J&
For J = 0 To UB(Ay)
    If Ay(J) = Itm Then Ay_Idx = J: Exit Function
Next
Ay_Idx = -1
End Function

Function Ay_IdxAy(FullAy, SubAy) As Long()
Dim U&
    U = UB(SubAy)
If U = -1 Then Exit Function
Dim O&()
    ReDim O(U)
    Dim J&
    For J = 0 To U
        O(J) = Ay_Idx(FullAy, SubAy(J))
    Next
Ay_IdxAy = O
End Function

Function Ay_IdxAy_OfInt(FullAy, SubAy) As Integer()
Dim U%
    U = UB(SubAy)
If U = -1 Then Exit Function
Dim O%()
    ReDim O(U)
    Dim J&
    For J = 0 To U
        O(J) = Ay_Idx(FullAy, SubAy(J))
    Next
Ay_IdxAy_OfInt = O
End Function

Function Ay_IsEmpty(Ay) As Boolean
Ay_IsEmpty = Sz(Ay) = 0
End Function

Function Ay_LastEle(Ay)
Dim U&
    U = UB(Ay)
If U = -1 Then Exit Function
If IsObject(Ay(U)) Then
    Set Ay_LastEle = Ay(U)
Else
    Ay_LastEle = Ay(U)
End If
End Function

Function Ay_Quote(Ay, Q$) As String()
Dim O$(), U&
U = UB(Ay)
If U = -1 Then Exit Function
ReDim O(U)
Dim J&
For J = 0 To U
    O(J) = Str_Quote(Ay(J), Q)
Next
Ay_Quote = O
End Function

Function Ay_RmvBlankEle(Ay) As String()
Dim N&
    N = Sz(Ay)
If N = 0 Then Exit Function
Dim O$()
    Dim J&
    For J = 0 To N - 1
        If Trim(Ay(J)) <> "" Then Push O, Ay(J)
    Next
Ay_RmvBlankEle = O
End Function

Function Ay_RmvEleAt(Ay, At&)
Dim O, J&, U&, Rmv As Boolean
O = Ay
U = UB(O)
For J = At + 1 To U
    O(J - 1) = O(J)
    Rmv = True
Next
If Rmv Then ReDim Preserve O(U - 1)
Ay_RmvEleAt = O
End Function

Function Ay_Sqv(Ay)
Dim N&
N = Sz(Ay)
ReDim O(1 To 1, 1 To N)
Dim J&
For J = 0 To N - 1
    O(1, J + 1) = Ay(J)
Next
Ay_Sqv = O
End Function

Function Ay_StrEsc(Ay) As String()
Dim U&
    U = UB(Ay)
If U = -1 Then Exit Function
Dim O$()
    ReDim O(U)
Dim J&
For J = 0 To U
    O(U) = Str_Esc(Ay(J))
Next
Ay_StrEsc = O
End Function

Function Ay_SubSet_ByPfx(Ay, Pfx$) As String()
Dim O$(), J&
For J = 0 To UB(Ay)
    If IsPfx(Ay(J), Pfx) Then Push O, Ay(J)
Next
Ay_SubSet_ByPfx = O
End Function

Function Ay_ToStr$(Ay)
Ay_ToStr = Join(Ay_Quote(Ay, "[]"))
End Function

Function LB&(Ay)
On Error GoTo X
LB = LBound(Ay)
Exit Function
X: LB = -2
End Function

Sub Push(Ay, Itm)
Dim N&
N = Sz(Ay)
ReDim Preserve Ay(N)
If IsObject(Itm) Then
    Set Ay(N) = Itm
Else
    Ay(N) = Itm
End If
End Sub

Sub PushAy(Ay, Ay1)
Dim J&
For J = 0 To UB(Ay1)
    Push Ay, Ay1(J)
Next
End Sub

Sub PushAy_NoDup(O, Ay)
Dim J&
For J = 0 To UB(Ay)
    Push_NoDup O, Ay(J)
Next
End Sub

Sub Push_NoDup(Ay, Itm)
If Not Ay_Has(Ay, Itm) Then
    Push Ay, Itm: Exit Sub
End If
End Sub

Function StrAy(ParamArray Ap()) As String()
Dim Av()
    Av = Ap
Dim O$()
    ReDim O(UBound(Av))
Dim J%
For J = 0 To UBound(Av)
    O(J) = Av(J)
Next
StrAy = O
End Function

Function Sz&(Ay)
If IsEmpty(Ay) Then Exit Function
If Not IsArray(Ay) Then Er "{Type} of Ay is not array", TypeName(Ay)
On Error Resume Next
Sz = UBound(Ay) + 1
End Function

Function UB&(Ay)
UB = Sz(Ay) - 1
End Function
