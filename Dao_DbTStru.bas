Attribute VB_Name = "Dao_DbTStru"
Option Explicit
Private A_T$
Private A_D As Database
Private ZFldNmAy_Fk$()
Private ZFldNmAy_Sk$()
Private ZTbl As TableDef

Function Ay_MaxLen&(Ay)
Dim O&
Dim J&
For J = 0 To UB(Ay)
    If Len(Ay(J)) > O Then O = Len(Ay(J))
Next
Ay_MaxLen = O
End Function

Function Ay_Minus(Ay, ParamArray Ap())
Dim Av()
    Av = Ap
Dim J%
Dim O
    O = Ay
    For J = 0 To UB(Av)
        O = Ay_MinusOne(O, Av(J))
    Next
Ay_Minus = O
End Function

Function Ay_MinusOne(Ay, Ay1)
If Sz(Ay1) = 0 Then
    Ay_MinusOne = Ay
    Exit Function
End If
Dim O
O = Ay
Erase O
Dim J%
For J = 0 To UB(Ay)
    If Not Ay_Has(Ay1, Ay(J)) Then
        Push O, Ay(J)
    End If
Next
Ay_MinusOne = O
End Function

Function DbT_StruStr$(T As DbT)
A_T = T.T
Set A_D = T.D
Set ZTbl = ZZTbl
ZFldNmAy_Sk = ZZFldNmAy_Sk
ZFldNmAy_Fk = ZZFldNmAy_Fk
DbT_StruStr = T.T & " = " & ZPart_Pk & " | " & ZPart_Fk & " | " & ZPart_Rest & " | " & ZPart_UKey & " | " & ZPart_Key
End Function

Sub Tst()
Dim D As Database
    Set D = OpenDatabase("C:\temp\A.accdb")
Dim Act$
    'Act = DbT_StruStr(DbT(OpenDatabase(Fs_Fb), "SkuCostChr"))
    Act = DbT_StruStr(DbT(D, "Table1"))
MsgBox Act

Dim P As Prp
    P = Prp(ZZTbl, "Fk1", "Ele")
Debug.Print "Ele=" & Prp_Val(P)
Stop
End Sub

Private Property Get ZFldNmAy_All() As String()
ZFldNmAy_All = ZFlds_NmAy(ZTbl.Fields)
End Property

Private Property Get ZFldNmAy_Pk() As String()
Dim Idx As Index
For Each Idx In ZTbl.Indexes
    If Idx.Primary Then ZFldNmAy_Pk = ZIdx_FldNmAy(Idx)
Next
End Property

Private Property Get ZFldNmAy_Rest() As String()
ZFldNmAy_Rest = Ay_Minus(ZFldNmAy_All, Array(A_T), ZFldNmAy_Sk, ZFldNmAy_Fk)
End Property

Private Function ZFldNmAy_Str$(FldNmAy$())
Dim O$()
    O = FldNmAy
Dim J%
For J = 0 To UB(O)
    O(J) = Replace(O(J), A_T, "*")
Next
ZFldNmAy_Str = Join(O)
End Function

Private Function ZFld_IsFk(F As Field) As Boolean
Dim Ele$
    Ele = Prp_Val(Prp(ZTbl, F.Name, "Ele"))
ZFld_IsFk = Ele Like "Id*"
End Function

Private Function ZFlds_NmAy(F As Fields) As String()
Dim I As Field
Dim O$()
For Each I In F
    Push O, I.Name
Next
ZFlds_NmAy = O
End Function

Private Function ZIdx_FldNmAy(Idx As Index) As String()
Dim F As Field
Dim O$()
For Each F In Idx.Fields
    Push O, F.Name
Next
ZIdx_FldNmAy = O
End Function

Private Property Get ZKeyNmAy_Key() As String()
Dim O$()
    Dim Idx As Index
    For Each Idx In ZTbl.Indexes
        If Not Idx.Unique Then
            Push O, Idx.Name
        End If
    Next
ZKeyNmAy_Key = O
End Property

Private Property Get ZKeyNmAy_UKey() As String()
Dim O$()
    Dim Idx As Index
    For Each Idx In ZTbl.Indexes
        If Idx.Unique Then
            If Idx.Name <> A_T Then
                If Not Idx.Primary Then
                    Push O, Idx.Name
                End If
            End If
        End If
    Next
ZKeyNmAy_UKey = O
End Property

Private Function ZKeyStr_ByKeyNmAy$(KeyNmAy$())
Dim A$()
    A = KeyNmAy
Dim N%
    N = Sz(A)
If N = 0 Then Exit Function
Dim O$()
    ReDim O(N - 1)
    Dim J%
    For J = 0 To N - 1
        O(J) = ZKeyStr_OneKey(A(J))
    Next
ZKeyStr_ByKeyNmAy = Join(O, ", ")
End Function

Private Function ZKeyStr_OneKey$(UKeyNm$)
Dim F$()
    F = ZKey_FldNmAy(UKeyNm)
ZKeyStr_OneKey = UKeyNm & "(" & Join(F) & ")"
End Function

Private Function ZKey_FldNmAy(KeyNm$) As String()
ZKey_FldNmAy = ZIdx_FldNmAy(ZTbl.Indexes(KeyNm))
End Function

Private Property Get ZPart_Fk$()
ZPart_Fk = ZFldNmAy_Str(ZFldNmAy_Fk)
End Property

Private Property Get ZPart_Key$()
ZPart_Key = ZKeyStr_ByKeyNmAy(ZKeyNmAy_Key)
End Property

Private Property Get ZPart_Pk$()
Dim A$()
    A = ZFldNmAy_Pk
If Sz(A) <> 1 Then Er "{Db}.{T} has {n}-Pk fields, which should be 1", A_D.Name, A_T, Sz(A)
If A(0) <> A_T Then Er "{Db}.{T} has 1-Pk-field-{Name}, which should be same as [T]", A_D.Name, A_T, A(0)
ZPart_Pk = "* " & ZFldNmAy_Str(ZFldNmAy_Sk)
End Property

Private Property Get ZPart_Rest$()
ZPart_Rest = ZFldNmAy_Str(ZFldNmAy_Rest)
End Property

Private Property Get ZPart_UKey$()
ZPart_UKey = ZKeyStr_ByKeyNmAy(ZKeyNmAy_UKey)
End Property

Private Property Get ZZFldNmAy_Fk() As String()
Dim O$()
Dim F As Field
For Each F In ZTbl.Fields
    If ZFld_IsFk(F) Then Push O, F.Name
Next
ZZFldNmAy_Fk = O
End Property

Private Property Get ZZFldNmAy_Sk() As String()
Dim Idx As Index
For Each Idx In ZTbl.Indexes
    If Idx.Name = A_T Then
        If Not Idx.Unique Then Er "{Db}.{T} has Index of same name as [T] not unique (which is same-name-key should be unique)", A_D.Name, A_T
        ZZFldNmAy_Sk = ZIdx_FldNmAy(Idx)
        Exit Property
    End If
Next
End Property

Private Property Get ZZTbl() As TableDef
Set ZZTbl = A_D.TableDefs(A_T)
End Property
