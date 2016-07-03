Attribute VB_Name = "Dao_DbTStru"
Option Explicit

Function DbT_StruStr$(T As DbT)
Dim Tbl As Dao.TableDef
Dim TblNm$
Dim Idx As Index
Dim F As Field
Dim Ele$
Dim KeyNmAy_Key$()
Dim KeyNmAy_UKey$()
Dim FldNmAy_Sk$()
Dim FldNmAy_Fk$()
Dim FldNmAy_Rest$()
Dim FldNmAy_All$()
Dim FldNmAy_Pk$()
Dim Part_Key$
Dim Part_Pk$
Dim Part_UKey$
Dim Part_Rest$
Dim Part_Fk$
Dim Part$

TblNm = T.T
Set Tbl = T.D.TableDefs(TblNm)
Erase FldNmAy_Pk
    For Each Idx In Tbl.Indexes
        If Idx.Primary Then FldNmAy_Pk = ZIdx_FldNmAy(Idx): Exit For
    Next
    If Sz(FldNmAy_Pk) <> 1 Then Er "{Db}.{T} has {n}-Pk fields, which should be 1", T.D.Name, TblNm, Sz(FldNmAy_Pk)
    If FldNmAy_Pk(0) <> TblNm Then Er "{Db}.{T} has 1-Pk-field-{Name}, which should be same as [T]", T.D.Name, TblNm, FldNmAy_Pk(0)

Erase FldNmAy_Sk
    For Each Idx In Tbl.Indexes
        If Idx.Name = TblNm Then
            If Not Idx.Unique Then Er "{Db}.{T} has Index of same name as [T] not unique (which is same-name-key should be unique)", T.D.Name, TblNm
            FldNmAy_Sk = ZIdx_FldNmAy(Idx) '<===
            Exit For
        End If
    Next

Erase FldNmAy_Fk
    For Each F In Tbl.Fields
        Ele = Prp_Val(Prp(Tbl, F.Name, "Ele"))
        If Ele Like "Id*" Then Push FldNmAy_Fk, F.Name    '<==
    Next
    
Erase FldNmAy_All
    For Each F In Tbl.Fields
        Push FldNmAy_All, F.Name    '<===
    Next

FldNmAy_Rest = Ay_Minus(FldNmAy_All, Array(TblNm), FldNmAy_Sk, FldNmAy_Fk)

Erase KeyNmAy_Key
    For Each Idx In Tbl.Indexes
        If Not Idx.Unique Then
            Push KeyNmAy_Key, Idx.Name  '<===
        End If
    Next

Erase KeyNmAy_UKey
    For Each Idx In Tbl.Indexes
        If Idx.Unique Then
            If Idx.Name <> TblNm Then
                If Not Idx.Primary Then
                    Push KeyNmAy_UKey, Idx.Name '<===
                End If
            End If
        End If
    Next

Part_Fk = Join(FldNmAy_Fk)
Part_Pk = "* " & Join(FldNmAy_Sk)
Part_Rest = Join(FldNmAy_Rest)
Part_UKey = ZKeyStr(Tbl, KeyNmAy_UKey)
Part_Key = ZKeyStr(Tbl, KeyNmAy_Key)

Part = Part_Pk & " | " & Part_Fk & " | " & Part_Rest & " | " & Part_UKey & " | " & Part_Key
Part = Replace(Part, TblNm, "*")
DbT_StruStr = TblNm & " = " & Part
End Function

Function Db_StruStr$(Db As Database)
Dim O$(), T As TableDef
For Each T In Db.TableDefs
    Push O, DbT_StruStr(DbT(Db, T.Name))
Next
Db_StruStr = Join(O, vbCrLf)
End Function

Sub Tst()
Dim D As Database
Dim Act$
Dim Tbl As TableDef
Dim P As Prp
Set D = OpenDatabase("C:\temp\A.accdb")
'Act = DbT_StruStr(DbT(OpenDatabase(Fs_Fb), "SkuCostChr"))
Act = DbT_StruStr(DbT(D, "Table1"))
MsgBox Act
Set Tbl = D.TableDefs("Table1")
    P = Prp(Tbl, "Fk1", "Ele")
Debug.Print "Ele=" & Prp_Val(P)
Stop
End Sub

Private Sub Db_StruStr__Tst()
Dim Db As Database
Set Db = OpenDatabase(Fs_Fb)
Str_Brw Db_StruStr(Db)
End Sub

Private Function ZIdx_FldNmAy(Idx As Index) As String()
Dim F As Field, O$()
For Each F In Idx.Fields
    Push O, F.Name
Next
ZIdx_FldNmAy = O
End Function

Private Function ZKeyStr$(Tbl As TableDef, KeyNmAy$())
Dim F$(), N%, O$(), KeyNm$, J%

N = Sz(KeyNmAy)
If N = 0 Then Exit Function
ReDim O(N - 1)
For J = 0 To N - 1
    KeyNm = KeyNmAy(J)
    F = ZIdx_FldNmAy(Tbl.Indexes(KeyNm))
    O(J) = KeyNm & "(" & Join(F) & ")"
Next
ZKeyStr = Join(O, ", ")
End Function
