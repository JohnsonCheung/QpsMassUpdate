Attribute VB_Name = "Dao_TblDoCrt"
Option Explicit
Private Type TblCreationInf
    CrtSql As String
    PkSql As String
    SkSql As String
    KeySql() As String
    UIdxSql() As String
    FldPrp() As FldPrpDef
    TblPrp As TblPrpDef
End Type
Private A_T As TblDef

Function Ay_Add(Ay, ParamArray Ap())
Dim Av()
    Av = Ap
    
Dim O
    O = Ay
    Dim J%
    For J = 0 To UB(Av)
        O = Ay_AddOne(O, Av(J))
    Next
Ay_Add = O
End Function

Function Ay_AddOne(Ay, Ay1)
Dim O
    O = Ay
    PushAy O, Ay1
Ay_AddOne = O
End Function

Function Ay_BegEnd(Ay, BegIdx&, EndIdx&)
Dim O
O = Ay
Erase O
Dim J&
For J = BegIdx To EndIdx
    Push O, Ay(J)
Next
Ay_BegEnd = O
End Function

Sub Ay_Brw(Ay, Optional BrwNm$)
Dim S$
    S = Join(Ay, vbCrLf)
Str_Brw S, BrwNm
End Sub

Sub Ay_InsAt(OAy, Ele, At&)
Dim U&, J&
U = UB(OAy)
ReDim Preserve OAy(U + 1)
For J = U To At Step -1
    OAy(J + 1) = OAy(J)
Next
OAy(At) = Ele
End Sub

Function Ay_Join$(Ay, Optional Sep$ = " ")
Ay_Join = Join(Ay_StrAy(Ay), Sep)
End Function

Function Ay_Srt(Ay)
Dim O, J&, At&, I&, Ele
O = Ay
Erase O
For J = 0 To UB(Ay)
    Ele = Ay(J)
    For At = 0 To UB(O)
        If O(At) > Ele Then Exit For
    Next
    Ay_InsAt O, Ele, At
Next
Ay_Srt = O
End Function

Function Ay_Srt_IntoIdxAy(Ay) As Long()
Dim A, J&, U&, O&()
U = UB(Ay)
If U = -1 Then Exit Function
ReDim O(U)
A = Ay_Srt(Ay)
For J = 0 To U
    O(J) = Ay_Idx(Ay, A(J))
Next
Ay_Srt_IntoIdxAy = O
End Function

Function Ay_StrAy(Ay) As String()
Dim U&
U = UB(Ay)
If U = -1 Then Exit Function
Dim O$()
ReDim O(U)
Dim J&
For J = 0 To U
    O(J) = Ay(J)
Next
Ay_StrAy = O
End Function

Function Ay_ToStrAy(Ay()) As String()
Dim U&
    U = UB(Ay)
If U = -1 Then Exit Function
Dim O$()
    ReDim O(U)
Dim J&
For J = 0 To U
    O(J) = Ay(J)
Next
Ay_ToStrAy = O
End Function

Function IntAy(ParamArray Ap()) As Integer()
Dim Av()
Dim O%()
Dim J%
Dim U%
Av = Ap
U = UB(Av)
ReDim O(U)
For J = 0 To U
    O(J) = Av(J)
Next
IntAy = O
End Function

Function LngAy(ParamArray Ap()) As Long()
Dim Av()
Dim O&()
Dim J%
Dim U%
Av = Ap
U = UB(Av)
ReDim O(U)
For J = 0 To U
    O(J) = Av(J)
Next
LngAy = O
End Function

Sub TblDef_Crt(T As TblDef)
A_T = T
ZCrt ZInf
End Sub

Sub Tbl_DoCrt(Db As Database, T As TblDef)

End Sub

Sub ZCreationInf__Tst()
Dim F0 As FldDef
    With F0
        .Nm = "Table1"
        .Des = "Fld1 Des"
        .DftVal = CByte(1)
        .NotNul = True
        .Ty = dbByte
    End With
Dim F1 As FldDef
    With F1
        .Nm = "Fld1"
        .Des = "Fld1 Des"
        .DftVal = "Fld1 Default Value"
        .NotNul = False
        .Ty = dbText
        .TxtSz = 10
    End With
Dim F2 As FldDef
    With F2
        .Nm = "Fld2"
        .Des = "Fld2 Des"
        .DftVal = "Now()"
        .NotNul = True
        .Ty = dbDate
    End With

Dim Fld() As FldDef
    ReDim Fld(2)
    Fld(0) = F0
    Fld(1) = F1
    Fld(2) = F2
Dim SkFld() As String
Dim KIdx() As IdxDef
Dim UIdx() As IdxDef

Dim T As TblDef
    T.Nm = "Table1"
    T.Fld = Fld
    T.SkFld = SkFld
    T.KIdx = KIdx
    T.UIdx = UIdx

A_T = T
Dim Act As TblCreationInf
    Act = ZInf
Debug.Print Act.CrtSql
End Sub

Private Sub Ay_Srt_IntoIdxAy__Tst()
Ay_Dmp Ay_Srt_IntoIdxAy(Array(2, 5, 1, 2))
End Sub

Private Sub Ay_Srt__Tst()
Ay_Dmp Ay_Srt(Array(2, 5, 1, 2))
End Sub

Private Function DaoTy_SqlTyStr$(Ty As DataTypeEnum)
'BINARY 1 byte per character Any type of data may be stored in a field of this type. No translation of the data (for example, to text) is made. How the data is input in a binary field dictates how it will appear as output.
'BIT 1 byte Yes and No values and fields that contain only one of two values.
'TINYINT 1 byte An integer value between 0 and 255.
'MONEY 8 bytes A scaled integer between ¡V 922,337,203,685,477.5808 and 922,337,203,685,477.5807.
'DATETIME (See DOUBLE) 8 bytes A date or time value between the years 100 and 9999.
'UNIQUEIDENTIFIER 128 bits A unique identification number used with remote procedure calls.
'REAL 4 bytes A single-precision floating-point value with a range of ¡V 3.402823E38 to ¡V 1.401298E-45 for negative values, 1.401298E-45 to 3.402823E38 for positive values, and 0.
'FLOAT 8 bytes A double-precision floating-point value with a range of ¡V 1.79769313486232E308 to ¡V 4.94065645841247E-324 for negative values, 4.94065645841247E-324 to 1.79769313486232E308 for positive values, and 0.
'SMALLINT 2 bytes A short integer between ¡V 32,768 and 32,767. (See Notes)
'INTEGER 4 bytes A long integer between ¡V 2,147,483,648 and 2,147,483,647. (See Notes)
'DECIMAL 17 bytes An exact numeric data type that holds values from 1028 - 1 through - 1028 - 1. You can define both precision (1 - 28) and scale (0 - defined precision). The default precision and scale are 18 and 0, respectively.
'TEXT 2 bytes per character (See Notes) Zero to a maximum of 2.14 gigabytes.
'IMAGE As required Zero to a maximum of 2.14 gigabytes. Used for OLE objects.
'CHARACTER 2 bytes per character (See Notes) Zero to 255 characters.
Dim O$
Select Case Ty
Case DataTypeEnum.dbBoolean: O = "BIT"
Case DataTypeEnum.dbBigInt: O = "DECIMAL"
Case DataTypeEnum.dbBinary: O = "BINARY"
Case DataTypeEnum.dbByte: O = "TINYINT"
Case DataTypeEnum.dbChar: O = "TEXT"
Case DataTypeEnum.dbCurrency: O = "MONEY"
Case DataTypeEnum.dbDate: O = "DATETIME"
Case DataTypeEnum.dbDecimal: O = "DECIMAL"
Case DataTypeEnum.dbDouble: O = "FLOAT"
Case DataTypeEnum.dbFloat: O = "FLOAT"
Case DataTypeEnum.dbGUID: O = "UNIQUEIDENTIFIER"
Case DataTypeEnum.dbInteger: O = "INTEGER"
Case DataTypeEnum.dbLong: O = "DECIMAL"
Case DataTypeEnum.dbLongBinary: O = "BINARY"
Case DataTypeEnum.dbMemo: O = "TEXT"
Case DataTypeEnum.dbNumeric: O = "DECIMAL"
Case DataTypeEnum.dbSingle: O = "REAL"
Case DataTypeEnum.dbText: O = "CHARACTER"
Case DataTypeEnum.dbTime: O = "DATETIME"
Case DataTypeEnum.dbTimeStamp: O = "DATETIME"
Case DataTypeEnum.dbVarBinary: O = "BINARY"
Case Else: Er "Invalid Dao-{Ty}", Ty
End Select
DaoTy_SqlTyStr = O
End Function

Private Sub ZCrt(Inf As TblCreationInf)

End Sub

Private Property Get ZFldPrp_Des() As FldPrpDef()

End Property

Private Property Get ZFldPrp_DftVal() As FldPrpDef()

End Property

Private Sub ZFldPrp_PushAy(OAy() As FldPrpDef, Ay() As FldPrpDef)

End Sub

Private Property Get ZInf() As TblCreationInf
Dim O As TblCreationInf
    O.CrtSql = ZInf_CrtSql
    O.TblPrp = ZTblPrp_Des
    O.FldPrp = ZInf_FldPrp
ZInf = O
End Property

Private Property Get ZInf_CrtSql$()
Dim A$
    A = ZSqlFldLst$
Dim B$
    B = ZSqlConstrain
ZInf_CrtSql = Fmt_QQ("Create Table ? (?)?", A_T.Nm, A, B)
End Property

Private Property Get ZInf_FldPrp() As FldPrpDef()
Dim O() As FldPrpDef
ZFldPrp_PushAy O, ZFldPrp_Des
ZFldPrp_PushAy O, ZFldPrp_DftVal
ZInf_FldPrp = O
End Property

Private Property Get ZSqlConstrain$()
Dim O$()
    O = ZSqlConstrain_Ay
    O = Ay_AddPfx(O, "," & vbCrLf)
ZSqlConstrain = Join(O)
End Property

Private Property Get ZSqlConstrain_Ay() As String()
ZSqlConstrain_Ay = Ay_Add(ZSqlConstrain_KIdx, ZSqlConstrain_UIdx)
End Property

Private Property Get ZSqlConstrain_KIdx() As String()
ZSqlConstrain_KIdx = ZSqlConstrain_OneIdxAy(A_T.KIdx, False)
End Property

Private Property Get ZSqlConstrain_NKey%(K() As IdxDef)
On Error Resume Next
ZSqlConstrain_NKey = UBound(K) + 1
End Property

Private Function ZSqlConstrain_OneIdx$(Idx As IdxDef, IsUnique As Boolean)
Const C1 = "Constrain ? (?)"
Const C2 = "Constrain ? Unique (?)"
Dim C$
    C = IIf(IsUnique, C2, C1)
Dim Nm$
    Nm = Idx.Nm
Dim F$
    F = Join(Idx.FldNm, ", ")
ZSqlConstrain_OneIdx = Fmt_QQ(C, Nm, F)
End Function

Private Function ZSqlConstrain_OneIdxAy(K() As IdxDef, IsUnique As Boolean) As String()
Dim N%
    N = ZSqlConstrain_NKey(K)
If N = 0 Then Exit Function
ReDim O$(N - 1)
    Dim J%
    For J = 0 To N - 1
        O(J) = ZSqlConstrain_OneIdx(K(J), IsUnique)
    Next
ZSqlConstrain_OneIdxAy = O
End Function

Private Property Get ZSqlConstrain_UIdx() As String()
ZSqlConstrain_UIdx = ZSqlConstrain_OneIdxAy(A_T.UIdx, True)
End Property

Private Property Get ZSqlFldLst$()
ZSqlFldLst = Join(ZSqlFldLst_Ay, "," & vbCrLf)
End Property

Private Property Get ZSqlFldLst_Ay() As String()
Dim U%
    U = UBound(A_T.Fld)
ReDim O$(U)
    Dim J%
    For J = 0 To U
        O(J) = ZSqlFldLst_OneFld(J)
    Next
ZSqlFldLst_Ay = O
End Property

Private Function ZSqlFldLst_OneFld$(Idx%)
Dim F As FldDef
    F = A_T.Fld(Idx)
Dim T$
    If Idx = 0 Then
        If F.Nm <> A_T.Nm Then Er "First field-{Name} is should be same as table-{Name}", F.Nm, A_T.Nm
        T = "AUTOINCREMENT Not Null Primary Key"
    Else
        Dim Ty$
            Ty = ZSqlFldLst_Ty$(F.Ty, F.TxtSz)
        Dim NotNul$
            If F.NotNul Then NotNul = " Not Null"
        T = Ty & NotNul
    End If
ZSqlFldLst_OneFld = F.Nm & " " & T
End Function

Private Function ZSqlFldLst_OthTy$(Ty As Dao.DataTypeEnum)
ZSqlFldLst_OthTy = DaoTy_SqlTyStr(Ty)
End Function

Private Function ZSqlFldLst_TxtTy$(Sz As Byte)
ZSqlFldLst_TxtTy = "TEXT(" & Sz & ")"
End Function

Private Function ZSqlFldLst_Ty$(Ty As DataTypeEnum, Sz As Byte)
Dim O$
If Ty = dbText Then
    O = ZSqlFldLst_TxtTy(Sz)
Else
    O = ZSqlFldLst_OthTy(Ty)
End If
ZSqlFldLst_Ty = O
End Function

Private Property Get ZTblPrp_Des() As TblPrpDef

End Property
