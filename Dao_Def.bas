Attribute VB_Name = "Dao_Def"
Option Explicit

Type FldDef
    Nm As String
    Ty As Dao.DataTypeEnum
    TxtSz As Byte  ' Only for Text
    DftVal As Variant
    Des As String
    NotNul As Boolean
End Type
Type IdxDef
    Nm As String
    FldNm() As String
End Type
Type TblDef
    Nm As String
    Fld() As FldDef
    SkFld() As String
    UIdx() As IdxDef
    KIdx() As IdxDef
End Type
Type DbDef
    Nm As String
    Tbl() As TblDef
End Type
Type FldPrpDef
    FldNm As String
    PrpNm As String
    V As Variant
End Type
Type TblPrpDef
    PrpNm As String
    V As Variant
End Type
Type RelDef
    Nm As String
    PTbl As String
    FTbl As String
    FFld As String
End Type
Enum eSecTy
    eEle = 1
    eFld = 2
    eTbl = 3
End Enum
Type Section
    Ty As eSecTy
    LinAy() As String
End Type
Private Type EleLin
    Nm As String
    
End Type
Private Type FldLin
    Nm As String
    
End Type
Private Type TblLin
    Nm As String
End Type
Private O As DbDef
Private Type B
    A As String
End Type
Private A_DbDef$()
Private B As B

Function Db_Def(DbDef$()) As DbDef
A_DbDef = DbDef
Dim Section() As Section
    Section = ZSection()
Dim J%
Dim Ele() As EleLin
Dim Fld() As FldLin

For J = 0 To UBound(Section)
    Dim Ty As eSecTy
    Dim LinAy$()
    Dim I%
    For I = 0 To UB(LinAy)
        Dim L$
        L = LinAy(I)
            Select Case Ty
            Case eEle: 'ZPush_Ele ZEle_BrkLin(L)
            Case eFld: 'ZPush_Fld ZFld_BrkLin(L)
            Case eTbl: 'ZBld_Tbl L
            Case Else: Er "Invalid {SecTy}", Ty
            End Select
    Next
Next
Db_Def = O
End Function

Private Property Get ZSection() As Section()

End Property
