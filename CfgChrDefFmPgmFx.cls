VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CfgChrDefFmPgmFx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private A_Dic_OfChrCd_ToChrDefIdx As Dictionary
Private A_Dic_OfChrNm_ToChrCd As Dictionary
Private A_ChrDefAy() As ChrDef
Private A_ChrCdDotNmAy$()
Private A_VdtChrCdAy$()

Property Get ChrCdAy() As String()
Dim U&
    U = UBound(A_ChrCdDotNmAy)
Dim O$()
    ReDim O(U)
Dim J%
For J = 0 To U
    With Brk(A_ChrCdDotNmAy(J), ".")
        O(J) = .S1
    End With
Next
ChrCdAy = O
End Property

Property Get ChrCdDotNmAy() As String()
ChrCdDotNmAy = A_ChrCdDotNmAy
End Property

Function ChrCd_ChrDef(ChrCd$) As ChrDef
'Const Pfx = "ZBRACELET_FAB"
'If IsPfx(ChrCd, Pfx) Then Stop
If Not A_Dic_OfChrCd_ToChrDefIdx.Exists(ChrCd) Then Er "Given {ChrCd} not exist in A_Dic_OfChrCd_ToChrDefIdx", ChrCd
ChrCd_ChrDef = A_ChrDefAy(A_Dic_OfChrCd_ToChrDefIdx(ChrCd))
End Function

Function ChrDef(Idx%) As ChrDef
ChrDef = A_ChrDefAy(Idx)
End Function

Property Get ChrNmAy() As String()
Dim U&
    U = UBound(A_ChrCdDotNmAy)
Dim O$()
    ReDim O(U)
Dim J%
For J = 0 To U
    With Brk(A_ChrCdDotNmAy(J), ".")
        O(J) = .S2
    End With
Next
ChrNmAy = O
End Property

Function ChrNm_ChrDef(ChrNm$) As ChrDef
Dim ChrCd$
ChrCd = A_Dic_OfChrNm_ToChrCd(ChrNm)
ChrNm_ChrDef = ChrCd_ChrDef(ChrCd)
End Function

Function ChrNm_ChrValNmAy(ChrNm$) As String()
ChrNm_ChrValNmAy = ChrNm_ChrDef(ChrNm).Dic_OfValNm_ToValCd.Keys
End Function

Function VdtChrCdAy() As String()
VdtChrCdAy = A_VdtChrCdAy
End Function

Private Sub Class_Initialize()
'#Read PgmFx-Ws-{Ctl ChrV} into A_*
' A_{Dic_OfChrCd_ToChrDefIdx
'    Dic_OfChrNm_ToChrCd
'    ChrDefAy
'    ChrCdDotNmAy}
'#Notes:
'A_ChrDefAy is primary.  All-other-A_* comes from this primary.
'PgmFx-Ws-Ctl-fields = ChrCd CstGp CstEle | ChrNm IsMulti IsMust
'PgmFx-Ws-ChrV-fields = ChrCd ChrValNm | ChrValCd
'PgmFx-Ws-ChrV goes into A_ChrDefAy().Dic_Of
Dim D As ChrDef
Dim J%
Dim Dt As TDt
Dim U%
Dim DrAy()
Dim Dic As Dictionary
Dim FldNmAy$()
Dim CnoChrCd%
Dim CnoChrNm%
Dim CnoCstGp%
Dim CnoCstEle%
Dim CnoIsMulti%
Dim CnoIsMust%
Dim Dr()
Dim O_Dic1 As New Dictionary
Dim O_Dic2 As New Dictionary


Dt = Ws_Dt(WsCtl)
U = UB(Dt.DrAy)
DrAy = Dt.DrAy ' PgmFx-Ws-Ctl
FldNmAy = Dt.FldNmAy
CnoChrCd = Ay_Idx(FldNmAy, "ChrCd")
CnoChrNm = Ay_Idx(FldNmAy, "ChrNm")
CnoCstGp = Ay_Idx(FldNmAy, "CstGp")
CnoCstEle = Ay_Idx(FldNmAy, "CstEle")
CnoIsMulti = Ay_Idx(FldNmAy, "IsMulti")
CnoIsMust = Ay_Idx(FldNmAy, "IsMust")
Set Dic = ZDic_OfChrCd_ToDicOfValNmToValCd

ReDim A_ChrDefAy(U)     '<== A_ChrDefAy has same size of PgmFx-Ws-Ctl
For J = 0 To U
    Dr = DrAy(J)    ' PgmFx-Ws-Ctl-Row
    With D
        .CharCode = Dr(CnoChrCd)
        .CharName = Dr(CnoChrNm)
        .CostEle = Dr(CnoCstEle)
        .CostGp = Dr(CnoCstGp)
        Set .Dic_OfValNm_ToValCd = Dic(.CharCode)   'Dic = ZDic_OfChrCd_ToDicOfValNmToValCd
        .IsMulti = Dr(CnoIsMulti)
        .IsMust = Dr(CnoIsMust)
    End With
    A_ChrDefAy(J) = D       '<== A_ChrDefAy, the primary
Next
ReDim A_ChrCdDotNmAy(UBound(A_ChrDefAy))
For J = 0 To UBound(A_ChrDefAy)
    D = A_ChrDefAy(J)
    A_ChrCdDotNmAy(J) = D.CharCode & "." & D.CharName       '<=== A_ChrCdDotNmAy
    O_Dic1.Add D.CharCode, J
    O_Dic2.Add D.CharName, D.CharCode
Next

Set A_Dic_OfChrCd_ToChrDefIdx = O_Dic1  '<===A_Dic_OfChrCd_ToChrDefIdx
Set A_Dic_OfChrNm_ToChrCd = O_Dic2      '<===A_Dic_OfChrNm_ToChrCd
A_VdtChrCdAy = Ay_ToStrAy(Dt_ColAy(Ws_Dt(ZPgmFx_Ws_VdtChrCd), "ChrCd"))
End Sub

Private Function ZDic_OfChrCd_ToDicOfValNmToValCd() As Dictionary
Dim Ws As Worksheet
    Set Ws = ZPgmFx_Ws_ChrV
Dim Dt As TDt
    Dt = Ws_Dt(Ws)
Dim CnoChrCd%
Dim CnoChrValNm%
Dim CnoChrValCd%
    CnoChrCd = Ay_Idx(Dt.FldNmAy, "ChrCd")
    CnoChrValNm = Ay_Idx(Dt.FldNmAy, "ChrValNm")
    CnoChrValCd = Ay_Idx(Dt.FldNmAy, "ChrValCd")
    
Dim O As New Dictionary
    Dim J%
    For J = 0 To UB(Dt.DrAy)
        Dim Dr()
            Dr = Dt.DrAy(J)
        Dim ChrCd$, ChrValNm$, ChrValCd$
        ChrCd = Dr(CnoChrCd)
        ChrValNm = Dr(CnoChrValNm)
        ChrValCd = Dr(CnoChrValCd)
        Dim Dic As Dictionary
        If O.Exists(ChrCd) Then
            Set Dic = O(ChrCd)
        Else
            Set Dic = New Dictionary
            O.Add ChrCd, Dic
        End If
        Dic.Add ChrValNm, ChrValCd
    Next
Set ZDic_OfChrCd_ToDicOfValNmToValCd = O
End Function

Private Function ZPgmFx_Ws_ChrV() As Worksheet
Set ZPgmFx_Ws_ChrV = yWbMassUpd.Sheets("ChrV")
End Function

Private Property Get ZPgmFx_Ws_Ctl() As Worksheet
Set ZPgmFx_Ws_Ctl = yWbMassUpd.Sheets("Ctl")
End Property

Private Property Get ZPgmFx_Ws_VdtChrCd() As Worksheet
Set ZPgmFx_Ws_VdtChrCd = yWbMassUpd.Sheets("VdtChrCd")
End Property
