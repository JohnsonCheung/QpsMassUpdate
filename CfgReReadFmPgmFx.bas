Attribute VB_Name = "CfgReReadFmPgmFx"
Option Explicit
'
'Sub Tst()
'Dim Act As CfgInfo
'Act = Cfg_ReRead_FmPgmFx
'Stop
'End Sub
''This in is re-read the 2 tables: Ctt & ChrV from MuPgm Wb into CfgInfo
''This have no used, but just a reverse of writting
''These Two Ws will be read as @ChrInfo
'Property Get Cfg_ReRead_FmPgmFx() As CfgInfo
'Dim Wb As Workbook
'Dim WsCtl As Worksheet
'Dim WsChrV As Worksheet
'Dim SqvCtl
'Dim SqvChrV
'Dim O As CfgInfo
'SqvCtl = WsCtl.ListObjects(1).DataBodyRange.Value
'SqvChrV = WsChrV.ListObjects(1).DataBodyRange.Value
'O.ChrV = ZChrV(SqvChrV)
'O.Ctl = ZCtl(SqvCtl)
'Cfg_ReRead_FmPgmFx = O
'End Property
'
'Private Function ZChrV(Sqv) As ChrVRec()
'Dim O() As ChrVRec, U%, J%
'U = UBound(Sqv, 1)
'ReDim O(U - 1)
'For J = 1 To U
'    With O(J - 1)
'        .CharCode = Sqv(J, 1)
'        .ValName = Sqv(J, 2)
'        .ValCode = Sqv(J, 3)
'    End With
'Next
'ZChrV = O
'End Function
'Private Function ZCtl(Sqv) As CtlRec()
'Dim O() As CtlRec, U%, J%
'U = UBound(Sqv, 1)
'ReDim O(U - 1)
''1           2       3       4           5       6      7
''CharCode    CostGp  CostEle CharName    IsMulti IsMust  CtlType
'For J = 1 To U
'    With O(J - 1)
'        .CharName = Sqv(J, 1)
'        .CharCode = Sqv(J, 2)
'        .IsMulti = Sqv(J, 5)
'        .IsMust = Sqv(J, 6)
'    End With
'Next
'ZCtl = O
'End Function



