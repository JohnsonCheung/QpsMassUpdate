Attribute VB_Name = "SrcCnoDef"
Option Explicit
Private A_Ay() As TCnoDef
Enum eValTy
    eStr = 1
    eStrOpt = 2
    eNbr = 3
    eNbrOpt = 4
    eDte = 5
    eDteOpt = 6
    ePos = 7
    ePosOpt = 8
End Enum
Enum eColTy
    eCT_Chr = 1
    eCT_CstRmk = 2
    eCT_CstTot = 3
    eCT_CstVal = 4
    eCT_Key = 5
    eCT_One = 6
    eCT_PjQ = 7
    eCT_Sku = 8
End Enum

Type TCnoDef
    Nm As String
    CostGp As String
    CostEle As String
    CharName As String
    ColTy As eColTy
    ValTy As eValTy
    Cno As Integer
End Type

Sub CnoDef_Dmp(Ay() As TCnoDef)
Dim J%
For J = 0 To ZAy_Sz(Ay) - 1
    ZDmp_Itm Ay(J)
Next

End Sub

Function Enm_ColTy(S$) As eColTy
Dim O As eColTy
Select Case S
Case "eCT_Chr":    O = eColTy.eCT_Chr
Case "eCT_CstRmk": O = eColTy.eCT_CstRmk
Case "eCT_CstTot": O = eColTy.eCT_CstTot
Case "eCT_CstVal": O = eColTy.eCT_CstVal
Case "eCT_Key":    O = eColTy.eCT_Key
Case "eCT_One":    O = eColTy.eCT_One
Case "eCT_PjQ":    O = eColTy.eCT_PjQ
Case "eCT_Sku":    O = eColTy.eCT_Sku
Case Else: Er "Given {S} is a not in valid Enm-eColTy-{MbrNmList}", S, "[eCT_Chr eCT_CstRmk eCT_CstTot eCT_CstVal eCT_Key eCT_One eCT_PjQ eCT_Sku]"
End Select
Enm_ColTy = O
End Function

Function Enm_ColTy_ToStr(P As eColTy)
Dim O$
Select Case P
Case eColTy.eCT_Chr:    O = "eCT_Chr"
Case eColTy.eCT_CstRmk: O = "eCT_CstRmk"
Case eColTy.eCT_CstTot: O = "eCT_CstTot"
Case eColTy.eCT_CstVal: O = "eCT_CstVal"
Case eColTy.eCT_Key:    O = "eCT_Key"
Case eColTy.eCT_One:    O = "eCT_One"
Case eColTy.eCT_PjQ:    O = "eCT_PjQ"
Case eColTy.eCT_Sku:    O = "eCT_Sku"
Case Else: Er "Enm-eColTy-{MbrVal} not in valid {MbrVal-List} of {MbrNm-List}", P, "[1 2 3 4 5 6 7 8]", "[eCT_Chr eCT_CstRmk eCT_CstTot eCT_CstVal eCT_Key eCT_One eCT_PjQ eCT_Sku]"
End Select
Enm_ColTy_ToStr = O
End Function

Function Enm_ValTy(S$) As eValTy
Dim O As eValTy
Select Case S
Case "eStr":    O = eValTy.eStr
Case "eStrOpt": O = eValTy.eStrOpt
Case "eNbr":    O = eValTy.eNbr
Case "eNbrOpt": O = eValTy.eNbrOpt
Case "eDte":    O = eValTy.eDte
Case "eDteOpt": O = eValTy.eDteOpt
Case "ePos":    O = eValTy.ePos
Case "ePosOpt": O = eValTy.ePosOpt
Case Else: Er "Given {S} is a not in valid Enm-eValTy-{MbrNmList}", S, "[eStr eStrOpt eNbr eNbrOpt eDte eDteOpt ePos ePosOpt]"
End Select
Enm_ValTy = O
End Function

Function Enm_ValTy_ToStr(P As eValTy)
Dim O$
Select Case P
Case eValTy.eStr:    O = "eStr"
Case eValTy.eStrOpt: O = "eStrOpt"
Case eValTy.eNbr:    O = "eNbr"
Case eValTy.eNbrOpt: O = "eNbrOpt"
Case eValTy.eDte:    O = "eDte"
Case eValTy.eDteOpt: O = "eDteOpt"
Case eValTy.ePos:    O = "ePos"
Case eValTy.ePosOpt: O = "ePosOpt"
Case Else: Er "Enm-eValTy-{MbrVal} not in valid {MbrVal-List} of {MbrNm-List}", P, "[1 2 3 4 5 6 7 8]", "[eStr eStrOpt eNbr eNbrOpt eDte eDteOpt ePos ePosOpt]"
End Select
Enm_ValTy_ToStr = O
End Function

Function TCnoDef(C As TCno) As TCnoDef()
Erase A_Ay
ZFld_PjQ C.PjQ
ZFld_One C.One
ZFld_Sku C.Sku
ZFld_Chr C.Chr
ZFld_CstVal C.CstVal
ZFld_CstRmk C.CstRmk
TCnoDef = A_Ay
End Function

Private Sub CnoDef_PushAy(Ay() As TCnoDef)
Dim J%
For J = 0 To ZAy_Sz(Ay) - 1
    ZAy_Push Ay(J)
Next
End Sub

Private Sub ZAy_Push(M As TCnoDef)
Dim N%
'    N = ZAy_Sz(A_Ay)
ReDim Preserve A_Ay(N)
A_Ay(N) = M
End Sub

Private Function ZAy_Sz%(Ay() As TCnoDef)
On Error Resume Next
ZAy_Sz = UBound(Ay) + 1
End Function

Private Sub ZDmp_Itm(F As TCnoDef)
With F
    'Debug.Print .Nm, Enm.ValTyStr(F.Ty), , .CostGp, .CostEle, .CharName
End With
End Sub

Private Sub ZFld(Nm$, ValTy As eValTy, Cno%, Optional CostGp$, Optional CostEle$, Optional CharName$)
Dim O As TCnoDef
With O
    .Nm = Nm
    .ValTy = ValTy
    .Cno = Cno
    .CharName = CharName
    .CostGp = CostGp
    .CostEle = CostEle
End With
ZAy_Push O
End Sub

Private Sub ZFld_Chr(C() As ChrCno)
Dim O() As TCnoDef, J%, Nm$, Cno%
For J = 0 To UBound(C)
    With C(J)
        Nm = "Char"
        Cno = .Cno
        ZFld Nm, eValTy.eStrOpt, Cno, .CostGp, .CostEle, .CharName
    End With
Next
End Sub

Private Sub ZFld_CstRmk(C() As CstRmkCno)
Dim O() As TCnoDef, J%, Nm$, Cno%
For J = 0 To UBound(C)
    With C(J)
        Nm = "Cost Rmk"
        Cno = .Cno
        ZFld Nm, eValTy.eStrOpt, Cno, .CostGp, .CostEle
    End With
Next
End Sub

Private Sub ZFld_CstVal(C() As CstValCno)
Dim O() As TCnoDef, J%, Nm$, Cno%
For J = 0 To UBound(C)
    With C(J)
        Nm = "Cost"
        Cno = .Cno
        ZFld Nm, eValTy.eNbrOpt, Cno, .CostGp, .CostEle
    End With
Next
End Sub

Private Sub ZFld_One(C As OneCno)
ZFld "ToolCst", eValTy.ePosOpt, C.ToolCst
ZFld "ProtCst", eValTy.ePosOpt, C.ProtCst
ZFld "ToolRmk", eValTy.eStrOpt, C.ToolRmk
ZFld "ProtRmk", eValTy.eStrOpt, C.ProtRmk
End Sub

Private Sub ZFld_Opt(Nm$, Ty As eValTy, Cno%)
If Cno = 0 Then Exit Sub
ZFld Nm, Ty, Cno
End Sub

Private Sub ZFld_PjQ(C As PjQCno)
Dim O() As TCnoDef
ZFld "Brand", eValTy.eStr, C.Brand
ZFld "RateCHF", eValTy.eNbrOpt, C.RateCHF
ZFld "RateJPY", eValTy.eNbrOpt, C.RateJPY
ZFld "RateUSD", eValTy.eNbrOpt, C.RateUSD
ZFld "Supplier", eValTy.eStr, C.Supplier
End Sub

Private Sub ZFld_Sku(C As SkuCno)
ZFld "SkuCst", eValTy.ePos, C.Cost
ZFld "PotentialQty", eValTy.ePosOpt, C.PotentialQty
ZFld_Opt "AssWatchHKD", eValTy.eNbrOpt, C.AssWatchHKD
ZFld_Opt "AssWatchUSD", eValTy.eNbrOpt, C.AssWatchUSD
ZFld_Opt "CompleteWatchHKD", eValTy.ePosOpt, C.CompleteWatchHKD
ZFld_Opt "CompleteWatchUSD", eValTy.ePosOpt, C.CompleteWatchUSD
ZFld_Opt "SalesmanHKD", eValTy.eNbrOpt, C.SalesmanHKD
ZFld_Opt "SalesmanUSD", eValTy.eNbrOpt, C.SalesmanUSD
End Sub
