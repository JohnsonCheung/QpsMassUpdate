Attribute VB_Name = "SrcCno"
Option Explicit
Private A_HdSqv
Private A_WsNm$
'====???Cno ====================================
Type KeyCno
    Pj As Integer
    Sku As Integer
    QDte As Integer
End Type
Type SkuCno
    Cost As Integer
    PotentialQty  As Integer
    CompleteWatchUSD  As Integer
    CompleteWatchHKD  As Integer
    AssWatchUSD  As Integer
    AssWatchHKD  As Integer
    SalesmanHKD  As Integer
    SalesmanUSD  As Integer
End Type
Type PjQCno
    Brand As Integer
    Supplier As Integer
    RateUSD As Integer
    RateCHF As Integer
    RateJPY As Integer
End Type
Type OneCno
    ProtCst As Integer
    ProtRmk As Integer
    ToolCst As Integer
    ToolRmk As Integer
End Type
Type Cst1Cno                ' Cst1 is Multi-Cost-Ele-Cost: Row-5 is EleGp??Ele??'
    CostGp As String        ' Row-1
    CostEle As String       ' Row-6
    Cno As Integer
    FldNmAtRow5 As String
End Type
Type Cst2Cno                ' Cst2 is Single-Cost-Ele-Cost: Row-5 is 'EleGp??Tot' & No '{A}Ele??', where A='EleGp??Tot'
    CostGp As String        ' Row-1
    CostEle As String       ' Row-6, remove Sfx(" $").
    Cno As Integer
    FldNmAtRow5 As String
    FldNmAtRow6 As String   ' Must have Sfx-" $", otherwise, it is error
End Type
Type CstValCno ' Combine Cst1Cno and Cst2Cno as one
    IsSingleEle As Boolean
    CostGp As String
    CostEle As String
    Cno As Integer
End Type
Type CstRmkCno            ' CstRmk is Cost-Ele-Rmk:  Row-5 is "EleGp??Ele??Rmk
    CostGp As String      ' Row 1
    CostEle As String     ' Row-6, remove Sfx(" Rmk").  If ='Other Cost (if any)', adjust to 'Other Cost#1 (if any)'
                          ' The ExportQry need to fix to (1) export 2 OtherCost (2) Make Sure the Row-6 is using '{CostEle} Rmk'
    Cno As Integer
    FldNmAtRow5 As String
End Type
Type CstTotCno            ' CstTot for multiple Cost-Ele-Cost: Row-5 is 'EleGp??Tot' & with '{A}Ele??', where A='EleGp??Tot'
                          ' It is for validate DifR1FormulaEr
    CostGp As String      ' Row 1
    CostEle As String     ' Row-6, remove Sfx(" $").
    Cno As Integer
    FldNmAtRow5 As String
    FldNmAtRow6 As String   ' Must have Sfx-" $", otherwise, it is error
End Type
Type ChrCno
    CostGp As String
    CostEle As String
    CharName As String  '
    CharCode As String
    FldNmAtRow5 As String
    Cno As Integer
End Type
'=====


Type TCno
    Key As KeyCno
    Sku As SkuCno
    PjQ As PjQCno
    One As OneCno
    CstTot() As CstTotCno   ' Cost-Total for Multi-Cost-Element
    CstVal() As CstValCno
    CstRmk() As CstRmkCno   ' Cost-Rmk
    Chr() As ChrCno
    NCstTot As Integer
    NCstVal As Integer
    NCstRmk As Integer
    NChr As Integer
End Type

Function TCno(HdSqv, WsNm$) As TCno
A_HdSqv = HdSqv

A_WsNm = WsNm
Dim O As TCno
Dim Cst1() As Cst1Cno
Dim Cst2() As Cst2Cno
With O
    .Chr = Z_ChrCol
    Cst1 = Z_Cst1Col          ' Cst1 = Many-Cost-Ele cost
    Cst2 = Z_Cst2Col          ' Cst2 = One-Cost-Ele  cost
    .CstVal = Z_CstValCol(Cst1, Cst2) ' CstVal = Combine Cst1 & Cst2
    .CstRmk = Z_CstRmkCol  ' CstRmk = Cost Rmk
    .CstTot = Z_CstTotCol
    .Key = Z_KeyCol
    .One = Z_OneCol      ' OneCol is for Tbl-ProjOneTimeCost
    .PjQ = Z_PjQCol        ' PjQCol is for Tbl-ProjQ
    .Sku = Z_SkuCol  ' SkuCol is for Tbl-Sku
    .NChr = ZSz_Chr(.Chr)
    .NCstRmk = ZSz_CstRmk(.CstRmk)
    .NCstTot = ZSz_CstTot(.CstTot)
    .NCstVal = ZSz_CstVal(.CstVal)
End With
TCno = O
End Function

Private Function ZSz_Chr(Ay() As ChrCno)
On Error Resume Next
ZSz_Chr = UBound(Ay) + 1
End Function

Private Function ZSz_Cst1(Ay() As Cst1Cno)
On Error Resume Next
ZSz_Cst1 = UBound(Ay) + 1
End Function

Private Function ZSz_Cst2(Ay() As Cst2Cno)
On Error Resume Next
ZSz_Cst2 = UBound(Ay) + 1
End Function

Private Function ZSz_CstRmk(Ay() As CstRmkCno)
On Error Resume Next
ZSz_CstRmk = UBound(Ay) + 1
End Function

Private Function ZSz_CstTot(Ay() As CstTotCno)
On Error Resume Next
ZSz_CstTot = UBound(Ay) + 1
End Function

Private Function ZSz_CstVal(Ay() As CstValCno)
On Error Resume Next
ZSz_CstVal = UBound(Ay) + 1
End Function

Private Function Z_ChrCol() As ChrCno()
Dim CnoAy%()
    Dim J&
    For J = 1 To UBound(A_HdSqv, 2)
        If A_HdSqv(5, J) Like "ChrGp??Ele??Chr??" Then
            Push CnoAy, J
        End If
    Next
    
Dim U%
    U = UB(CnoAy)
    
If U < 0 Then Er "No field-label like ChrGp??Ele??Chr?? in row-5"

Dim O() As ChrCno
    ReDim O(U) As ChrCno
    For J = 0 To U
        Dim C%
            C = CnoAy(J)
        Dim M As ChrCno
            M.Cno = C
            M.CostGp = A_HdSqv(2, C)
            M.CostEle = A_HdSqv(3, C)
            M.CharCode = A_HdSqv(4, C)
            M.CharName = A_HdSqv(6, C)
        O(J) = M
    Next
Z_ChrCol = O
End Function

Private Function Z_Cst1Col() As Cst1Cno()  ' Cst1 = Many-Cost-Ele-Cost
Dim CnoAy%()
    Dim J&
    For J = 1 To UBound(A_HdSqv, 2)
        If A_HdSqv(5, J) Like "EleGp??Ele??" Then
            Push CnoAy, J
        End If
    Next
    
Dim U%
    U = UB(CnoAy)
If U = -1 Then Er "No field-name like EleGp??Ele?? at row 5"
Dim O() As Cst1Cno
    ReDim O(U) As Cst1Cno
    For J = 0 To U
        Dim C%
            C = CnoAy(J)
        Dim M As Cst1Cno
            M.Cno = C
            M.CostGp = A_HdSqv(1, C)
            M.CostEle = A_HdSqv(6, C)
            M.FldNmAtRow5 = A_HdSqv(5, C)
        O(J) = M
    Next
Z_Cst1Col = O
End Function

Private Function Z_Cst2Col() As Cst2Cno()  ' Cst2 = Single-Cost-Ele-Cost
Dim CnoAy%()
    Dim J&, I&
    For J = 1 To UBound(A_HdSqv, 2)
        If Not A_HdSqv(5, J) Like "EleGp??Tot" Then GoTo Nxt
        Dim A$
            A = Left(A_HdSqv(5, J), 7) & "Ele??"
        Dim IsSingleCstEle As Boolean
            IsSingleCstEle = True
            For I = 1 To UBound(A_HdSqv, 2)
                If A_HdSqv(5, I) Like A Then IsSingleCstEle = False: Exit For
            Next
        If IsSingleCstEle Then Push CnoAy, J
Nxt:
    Next
Dim U%
    U = UB(CnoAy)
If U = -1 Then Er "No total-cost-column with single cost element"
ReDim O(U) As Cst2Cno
    For J = 0 To U
        Dim C%
            C = CnoAy(J)
        Dim M As Cst2Cno
            M.Cno = C
            M.CostGp = A_HdSqv(1, C)
            M.CostEle = Str_RmvSfx(A_HdSqv(6, C), " $")
            M.FldNmAtRow5 = A_HdSqv(5, C)
            M.FldNmAtRow6 = A_HdSqv(6, C)
        O(J) = M
    Next
    
For J = 0 To U
    If Not Str_IsSfx(O(J).FldNmAtRow6, " $") Then Er "{FldNmAtRow6} should have [ $] as sfx", O(J).FldNmAtRow6
Next

Z_Cst2Col = O
End Function

Private Function Z_CstRmkCol() As CstRmkCno() ' CstRmk = Cost-Rmk
Dim CnoAy%()
    Dim J&
    For J = 1 To UBound(A_HdSqv, 2)
        If A_HdSqv(5, J) Like "EleGp??Ele??Rmk" Then
            Push CnoAy, J
        End If
    Next

Dim U%
    U = UB(CnoAy)
    If U = -1 Then Er "No field-name like EleGp??Ele??Rmk in row-5"

ReDim O(U) As CstRmkCno
    Dim M As CstRmkCno
    Dim Cno%
    For J = 0 To U
        Cno = CnoAy(J)
        M.Cno = Cno
        M.CostGp = A_HdSqv(1, Cno)
        M.CostEle = Z_CstRmkCol__CostEle(A_HdSqv, Cno)
        M.FldNmAtRow5 = A_HdSqv(5, Cno)
        O(J) = M
    Next
Z_CstRmkCol = O
End Function

Private Function Z_CstRmkCol__CostEle$(A_HdSqv, Cno%)
'CostGp = Other Cost (if any)
'CostEle = Other Cost#1 (if any)

Dim IsSingle As Boolean
    IsSingle = A_HdSqv(5, Cno - 1) Like "*Tot"

Dim O$
    If IsSingle Then    ' Single mean the cost-ele under the cost gp is single element
        O = A_HdSqv(6, Cno)
        O = RmvSfx(O, " Rmk")
    Else
        O = A_HdSqv(2, Cno - 1)
    End If

    If O = "Other Cost (if any)" Then
        O = "Other Cost#1 (if any)"
    End If
Z_CstRmkCol__CostEle = O
End Function

Private Function Z_CstTotCol() As CstTotCno()  ' Cst1 = Total of Many-Cost-Ele-Cost
Dim CnoAy%(), J&, I&
For J = 1 To UBound(A_HdSqv, 2)
    If Not (A_HdSqv(5, J) Like "EleGp??Tot") Then GoTo Nxt
    
    Dim A$
        A = Left(A_HdSqv(5, J), 7) & "Ele??"
    Dim IsMultiCstEle As Boolean
        IsMultiCstEle = False
        For I = 1 To UBound(A_HdSqv, 2)
            If A_HdSqv(5, I) Like A Then IsMultiCstEle = True: Exit For
        Next
    If IsMultiCstEle Then Push CnoAy, J
Nxt:
Next
Dim U%
    U = UB(CnoAy)
    
If U = -1 Then Er "No many-cost-ele-cost field-name.  That field with EleGp??Tot, but no Ele??"

ReDim O(U) As CstTotCno
    For J = 0 To U
        Dim C%
            C = CnoAy(J)
        Dim M As CstTotCno
            M.Cno = C
            M.CostGp = A_HdSqv(1, C)
            M.CostEle = Str_RmvSfx(A_HdSqv(6, C), " $")
            M.FldNmAtRow5 = A_HdSqv(5, C)
            M.FldNmAtRow6 = A_HdSqv(6, C)
        O(J) = M
    Next

For J = 0 To U
    If Not Str_IsSfx(O(J).FldNmAtRow6, " $") Then Er "{FldNmAtRow6} should have [ $] as sfx", O(J).FldNmAtRow6
Next
Z_CstTotCol = O
End Function

Private Function Z_CstValCol(C1() As Cst1Cno, C2() As Cst2Cno) As CstValCno()
Dim U%, U1%, U2%, I%, J%
U1 = UBound(C1)
U2 = UBound(C2)
U = U1 + U2 + 1
Dim O() As CstValCno
ReDim O(U)
For J = 0 To U1
    With O(I)
        .Cno = C1(J).Cno
        .CostEle = C1(J).CostEle
        .CostGp = C1(J).CostGp
        .IsSingleEle = False
    End With
    I = I + 1
Next
For J = 0 To U2
    With O(I)
        .Cno = C2(J).Cno
        .CostEle = C2(J).CostEle
        .CostGp = C2(J).CostGp
        .IsSingleEle = True
    End With
    I = I + 1
Next
Dim O1() As CstValCno
O1 = Z_CstValCol__OtherCostIfAny(O)
Z_CstValCol = O1
End Function

Private Function Z_CstValCol__OtherCostIfAny(P() As CstValCno) As CstValCno()
Dim O() As CstValCno
O = P
Dim J%
For J = 0 To UBound(P)
    If O(J).CostEle = "Other Cost (if any)" Then
        O(J).CostEle = "Other Cost#1 (if any)"
    End If
Next
Z_CstValCol__OtherCostIfAny = O
End Function

Private Function Z_KeyCol() As KeyCno
Dim O As KeyCno
O.Pj = Z_XXXColCno(A_HdSqv, "ProjNo")
O.Sku = Z_XXXColCno(A_HdSqv, "Sku")
O.QDte = Z_XXXColCno(A_HdSqv, "QuoteDate")
With O
    If .Sku = 0 Then Z_MissingEr "ProjQ", "Sku"
    If .Pj = 0 Then Z_MissingEr "ProjQ", "ProjNo"
    If .QDte = 0 Then Z_MissingEr "ProjQ", "QuoteDate"
End With
Z_KeyCol = O
End Function

Private Sub Z_MissingEr(TblNm$, ColNm$)
Er "{Col} of {Tbl} is missing in {Ws}, ColNm, TblNm, A_WsNm)"
End Sub

Private Sub Z_MissingWarning(TblNm$, ColNm$)
Debug.Print Fmt_QQ("Warning: Ws(?).Tbl(?).Col(?) of is missing", A_WsNm, TblNm, ColNm)
End Sub

Private Function Z_OneCol() As OneCno
Dim O As OneCno
O.ProtCst = Z_XXXColCno(A_HdSqv, "OneTimeCost01")
O.ProtRmk = Z_XXXColCno(A_HdSqv, "OneTimeCost01Rmk")
O.ToolCst = Z_XXXColCno(A_HdSqv, "OneTimeCost02")
O.ToolRmk = Z_XXXColCno(A_HdSqv, "OneTimeCost02Rmk")
If O.ProtCst = 0 Then Z_MissingWarning "ProjOneTimeCost", "Prototype Cost"
If O.ProtRmk = 0 Then Z_MissingWarning "ProjOneTimeCost", "Prototype Cost Remark"
If O.ToolCst = 0 Then Z_MissingWarning "ProjOneTimeCost", "Tooling Cost"
If O.ToolRmk = 0 Then Z_MissingWarning "ProjOneTimeCost", "Tooling Cost Remark"
Z_OneCol = O
End Function

Private Function Z_PjQCol() As PjQCno
Dim O As PjQCno
O.RateCHF = Z_XXXColCno(A_HdSqv, "RateCHF")
O.RateJPY = Z_XXXColCno(A_HdSqv, "RateJPY")
O.RateUSD = Z_XXXColCno(A_HdSqv, "RateUSD")
O.Supplier = Z_XXXColCno(A_HdSqv, "Supplier")
O.Brand = Z_XXXColCno(A_HdSqv, "Brand")
With O
    If .Brand = 0 Then Z_MissingEr "ProjQ", "Brand"
    If .RateCHF = 0 Then Z_MissingEr "ProjQ", "RateCHF"
    If .RateJPY = 0 Then Z_MissingEr "ProjQ", "RateJPY"
    If .RateUSD = 0 Then Z_MissingEr "ProjQ", "RateUSD"
    If .Supplier = 0 Then Z_MissingEr "ProjQ", "Supplier"
End With
Z_PjQCol = O
End Function

Private Function Z_SkuCol() As SkuCno
Dim O As SkuCno
O.AssWatchHKD = Z_XXXColCno(A_HdSqv, "AssWatchHKD")
O.AssWatchUSD = Z_XXXColCno(A_HdSqv, "AssWatchUSD")
O.CompleteWatchHKD = Z_XXXColCno(A_HdSqv, "CompleteWatchHKD")
O.CompleteWatchUSD = Z_XXXColCno(A_HdSqv, "CompleteWatchUSD")
O.SalesmanHKD = Z_XXXColCno(A_HdSqv, "SalesmanHKD")
O.SalesmanUSD = Z_XXXColCno(A_HdSqv, "SalesmanUSD")
O.SalesmanUSD = Z_XXXColCno(A_HdSqv, "SalesmanUSD")
O.Cost = Z_XXXColCno(A_HdSqv, "SkuCost")
O.PotentialQty = Z_XXXColCno(A_HdSqv, "PotentialQty")
With O
    If .AssWatchHKD = 0 Then Z_MissingWarning "ProjQ", "AssWatchHKD"
    If .AssWatchUSD = 0 Then Z_MissingWarning "ProjQ", "AssWatchUSD"
    If .CompleteWatchHKD = 0 Then Z_MissingWarning "ProjQ", "CompleteWatchHKD"
    If .CompleteWatchUSD = 0 Then Z_MissingWarning "ProjQ", "RateUSD"
    If .SalesmanHKD = 0 Then Z_MissingWarning "ProjQ", "SalesmanHKD"
    If .SalesmanUSD = 0 Then Z_MissingWarning "ProjQ", "SalesmanUSD"
    If .PotentialQty = 0 Then Z_MissingEr "ProjQ", "PotentialQty"
    If .Cost = 0 Then Z_MissingEr "ProjQ", "SkuCost"
End With
Z_SkuCol = O
End Function

Private Function Z_XXXColCno%(A_HdSqv, Row5Val$)
Dim J%
For J = 1 To UBound(A_HdSqv, 2)
    If A_HdSqv(5, J) = Row5Val Then Z_XXXColCno = J: Exit Function
Next
End Function
