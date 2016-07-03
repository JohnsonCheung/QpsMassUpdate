Attribute VB_Name = "Dta_DtSel"
Option Explicit

Function Dr_Sel(Dr, IdxAy%())
'Return a new {ODr} by selecting the fields as given in {IdxAy} of {Dr}
'Note: If the idx within IdxAy is outside Dr, the element of ODr will be empty
'Note: If ending-empty-element of ODr will be discarded.
Dim U1%
    U1 = UB(Dr)
Dim I%
Dim U2%
    For U2 = UB(IdxAy) To 0 Step -1
        I = IdxAy(U2)
        If 0 <= I Or I <= U1 Then Exit For
    Next
Dim O()
    ReDim O(U2)
Dim J%
For J = 0 To U2
    I = IdxAy(J)
    If 0 <= I And I <= U1 Then O(J) = Dr(I)
Next
Dr_Sel = O
End Function

Function Dt_Sel(Dt As TDt, SelFldNmAy$(), Optional AsFldNmLvs$, Optional WhereFldNm_WithOptionalColonPfx$, Optional Operator As eOperator, Optional Operand, Optional NewDtNm$) As TDt
'Note: WhereFldNm_WithOptionalColonPfx$: If the field name has ":" is pfx, it is the AsFldNmLvs, if no colon, it is SelFldNm
Dim UR&  ' # of URow in {Dt}
    UR = UB(Dt.DrAy)
If UR = -1 Then GoTo X

Dim AsFldNmAy$()
    If AsFldNmLvs = "" Then
        AsFldNmAy = SelFldNmAy
    Else
        AsFldNmAy = Split(AsFldNmLvs)
    End If

If Sz(AsFldNmAy) <> Sz(SelFldNmAy) Then Er "{Sz1} of {AsFldNmAy} <> {Sz2} of {SelFldNmAy}", Sz(AsFldNmAy), Sz(SelFldNmAy), Ay_ToStr(AsFldNmAy), Ay_ToStr(SelFldNmAy)

Dim WhereFldIdx% ' The idx is based on Dt.FldNmAy
    WhereFldIdx = ZWhereFldIdx(WhereFldNm_WithOptionalColonPfx, SelFldNmAy, AsFldNmAy, Dt.FldNmAy)

Dim CnoAy%()
    CnoAy = Ay_IdxAy_OfInt(Dt.FldNmAy, SelFldNmAy)

Dim ErAy$()
    Erase ErAy
    Dim J%
    For J = 0 To UB(CnoAy)
        If CnoAy(J) = -1 Then Push ErAy, ""
    Next
    
Dim O_DrAy()
    Erase O_DrAy
    
    Dim R&
    For R = 0 To UR
        Dim Dr()
            Dr = Dt.DrAy(R)

        Dim Incl As Boolean  ' Include-in-this-row?
            If WhereFldIdx > -1 Then
                Dim V     ' WhereFldV
                    V = Dr(WhereFldIdx)
                Select Case Operator
                Case eOp_In: Incl = Ay_Has(Operand, V)
                Case eOp_Fn: Incl = Application.Run(Operand, V)
                Case Else: Er "Given {Operator} is invalid", Operator
                End Select
            Else
                Incl = True
            End If
            
        If Incl Then ' Include-in-this-row?
            Push O_DrAy, Dr_Sel(Dr, CnoAy)    '<==
        End If
    Next
X:

Dim O As TDt
         O.Nm = IIf(NewDtNm = "", Dt.Nm, NewDtNm)
    O.FldNmAy = AsFldNmAy
       O.DrAy = O_DrAy
Dt_Sel = O
End Function

Private Function ZWhereFldIdx%(WhereFldNm$, SelFldNmAy$(), AsFldNmAy$(), DtFldNmAy$())
Dim WhereFldNm_As$
    WhereFldNm_As = ""
    If Left(WhereFldNm, 1) = ":" Then
        WhereFldNm_As = Mid(WhereFldNm, 2)
    End If

Dim WhereFldNm_Ws$
    If WhereFldNm_As = "" Then
        WhereFldNm_Ws = WhereFldNm
    Else
        Dim Idx%
            Idx = Ay_Idx(AsFldNmAy, WhereFldNm_As)
        
        WhereFldNm_Ws = SelFldNmAy(Idx)
    End If

Dim O%
If WhereFldNm_Ws = "" Then
    O = -1
Else
    O = Ay_Idx(DtFldNmAy, WhereFldNm_Ws)
    Assert_NotEq O, -1, "Given {WhereFldNm_WithOptionalColonPfx} cannot be found in {DtFldNmAy} nor {AsFldNmAy}", _
        WhereFldNm, _
        Ay_ToStr(DtFldNmAy), _
        Ay_ToStr(AsFldNmAy)
End If
ZWhereFldIdx = O
End Function

