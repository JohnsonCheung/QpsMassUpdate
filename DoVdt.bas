Attribute VB_Name = "DoVdt"
Option Explicit
Option Base 0
Option Compare Text
Enum eErWsVer
    eV1ErAndLst = 0 ' ErWs and LstWs
    eV2DropDown = 1 ' Selection list as dropdown box in selection column of same Er worksheet
    eV3SelInEr = 2  ' Selection list in Selection column of same Er worksheet
    eV4SelInSep = 3 ' Selection list in separate worksheet
End Enum
Enum eWhichWs
    eWrkWs = 1
    eOrgWs = 2
End Enum

Type RedAdr
    WhichWs() As eWhichWs
    Adr() As String
End Type

Private A_Src As TSrc
Private A_ErVer As eErWsVer

Sub Do_Vdt(Src As TSrc, ErVer As eErWsVer)
A_Src = Src
A_ErVer = ErVer
Dim Wb As Workbook
    Set Wb = A_Src.Org.Ws.Parent
    
Fct.AlignTwoWb Wb

Dim WrkWs As Worksheet
    Set WrkWs = A_Src.Wrk.Ws

Dim OrgWs As Worksheet
    Set OrgWs = A_Src.Org.Ws

Dim WrkCno As TCno
    WrkCno = A_Src.Wrk.Cno

Dim OrgCno As TCno
    OrgCno = A_Src.Wrk.Cno
    
FmtFilter_DoRestore WrkWs       ' Done at beginning of validation
FmtColor_DoRestore WrkWs
FmtChrCol_AsTxt WrkWs, WrkCno
FmtChrCol_AsTxt OrgWs, OrgCno
WrkWs.Hyperlinks.Delete

VBA.Interaction.DoEvents
Wb_DltWs Wb, DtaChgWsNm
Wb_DltWs Wb, ErWsNm
Wb_DltWs Wb, LstWsNm
VBA.Interaction.DoEvents

Dim D As TDtaErOpt
    D = Src_DtaEr

DtaEr_DoCrt_ErWs D, WrkWs, A_ErVer

If D.Some Then
    Dim A As RedAdr
        A = RedAdr(D.Ay)
        
    FmtColor_DoPaint_Red OrgWs, WrkWs, A
    FmtFilter_DoSetRed WrkWs, A
Else
    DtaChg_DoCrt_Ws_and_PaintWrkWs Src_DtaChg, Wb
End If
Wb.Save
Application.ScreenUpdating = True
ZShwMsg D.Some
End Sub

Function Enm_WhichWs(S$) As eWhichWs
Dim O As eWhichWs
Select Case S
Case "eWrkWs": O = eWhichWs.eWrkWs
Case "eOrgWs": O = eWhichWs.eOrgWs
Case Else: Er "Given {S} is a not in valid Enm-eWhichWs-{MbrNmList}", S, "[eWrkWs eOrgWs]"
End Select
Enm_WhichWs = O
End Function

Function Enm_WhichWs_ToStr(P As eWhichWs)
Dim O$
Select Case P
Case eWhichWs.eWrkWs: O = "eWrkWs"
Case eWhichWs.eOrgWs: O = "eOrgWs"
Case Else: Er "Enm-eWhichWs-{MbrVal} not in valid {MbrVal-List} of {MbrNm-List}", P, "[1 2]", "[eWrkWs eOrgWs]"
End Select
Enm_WhichWs_ToStr = O
End Function

Private Property Get RedAdr(Ay() As TDtaEr) As RedAdr
Dim OA$()
Dim OW() As eWhichWs
    Dim J&
    For J = 0 To UBound(Ay)
        With Ay(J)
            Select Case Ay(J).Ty
            Case eDtaErTy.eChrCdNotFndEr:   Push OA, .ChrCdNotFnd.ShwFld.Adr:   Push OW, eWrkWs
            Case eDtaErTy.eChrEmptyEr:      Push OA, .ChrEmpty.ShwFld.Adr:      Push OW, eWrkWs
            Case eDtaErTy.eChrValEr:        Push OA, .ChrVal.ShwFld.Adr:        Push OW, eWrkWs
            Case eDtaErTy.eDifHdCellEr:     Push OA, .DifHdCell.ShwFld.Adr:     Push OW, eWrkWs
            Case eDtaErTy.eDifColCntEr
            Case eDtaErTy.eDifR1FormulaEr:  Push OA, .DifR1Formula.ShwFld.Adr:  Push OW, Enm.WhichWs(.DifR1Formula.ShwFld.Ws)
            Case eDtaErTy.eDifValEr:        Push OA, .DifVal.ShwFld.Adr:        Push OW, eWrkWs
            Case eDtaErTy.eDupSkuEr:        Push OA, .DupSku.ShwFld.Adr:        Push OW, Enm.WhichWs(.DupSku.ShwFld.Ws)
            Case eDtaErTy.eNoOrgRowEr:      Push OA, .NoOrgRow.ShwFld.Adr:      Push OW, eWrkWs
            Case eDtaErTy.eValTyEr:         Push OA, .ValTy.ShwFld.Adr:         Push OW, Enm.WhichWs(.ValTy.ShwFld.Ws)
            Case Else: Stop
            End Select
        End With
    Next
Dim O As RedAdr
    O.WhichWs = OW
    O.Adr = OA
RedAdr = O
End Property

Private Sub ZShwMsg(AnyEr As Boolean)
If AnyEr Then
    MsgBox "There are errors in the mass update excel file.  Please make correction and re-validate again", vbCritical:
Else
    MsgBox "There is no error, Worksheet(DataChanged) is generated", vbInformation
End If
End Sub
