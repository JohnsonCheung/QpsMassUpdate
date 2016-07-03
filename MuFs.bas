Attribute VB_Name = "MuFs"
Option Explicit
Public Const Fs_ImpFxLikStr$ = "Mass Update Data - *.xlsx"

Property Get Fs_CfgFdr$()
If Fs_IsSingleFdr Then
    Fs_CfgFdr = Fs_Hom
Else
    Fs_CfgFdr = Fs_Hom & "Generate QPS\"
End If
End Property

Property Get Fs_CfgFx$()
Const Lik = "*(Cfg)*.xlsm"
Dim Pth$
    Pth = Fs_CfgFdr
Dim Ay$()
    Ay = Pth_FnAy(Pth, Lik)
    
If Sz(Ay) = 0 Then Er "There is no {Like} file in {folder}", Lik, Pth

If Sz(Ay) > 1 Then
    Er "There are {n} [*(Cfg)*.xlsm] file in {folder}.  At most it can only be one Cfg-file", Sz(Ay), Pth
End If
Fs_CfgFx = Fs_CfgFdr & Ay(0)
End Property

Property Get Fs_Fb$()
Const Fn = "QPS Costing(Data).accdb"
Dim O$
O = Fs_FbFdr & Fn
If Dir(O) = "" Then Er "Cost {database} not exist in {folder}", Fn, Fs_FbFdr
Fs_Fb = O
End Property

Property Get Fs_FbFdr$()
If Fs_IsSingleFdr Then
    Fs_FbFdr = Fs_Hom
Else
    Fs_FbFdr = Fs_QpsImpExpFdr & "Data\"
End If
End Property

Property Get Fs_ImpDoneFdr$()
Dim O$
    O = Fs_ImpFdr & "Done\"

Pth_CrtIfNotExist O
Fs_ImpDoneFdr = O
End Property

Property Get Fs_ImpFdr$()
If Fs_IsSingleFdr Then
    Fs_ImpFdr = Fs_Hom
Else
    Fs_ImpFdr = Fs_QpsImpExpFdr & "Import - Mass Update data\"
End If
End Property

Function Fs_ImpFx$()
Dim FnAy$()
    FnAy = Pth_FnAy(Fs_ImpFdr, Fs_ImpFxLikStr)
Select Case Sz(FnAy)
    Case 0: Er "No files [Mass Update Data - *.xlsx] in {folder}", Fs_ImpFdr
    Case 1: Fs_ImpFx = Fs_ImpFdr & FnAy(0): Exit Function
End Select
Dim A$, Ay$(), J%
ReDim Ay(UB(FnAy))
For J = 0 To UB(FnAy)
    Ay(J) = J + 1 & ". " & FnAy(J)
Next
A = Join(Ay, vbLf)
Dim I$
Again:
    I = Val(InputBox("Select 1-" & Sz(FnAy) & vbLf & vbLf & A, "Which file?", 1))
    If I = 0 Then Err.Raise 1, , "No file is choosen"
    If 1 <= I And I <= Sz(FnAy) Then
        Fs_ImpFx = Fs_ImpFdr & FnAy(I - 1): Exit Function
    End If
    DoEvents
    GoTo Again
End Function

Property Get Fs_QpsImpExpFdr$()
If Fs_IsSingleFdr Then
    Fs_QpsImpExpFdr = Fs_Hom
Else
    Fs_QpsImpExpFdr = Fs_Hom & "QPS Costing ImportExport\"
End If
End Property

Private Property Get Fs_Hom$()
If Fs_IsSingleFdr Then
    Fs_Hom = Wb_Pth(yWbMassUpd)
Else
    Fs_Hom = Pth_Normalize(Wb_Pth(yWbMassUpd) & "..\..\..\")
End If
End Property

Private Property Get Fs_IsSingleFdr() As Boolean
Fs_IsSingleFdr = Not Str_IsSfx(Wb_Pth(yWbMassUpd), "\QPS Costing ImportExport\Program\MassUpdate\")
End Property
