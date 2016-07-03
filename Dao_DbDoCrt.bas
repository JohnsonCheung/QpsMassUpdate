Attribute VB_Name = "Dao_DbDoCrt"
Option Explicit
Private A_D As DbDef
Private A_Db As Database

Sub Db_DoCrt(D As DbDef)
ZFb_DoAssert
ZFb_DoCrt
ZTbl_DoCrt
ZRel_DoCrt
ZFb_DoClose
End Sub

Private Property Get ZFb$()

End Property

Private Sub ZFb_DoAssert()

End Sub

Private Sub ZFb_DoClose()
A_Db.Close
End Sub

Private Sub ZFb_DoCrt()
Set A_Db = CreateDatabase(ZFb, ZFb_Locale)
End Sub

Private Property Get ZFb_Locale$()

End Property

Private Property Get ZRel_Def() As RelDef()

End Property

Private Sub ZRel_DoCrt()
Dim R() As RelDef
    R = ZRel_Def
Dim J%
For J = 0 To UBound(R)
    ZRel_DoCrt_One R(J)
Next
End Sub

Private Sub ZRel_DoCrt_One(R As RelDef)

End Sub

Private Sub ZTbl_DoCrt()
Dim T() As TblDef
    T = A_D.Tbl
Dim J%
For J = 0 To UBound(T)
    Tbl_DoCrt A_Db, T(J)
Next
End Sub
