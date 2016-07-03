Attribute VB_Name = "DoDltLastImpQDte"
Option Explicit

Private A_Db As Database

Sub Do_Dlt_LastQDtaPj()
Set A_Db = Dao.DBEngine.OpenDatabase(Fs_Fb)
Dim LastQDte As Date
If Not ZConfirm(LastQDte) Then Exit Sub
' Must in this dependence order
A_Db.Execute Fmt("Delete from SkuCostChr      where QuoteDate=#{0}#", LastQDte)
A_Db.Execute Fmt("Delete from SkuCostEle      where QuoteDate=#{0}#", LastQDte)
A_Db.Execute Fmt("Delete from Sku             where QuoteDate=#{0}#", LastQDte)
A_Db.Execute Fmt("Delete from ProjOneTimeCost where QuoteDate=#{0}#", LastQDte)
A_Db.Execute Fmt("Delete from ProjQ           where QuoteDate=#{0}#", LastQDte)
A_Db.Close
End Sub

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()

End Sub

Private Function ZConfirm(ByRef OLastQDte) As Boolean
Dim MaxQDte As Date
Dim M$
Dim NSku%
    MaxQDte = RunSql_Val(A_Db, "Select Max(QuoteDate) from ProjQ")
    M = Format(MaxQDte, "yyyy-mm-dd")
    NSku = RunSql_Val(A_Db, Fmt_QQ("Select Count(*) from ProjQ where QuoteDate=#?#", M))

Dim A$
A = InputBox("Input [YES] to confirm delete" & vbLf & vbLf & "[" & NSku & "] Sku of Quote Date[" & M & "]", "Delete Project?")
If A <> "YES" Then Exit Function
A = InputBox("Input [YES] again to confirm delete" & vbLf & vbLf & "[" & NSku & "] Sku of Quote Date[" & M & "]", "Confirm Delete Project?")
If A <> "YES" Then Exit Function
OLastQDte = MaxQDte
ZConfirm = True
End Function
