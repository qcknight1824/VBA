Attribute VB_Name = "ATX_Runs"
Sub ATX_Macro_Button()

Sheets("TRADES Query").Range("H11:U1000").Clear

   Dim AV_TRADE As New ATX.QueryReportFunctions
   AV_TRADE.Run Sheets("Trades Query").Range("I7")
   
If Sheets("Trades Query").Range("I1") = 0 Then
    Sheets("ATXAnserPA").Range("B8:AY1000").Clear
    Sheets("RAW NII DATA").Range("A7:AX1000").Clear
    UserForm2.Show
    Exit Sub
    End If
    
    
Sheets("Positions Query").Range("H11:U1000").Clear

   Dim AV_POSITION As New ATX.QueryReportFunctions
   AV_POSITION.Run Sheets("Positions Query").Range("I7")
   
   'RAW NII DATA Tab

   Sheets("RAW NII DATA").Range("A7:Ax1000").Clear
   
    Sheets("ATXAnserPA").Range("B8:AY1000").Clear
    
    
If Sheets("Trades Query").Range("I1") > 1 Then
   Sheets("RAW NII DATA").Range("A6:Ax6").AutoFill Destination:=Sheets("RAW NII DATA").Range("A6:Ax6").Resize(Sheets("Trades QUERY").Range("TRADE_COUNT").Value)

' TRUNCATE AND DRAG FORMULAS IN ATX TABLE

   Sheets("ATXAnserPA").Range("B7:AY7").AutoFill Destination:=Sheets("ATXAnserPA").Range("B7:AY7").Resize(Sheets("Trades QUERY").Range("TRADE_COUNT").Value)
    Exit Sub
    End If

End Sub
