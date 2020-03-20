Attribute VB_Name = "ATX_Macro_Button"
Sub ATX_Macro_Button()
Attribute ATX_Macro_Button.VB_ProcData.VB_Invoke_Func = "V\n14"
'
' Keyboard Shortcut: Ctrl+Shift+V
'

Sheets("TRADES").Range("H11:AH1000").Clear

    Dim AV_TRADE As New ATX.QueryReportFunctions
    AV_TRADE.Run Sheets("Trades").Range("I7")
    
If Sheets("TRADES").Range("I1") = 1 Then
    Sheets("POSITION DATA").Range("A6:X1000").Clear
    Sheets("TRADING_ACTIVITY").Range("A11:Q1000").Clear
    Sheets("Bloomberg Pull").Range("A6:AL1000").Clear
    UserForm2.Show
Exit Sub
End If
    

' TRUNCATE AND DRAG FORMULAS IN TABLE

Sheets("POSITION DATA").Range("A6:X1000").Clear

Sheets("POSITION DATA").Range("A5:X5").AutoFill Destination:=Sheets("POSITION DATA").Range("A5:X5").Resize(Sheets("TRADES").Range("TRADE_COUNT").Value - 1)

'Trading Activity Tab

Sheets("TRADING_ACTIVITY").Range("A11:Q1000").Clear

Sheets("TRADING_ACTIVITY").Range("A10:Q10").AutoFill Destination:=Sheets("TRADING_ACTIVITY").Range("A10:Q10").Resize(Sheets("TRADES").Range("TRADE_COUNT").Value)

Sheets("TRADING_ACTIVITY").Range("A" & Range("TRADE_COUNT").Value + 9 & ":Q" & Range("TRADE_COUNT").Value + 9).Clear

'Bloomberg Pull

Sheets("Bloomberg Pull").Range("A6:AL1000").Clear

Sheets("Bloomberg Pull").Range("A5:AL5").AutoFill Destination:=Sheets("Bloomberg Pull").Range("A5:AL5").Resize(Sheets("TRADES").Range("TRADE_COUNT").Value - 1)



End Sub
