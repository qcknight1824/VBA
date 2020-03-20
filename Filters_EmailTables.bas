Attribute VB_Name = "Filters_EmailTables"
Sub Clear_Filter()
'
' Clear_Filter Macro
' clears filters on Position Data tab
'
' Keyboard Shortcut: Ctrl+p
'
On Error Resume Next

Sheets("Raw NII Data").ShowAllData

End Sub

Sub New_Hedge()
'
'Show_Buys_Only macro
'Filters Position Data to see only new trades

Sheets("Raw NII Data").Range("A5:AX5").AutoFilter Field:=50, Criteria1:="New Trade"

End Sub

Sub De_Designation()
'
'Show_Buys_Only macro
'Filters Position Data to see only de-designations

Sheets("Raw NII Data").Range("A5:AX5").AutoFilter Field:=50, Criteria1:="=*de-designation*"

End Sub

Sub Email_Tables()

Sheets("Email Tables").Range("A4:D1000").Clear
Sheets("Email Tables").Range("A3").ClearContents

Sheets("Email Tables").Range("F4:I1000").Clear
Sheets("Email Tables").Range("F3").ClearContents

If Sheets("Raw NII Data").Range("New_Trade_Count") > 1 Then
Sheets("Email Tables").Range("A3:D3").AutoFill Destination:=Sheets("Email Tables").Range("A3:D3").Resize(Sheets("Raw NII Data").Range("New_Trade_Count").Value)
If Sheets("Raw NII Data").Range("New_Trade_Count") = 1 Then
GoTo new_trade_paste
If Sheets("Raw NII Data").Range("De_Designation_Count") = 0 Then
GoTo De_Designations
End If
End If
End If

GoTo new_trade_paste

new_trade_paste: Sheets("Raw NII Data").Range("A5:AX5").AutoFilter Field:=50, Criteria1:="New Trade"

    Sheets("Raw NII Data").Range("B6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Email Tables").Range("A3").PasteSpecial xlPasteValues
    
De_Designations: Sheets("Raw NII Data").ShowAllData

If Sheets("Raw NII Data").Range("De_Designation_Count") > 1 Then
Sheets("Email Tables").Range("F3:I3").AutoFill Destination:=Sheets("Email Tables").Range("F3:I3").Resize(Sheets("Raw NII Data").Range("De_Designation_Count").Value)
If Sheets("Raw NII Data").Range("De_Designation_Count") = 1 Then
GoTo Continue
If Sheets("Raw NII Data").Range("De_Designation_Count") = 0 Then
Exit Sub
End If
End If
End If

GoTo Continue

Continue: Sheets("Raw NII Data").Range("A5:AX5").AutoFilter Field:=50, Criteria1:="=*de-designation*"

    Sheets("Raw NII Data").Range("B6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Email Tables").Range("F3").PasteSpecial xlPasteValues
    
Sheets("Raw NII Data").ShowAllData

End Sub
