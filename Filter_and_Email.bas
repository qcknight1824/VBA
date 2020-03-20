Attribute VB_Name = "Filter_and_Email"
Sub Clear_Filter()
Attribute Clear_Filter.VB_Description = "clears filters on Position Data tab"
Attribute Clear_Filter.VB_ProcData.VB_Invoke_Func = "p\n14"
'
' Clear_Filter Macro
' clears filters on Position Data tab
'
' Keyboard Shortcut: Ctrl+p
'
On Error Resume Next

Sheets("position data").ShowAllData

End Sub

Sub Show_Buys_Only()
'
'Show_Buys_Only macro
'Filters Position Data to see only buys

Range("A4:W4").AutoFilter Field:=4, Criteria1:="Buy"

End Sub

Sub Show_Sells_Only()
'
'Show_Buys_Only macro
'Filters Position Data to see only buys

Range("A4:W4").AutoFilter Field:=4, Criteria1:="Sell"

End Sub



Sub DRAFTEMAIL_2()

draft_portfolio.Show

Dim OApp As Object, OMail As Object, signature As String
Set OApp = CreateObject("Outlook.Application")
Set OMail = OApp.CreateItem(0)
    With OMail
    .Display
    End With
        signature = OMail.Body
    With OMail
    .To = Sheets("Email Draft").Range("A2").Value
    .CC = Sheets("Email Draft").Range("B2").Value
    .Subject = Sheets("Email Draft").Range("C2").Value
    '.Attachments.Add
    .Body = Sheets("Email Draft").Range("D2").Value
    
    End With
           
End Sub
