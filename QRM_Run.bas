Attribute VB_Name = "QRM_Run"


Sub RunQRM()
'Bloomberg Paste

Sheets("Bloomberg Paste").Range("A6:AL1000").Clear

BLOOMBERG_PULL = Sheets("Bloomberg Pull").Range("A5:AL" & Range("TRADE_COUNT").Value + 3).Value

Sheets("Bloomberg Paste").Range("A5:AL" & Range("TRADE_COUNT").Value + 3) = BLOOMBERG_PULL

Sheets("Bloomberg Paste").Range("AM6:AM1000").Clear

Sheets("Bloomberg Paste").Range("AM5").AutoFill Destination:=Sheets("Bloomberg Paste").Range("AM5").Resize(Sheets("TRADES").Range("TRADE_COUNT").Value - 1)

'QRM Linked Table

Sheets("QRM_Upload_Linked").Range("A3:AE1000").Clear

Sheets("QRM_Upload_Linked").Range("A2:AE2").AutoFill Destination:=Sheets("QRM_Upload_Linked").Range("A2:AE2").Resize(Sheets("TRADES").Range("TRADE_COUNT").Value - 1)


'QRM Upload

LINKED = Sheets("QRM_Upload_Linked").Range("A2:AE" & Range("TRADE_COUNT").Value + 2).Value

Sheets("QRM_UPLOAD").Range("B" & Range("PRIOR_VOL").Value + 2 & ":AF" & Range("PRIOR_VOL").Value + Range("TRADE_COUNT").Value) = LINKED

Range("PRIOR_VOL").Value = Range("CURRENT_VOL").Value


'COLOR FIRST ROW GRAY

'Workbooks("ATX Query IP - Data Pull.XLSB").Sheets("QRM_UPLOAD").Range("A" & Range("PRIOR_VOL").Value + 2 & ":AG" & Range("PRIOR_VOL").Value + 2).Select

'With Selection.Interior
      '  .Pattern = xlSolid
     '   .PatternColorIndex = xlAutomatic
      '  .ThemeColor = xlThemeColorDark1
    '    .TintAndShade = -0.149998474074526
    '    .PatternTintAndShade = 0
 '   End With
    
    
'ADD ROW FOR TRADE DATE ENTRIES

'LASTROW = Sheets("QRM_UPLOAD").Cells(Rows.Count, 2).End(xlUp).Row.Select

'Range ("A" & Range("PRIOR_VOL").Value + Range("TRADE_COUNT").Value + 1)


'TRANSFORM TRADE DATE ENTRIES
'COLOR TRADE DATE ENTRIES
'QRM ACCOUNT NAMES
'PULL DOWN COLUMN AH


End Sub
