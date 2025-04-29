PROMPT USER TO OPEN EXCEL FILES
```vba
Sub browseFileToOpen()

    Dim myFile As Variant
    Dim CD1 As String
    
'Open the QSOERP Raw Data entry
    Workbooks.Open Filename:="\\Operation\QSOERP.xlsx"

'Restrict browsing other files - Only excel files
    CD1 = InputBox("Enter Calndar date of the data in the format M/D/YYYY:")
    MsgBox "Please Select the ERP Raw Data File (APE) to open"
    myFile = Application.GetOpenFilename("Excel Files (*.xl*),*.xl*", , "Choose File", "Open", False)
    If myFile = False Then
'Do nothing
    Else
    Workbooks.Open (myFile)
    End If

'Copy Raw Data to paste
    Range("A6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A6:A110").Select
    Selection.Copy
    Windows("QSOERP.xlsx").Activate
    Sheets("TB").Select
    lastrow = Cells(Rows.Count, "C").End(xlUp).Row
    Cells(lastrow + 1, "C").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    SendKeys ("%{TAB}"), True
    
    Range("F6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("F6:F110").Select
    Selection.Copy
    Windows("QSOERP.xlsx").Activate
    Cells(lastrow + 1, "D").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Cells(rowIn, "A") = CD1
    Selection.Copy
   Cells(lastrow + 1, "A").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
End Sub


```
