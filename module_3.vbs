Sub main()

'Variable declaration
Dim work_sheet As Worksheet

'For loop to iterate over every sheet and
'call the modules to get the stock yearly overview
For Each work_sheet In Worksheets
    work_sheet.Activate
    Call module_1.yearly_overview
    Call module_2.Greatest
Next

End Sub