Sub PopulateAndSaveSlides()
    Dim pptApp As Object
    Dim pptBasePresentation As Object
    Dim pptSlide As Object
    Dim pptTextbox As Object
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim nameCell As Object
    Dim phoneCell As Object
    Dim newName As String
    Dim newPhone As String
    
    'In the "References" window, find and check the checkbox for "Microsoft Excel xx.x Object Library" (where xx.x represents the version of Excel you're using).
    
    ' Set base folder
    Dim baseFolder As String
    baseFolder = "C:\Users\User\MyFiles" ' Update as necessary
    
    ' Open Excel file
    Set excelApp = CreateObject("Excel.Application")
    inviteListExcel = baseFolder & "\InviteList_Guest.xlsx"
    Set excelWorkbook = excelApp.Workbooks.Open(inviteListExcel) ' Update the path
    
    ' Open base PowerPoint file
    Set pptApp = CreateObject("PowerPoint.Application")
    Set pptBasePresentation = pptApp.Presentations.Open(baseFolder & "\Invite_Guest.pptx") ' Update the path
    
    ' Loop through each name and phone number in Excel
    Dim lastRow As Long
    lastRow = excelWorkbook.Sheets("Sheet1").Cells(excelWorkbook.Sheets("Sheet1").Rows.Count, 2).End(-4162).Row ' Find the last used row in column B
    For rownum = 2 To lastRow ' Start from row 2 since row 1 is header
        newName = excelWorkbook.Sheets("Sheet1").Cells(rownum, 2).Value ' Column B contains names
        newPhone = excelWorkbook.Sheets("Sheet1").Cells(rownum, 3).Value ' Column C contains phone numbers

        If newName <> "" Then ' I'm using the name as the base for filename. So it cannot be blank.
            ' Clone the base presentation to a new presentation
            Set pptSlide = pptBasePresentation.Slides(1)
            Set pptTextbox = pptSlide.Shapes("Rectangle 10").TextFrame.TextRange ' Get the placeholder textbox name from the base Powerpoint
            pptTextbox.Text = newName & vbCrLf & newPhone 'Update as per your requirements/preference
            
            ' pptBasePresentation.SaveCopyAs baseFolder & "\" & newName & ".pptx" ' Update the path ' Option to create/save as ppt. Uncomment if you want to create a new ppt from the changes
            
            ' Print the slide to PDF using the printer settings
            pptBasePresentation.ExportAsFixedFormat Path:=baseFolder & "\" & newName & ".pdf", FixedFormatType:=ppFixedFormatTypePDF, PrintRange:=Nothing
        End If
    Next rownum
    
    ' Close Excel and PowerPoint
    excelWorkbook.Close False
    excelApp.Quit
    pptBasePresentation.Close
    pptApp.Quit
End Sub

