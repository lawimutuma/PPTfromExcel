# PPTfromExcel
This is simple VBA code to populate data into Powerpoint from an Excel table.

The code picks data from Excel - in this case, names from column B and phone numbers from column C. 

It loops through the names and updates a powerpoint slide with the details from Excel.

# Prerequisites:
- You should have designed your base powerpoint beforehand. This will be used as the template for the new ones to be created. Save it as a pptx.
- You should know the textbox/shape names for the items you are going update in PPT. (Use the HOME -> Editing tab -> Select -> Selection Pane to list your object names)
- You should have an excel with the data you want to copy across - obviously

# Steps
- Create a new powerpoint to run the VBA code. Don't use the base powerpoint.
-  Insert a VBA module module into the new ppt (ALT + F11 on Windows)
-  In the VBA Window, click Tools -> References and confirm the following are selected, if not, add them. (There will be others pre-selected by default.)
  -     i) "Microsoft Office xx.x Object Library" (where xx.x represents the version of Excel you're using)
  -      ii) "Microsoft Excel xx.x Object Library" (where xx.x represents the version of Excel you're using)
-  Paste the code into the module window. Amend variables where necessary (folder names, file names, etc)
-  Save the file as a PowerPoint Macro-enabled Presentation (*.pptm)
-  Make sure the source excel file is not open otherwise you'll get a file open error
-  Run your code by pressing F5
-  If you get errors, remember to force closure of excel - using task manager - before rerunning the code

Good luck
