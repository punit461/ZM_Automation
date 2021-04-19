# ZM_Automation_Excel

This Automation tool has a GUI interface .
Take up this package and firstly add a chrome canary build in the same folder or install and update the application link in the program( linklogic.py)
Every time you have to run linkui.py and only then the program will fucntion properly.
after clicking the linkui.py you will be welcomed with a GUI asking for whether you want a chrome to run in background or see the window.
If you wish to see the automation functioning click yes and if no then click no.
but if make sure the path of chrome canary is defined without any mistake 
Next It'll ask for Enter the Data. and As soon as we click the button It'll erase all the data from the LinkData.xlsx and open a fresh new sheet where you can enter the data 
then save and close the data file.
and click run.

Points to remeber where error might occur.
1.Chrome Canary Location
2. Username/ password in linklogic.py
3. Entering the Data in Excel and not saving and closing the file. #close the file after the data input and can reopen after the operation.
4. Chrome Driver Integratrion.
5. Entering Float value in the LinkData.xlsx due to excel converts data.( #bug)
