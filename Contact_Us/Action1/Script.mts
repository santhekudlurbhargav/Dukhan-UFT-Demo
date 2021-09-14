
strSheetName = "Pre_Login"
DataTable.AddSheet strSheetName
DataTable.ImportSheet "C:\Users\Test\Desktop\Dukhan_Bank Web Application\Test Data\DataSheet.xlsx",strSheetName,strSheetName

Datatable.SetCurrentRow(i)
StrFullName = DataTable.Value("FULL_NAME",strSheetName)
StrQID = DataTable.Value("QID",strSheetName) @@ script infofile_;_ZIP::ssf28.xml_;_
StrEmail = DataTable.Value("EMAIL",strSheetName)
StrMobileNumber = DataTable.Value("MOBILE_NUMBER",strSheetName)
StrReason = DataTable.Value("REASON",strSheetName)

Browser("Dukhan Bank").Page("Dukhan Bank").Link("Contact Us").Click @@ script infofile_;_ZIP::ssf99.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("full name*").Click @@ script infofile_;_ZIP::ssf101.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("WebEdit").Set StrFullName @@ script infofile_;_ZIP::ssf103.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("WebEdit_2").Set StrQID @@ script infofile_;_ZIP::ssf105.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("Email*").Click @@ script infofile_;_ZIP::ssf107.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("WebEdit_3").Set StrEmail @@ script infofile_;_ZIP::ssf109.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("WebEdit_4").Set StrReason @@ script infofile_;_ZIP::ssf111.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("12/09/2021").Click @@ script infofile_;_ZIP::ssf113.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("12/09/2021").Click @@ script infofile_;_ZIP::ssf115.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("3").Click @@ script infofile_;_ZIP::ssf117.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("30").Click @@ script infofile_;_ZIP::ssf119.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("WebEdit_5").Set "yejksjdjk" @@ script infofile_;_ZIP::ssf78.xml_;_
wait(3)
Browser("Dukhan Bank").Page("Dukhan Bank").Frame("a-3xauvmv0gayn").WebCheckBox("I'm not a robot_2").Set "ON"

'Browser("Dukhan Bank").Page("Dukhan Bank").Frame("a-3xauvmv0gayn").WebCheckBox("I'm not a robot").Set "ON" @@ script infofile_;_ZIP::ssf86.xml_;_
wait(3)
Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Submit").Highlight
wait(3)
Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Submit").Click @@ script infofile_;_ZIP::ssf83.xml_;_

Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Ok").Click @@ script infofile_;_ZIP::ssf85.xml_;_
