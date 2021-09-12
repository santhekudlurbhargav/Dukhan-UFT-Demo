strSheetName = "Pre_Login"
DataTable.AddSheet strSheetName
DataTable.ImportSheet "C:\Users\Test\Desktop\Dukhan_Bank WebApplication\Test Data\DataSheet.xlsx",strSheetName,strSheetName

strFirstname = DataTable.Value("FIRST_NAME",strSheetName)
strMiddlename = DataTable.Value("MIDDLE_NAME",strSheetName)
strLastname = DataTable.Value("LAST_NAME",strSheetName)
strEmail = DataTable.Value("EMAIL",strSheetName)
strQID = DataTable.Value("QID",strSheetName)
strPassport = DataTable.Value("PASSPORT",strSheetName)
strmobileNumber  = DataTable.Value("MOBILE_NUMBER",strSheetName)
strCompany = DataTable.Value("COMPANY_NAME",strSheetName)
strAddress = DataTable.Value("ADDRESS",strSheetName)
wait(5)
'Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("WebEdit").Highlight

setting.WebPackage("ReplayType") = 2
Browser("Dukhan Bank").Page("Dukhan Bank").Link("Open an Account").Click 20,20 @@ script infofile_;_ZIP::ssf154.xml_;_
setting.WebPackage("ReplayType") = 1
'Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("Current").Highlight
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("Current").Click @@ script infofile_;_ZIP::ssf155.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEement("html tag:=SPAN","innertext:=Saving").Click
Browser("Dukhan Bank").Page("Dukhan Bank").WebCheckBox("WebCheckBox").Set "ON" @@ script infofile_;_ZIP::ssf156.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("QAR 10,000 - QAR 14,999").Click @@ script infofile_;_ZIP::ssf157.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("Grand Hamad Street Branch").Click @@ script infofile_;_ZIP::ssf158.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Next").Click @@ script infofile_;_ZIP::ssf159.xml_;_

Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("firstName").Set strFirstname @@ script infofile_;_ZIP::ssf160.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("middleName").Set strMiddlename @@ script infofile_;_ZIP::ssf161.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("Last Name*").Click @@ script infofile_;_ZIP::ssf162.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("lastName").Set strLastname @@ script infofile_;_ZIP::ssf163.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("Email*").Click @@ script infofile_;_ZIP::ssf164.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("email").Set strEmail @@ script infofile_;_ZIP::ssf165.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("QID").Click @@ script infofile_;_ZIP::ssf166.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("qid").Set strQID @@ script infofile_;_ZIP::ssf167.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("Passport*").Click @@ script infofile_;_ZIP::ssf168.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("passport").Set strPassport @@ script infofile_;_ZIP::ssf169.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("Mobile Number*").Click @@ script infofile_;_ZIP::ssf170.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("mobileNumber").Set strmobileNumber @@ script infofile_;_ZIP::ssf171.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("Company Name*").Click @@ script infofile_;_ZIP::ssf172.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("companyName").Set strCompany @@ script infofile_;_ZIP::ssf173.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("Account Type Customer").Click @@ script infofile_;_ZIP::ssf174.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("address").Set strAddress @@ script infofile_;_ZIP::ssf175.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Next_2").Click @@ script infofile_;_ZIP::ssf176.xml_;_
wait(2)
setting.WebPackage("ReplayType") = 2
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("WebElement").Click 10,10
setting.WebPackage("ReplayType") = 1
wait(2)
Window("Google Chrome").Dialog("Open").WinObject("Items View").WinList("Items View").Select "Salary Certificate" @@ hightlight id_;_1904353552_;_script infofile_;_ZIP::ssf177.xml_;_
wait(2)
Window("Google Chrome").Dialog("Open").WinButton("Open").Click @@ hightlight id_;_2229566_;_script infofile_;_ZIP::ssf178.xml_;_

setting.WebPackage("ReplayType") = 2
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("WebElement_2").Click 10,20
setting.WebPackage("ReplayType") = 1

Browser("Dukhan Bank").Page("Dukhan Bank").WebFile("WebFile_2").Set "QID Image Back.jpg"
Browser("Dukhan Bank").Page("Dukhan Bank").WebCheckBox("WebCheckBox").Set "ON" @@ script infofile_;_ZIP::ssf181.xml_;_
wait(5)
Browser("Dukhan Bank").Page("Dukhan Bank").Frame("Iamnotrobot").WebCheckBox("I'm not a robot").Click
wait(2)

If Browser("Dukhan Bank").Page("Dukhan Bank").Frame("c-btzhrgmf6rbk").WebButton("recaptcha-help-button").Exist(5) then
   Reporter.ReportEvent micFail, "Capcha pop up dispalyed", "Submit button is disabled mode"

ElseIf  Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Submit").Click then
 Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Ok").Click
End If 


'Browser("Dukhan Bank").Page("Dukhan Bank").Frame("a-gl05k379optl").WebCheckBox("I'm not a robot").Set "ON" @@ script infofile_;_ZIP::ssf182.xml_;_
'setting.WebPackage("ReplayType") = 2
' Browser("Dukhan Bank").InsightObject("InsightObject").Click
' Browser("Dukhan Bank").Page("Dukhan Bank").Frame("Iamnotrobot").WebCheckBox("I'm not a robot").Click 10,10
' setting.WebPackage("ReplayType") = 1 @@ script infofile_;_ZIP::ssf184.xml_;_
