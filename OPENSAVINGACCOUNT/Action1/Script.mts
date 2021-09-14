
strSheetName = "Pre_Login"
DataTable.AddSheet strSheetName
DataTable.ImportSheet "C:\Users\Test\Desktop\Dukhan_Bank Web Application\Test Data\DataSheet.xlsx",strSheetName,strSheetName

TotalRows = Datatable.GetSheet(strSheetName).GetCurrentRow
 @@ hightlight id_;_6228424_;_script infofile_;_ZIP::ssf297.xml_;_
strFirstname = DataTable.Value("FIRST_NAME",strSheetName) @@ hightlight id_;_1837322_;_script infofile_;_ZIP::ssf224.xml_;_
strMiddlename = DataTable.Value("MIDDLE_NAME",strSheetName)
strLastname = DataTable.Value("LAST_NAME",strSheetName)
strEmail = DataTable.Value("EMAIL",strSheetName)
strQID = DataTable.Value("QID",strSheetName)
strPassport = DataTable.Value("PASSPORT",strSheetName)
strmobileNumber  = DataTable.Value("MOBILE_NUMBER",strSheetName)
strCompany = DataTable.Value("COMPANY_NAME",strSheetName)
strAddress = DataTable.Value("ADDRESS",strSheetName)
'strAccountType = Datatable.Value("Account_Type",strSheetName)
wait(5)
Browser("Dukhan Bank").Page("Dukhan Bank").Link("Open an Account").Click @@ script infofile_;_ZIP::ssf302.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("Saving").Click @@ script infofile_;_ZIP::ssf304.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("Faseel High Profit Savings").Click @@ script infofile_;_ZIP::ssf306.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebCheckBox("WebCheckBox").Set "ON" @@ script infofile_;_ZIP::ssf308.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("QAR 10,000 - QAR 14,999").Click @@ script infofile_;_ZIP::ssf310.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("▼Branch*Grand Hamad Street").Click @@ script infofile_;_ZIP::ssf312.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("Al Sadd Branch").Click @@ script infofile_;_ZIP::ssf314.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Next").Click @@ script infofile_;_ZIP::ssf316.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("firstName").Set strFirstname @@ script infofile_;_ZIP::ssf318.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("middleName").Set strMiddlename @@ script infofile_;_ZIP::ssf320.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("Last Name*").Click @@ script infofile_;_ZIP::ssf322.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("lastName").Set strLastname @@ script infofile_;_ZIP::ssf324.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("Email*").Click @@ script infofile_;_ZIP::ssf326.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("email").Set strEmail @@ script infofile_;_ZIP::ssf328.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("QID").Click @@ script infofile_;_ZIP::ssf330.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("qid").Set strQID @@ script infofile_;_ZIP::ssf332.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("Passport*").Click @@ script infofile_;_ZIP::ssf334.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("passport").Set strPassport @@ script infofile_;_ZIP::ssf336.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("mobileNumber").Set strmobileNumber @@ script infofile_;_ZIP::ssf338.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("Company Name*").Click @@ script infofile_;_ZIP::ssf340.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("companyName").Set strCompany @@ script infofile_;_ZIP::ssf342.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("address").Set strAddress @@ script infofile_;_ZIP::ssf344.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Next_2").Click @@ script infofile_;_ZIP::ssf346.xml_;_
wait(5)

setting.webpackage("ReplayType") = 2
Browser("Dukhan Bank").Page("Dukhan Bank").WebElement("WebElement").Click 10,20 @@ script infofile_;_ZIP::ssf348.xml_;_
setting.webpackage("ReplayType") = 1
Window("Google Chrome").Dialog("Open").WinObject("Items View").WinList("Items View").Select "Passport" @@ hightlight id_;_5702066_;_script infofile_;_ZIP::ssf352.xml_;_
Window("Google Chrome").Dialog("Open").WinButton("Open").Click @@ hightlight id_;_2165236_;_script infofile_;_ZIP::ssf353.xml_;_
wait(5)

Browser("Dukhan Bank").Page("Dukhan Bank").WebCheckBox("WebCheckBox").Set "ON" @@ script infofile_;_ZIP::ssf283.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").Frame("a-o8w7a89ix91").WebCheckBox("I'm not a robot").Set "ON" @@ script infofile_;_ZIP::ssf360.xml_;_

If Browser("Dukhan Bank").Page("Dukhan Bank").Frame("c-exednz337q5i").WebTable("WebTable").Exist(3) Then @@ script infofile_;_ZIP::ssf362.xml_;_
    Reporter.ReportEvent micFail, "Capcha pop up will be displayed","Submit button is disabled we can't select capcha images"
    ElseIf Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Submit").Exist(2) Then @@ script infofile_;_ZIP::ssf363.xml_;_
	       Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Submit").Click @@ script infofile_;_ZIP::ssf365.xml_;_
End If
           Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Ok").Click @@ script infofile_;_ZIP::ssf367.xml_;_
           
           
'If not Browser("Dukhan Bank").Page("Dukhan Bank").Frame("c-z1hlb1dtq8zz").exist(4) = True Then
'          Reporter.ReportEvent micFail, "Capcha pop up will be displayed","Submit button is disabled we can't select capcha images"
'	ElseIf Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Submit").Click Then
'	           Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Ok").Click
'End If
'
'If not Browser("Dukhan Bank").Page("Dukhan Bank").Frame("c-z1hlb1dtq8zz").Exist(5) Then @@ script infofile_;_ZIP::ssf136.xml_;_
'    Reporter.ReportEvent micFail, "Capcha pop up will be displayed","we can't select capcha images"
'ElseIf Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Submit").Click then
'Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Ok").Click @@ script infofile_;_ZIP::ssf60.xml_;_
'End If @@ script infofile_;_ZIP::ssf142.xml_;_
