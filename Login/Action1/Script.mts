
For i = 1 To 2
	
strSheetName = "Pre_Login"
DataTable.AddSheet strSheetName
DataTable.ImportSheet "C:\Users\Test\Desktop\Dukhan_Bank WebApplication\Test Data\DataSheet.xlsx",strSheetName,strSheetName
Datatable.SetCurrentRow(i)
StrUsername = DataTable.Value("USER_NAME",strSheetName)
StrPassword = DataTable.Value("PASSWORD",strSheetName)

Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Login").Click @@ script infofile_;_ZIP::ssf1.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("WebEdit").Set StrUsername @@ script infofile_;_ZIP::ssf2.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("WebEdit_2").SetSecure StrPassword @@ script infofile_;_ZIP::ssf3.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Login").Click @@ script infofile_;_ZIP::ssf4.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Ok").Click @@ script infofile_;_ZIP::ssf5.xml_;_

Next
