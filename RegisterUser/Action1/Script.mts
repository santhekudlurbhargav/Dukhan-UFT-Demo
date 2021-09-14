strSheetName = "Pre_Login"
DataTable.AddSheet strSheetName
DataTable.ImportSheet "C:\Users\Test\Desktop\Dukhan_Bank Web Application\Test Data\DataSheet.xlsx",strSheetName,strSheetName

StrCardNumber = DataTable.Value("CARD_NUMBER",strSheetName)
StrPassword = DataTable.Value("PIN_NUMBER",strSheetName)
 @@ script infofile_;_ZIP::ssf37.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").Link("Register").Click @@ script infofile_;_ZIP::ssf50.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("debitCard").Set StrCardNumber @@ script infofile_;_ZIP::ssf52.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("pin_number").SetSecure StrPassword @@ script infofile_;_ZIP::ssf54.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebCheckBox("WebCheckBox").Set "ON" @@ script infofile_;_ZIP::ssf56.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").Frame("a-t6c32rikh1hu").WebCheckBox("I'm not a robot").Set "ON" @@ script infofile_;_ZIP::ssf31.xml_;_
wait(3)
Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Next").Click @@ script infofile_;_ZIP::ssf61.xml_;_
wait(3)
Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Ok").Click @@ script infofile_;_ZIP::ssf63.xml_;_
