strSheetName = "Pre_Login"
DataTable.AddSheet strSheetName
DataTable.ImportSheet "C:\Users\Test\Desktop\Dukhan_Bank WebApplication\Test Data\DataSheet.xlsx",strSheetName,strSheetName

StrCardNumber = DataTable.Value("CARD_NUMBER",strSheetName)
StrPassword = DataTable.Value("PIN_NUMBER",strSheetName) @@ script infofile_;_ZIP::ssf4.xml_;_

Browser("Dukhan Bank").Page("Dukhan Bank").Link("Register").Click @@ script infofile_;_ZIP::ssf21.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("debitCard").Set StrCardNumber @@ script infofile_;_ZIP::ssf24.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("pin_number").SetSecure StrPassword @@ script infofile_;_ZIP::ssf26.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebCheckBox("WebCheckBox").Set "ON" @@ script infofile_;_ZIP::ssf28.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").Frame("a-t6c32rikh1hu").WebCheckBox("I'm not a robot").Set "ON" @@ script infofile_;_ZIP::ssf31.xml_;_
wait(3)
Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Next").Click @@ hightlight id_;_3343150_;_script infofile_;_ZIP::ssf36.xml_;_
wait(3)
Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Ok").Click @@ script infofile_;_ZIP::ssf11.xml_;_
