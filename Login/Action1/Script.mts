Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Login").Click @@ script infofile_;_ZIP::ssf1.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("WebEdit").Set "Banny" @@ script infofile_;_ZIP::ssf2.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebEdit("WebEdit_2").SetSecure "61404fdfd3f6f2f3ce42cff0078d35c2e810954d2e5f" @@ script infofile_;_ZIP::ssf3.xml_;_
wait(3)
Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Login").Click @@ script infofile_;_ZIP::ssf4.xml_;_
Browser("Dukhan Bank").Page("Dukhan Bank").WebButton("Ok").Click @@ script infofile_;_ZIP::ssf5.xml_;_
