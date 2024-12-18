Set X = WScript.CreateObject("WScript.Shell")
Do
    WScript.Sleep 50
    X.SendKeys "{CAPSLOCK}"
    X.SendKeys "{NUMLOCK}"
    X.SendKeys "sunil"
    X.SendKeys "{SCROLLLOCK}"
Loop
