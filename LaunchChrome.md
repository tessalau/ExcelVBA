``` vba

Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal dwData As Long, ByVal dwExtraInfo As Long)
Const MOUSEEVENTF_LEFTDOWN As Long = &H2
Const MOUSEEVENTF_LEFTUP As Long = &H4

Sub LoginToERPWithChromeV2()
    Dim chromePath As String
    Dim username As String
    Dim password As String
    Dim FinancialYear As String
    Dim FinancialPeriod As String
    Dim startTime As Single
    Dim detailsButtonX As Long
    Dim detailsButtonY As Long

    ' Get user credentials
    username = InputBox("Enter your ERP username:")
    password = InputBox("Enter your ERP password:")
    FinancialYear = InputBox("Enter the Financial Year for report extraction:")
    FinancialPeriod = InputBox("Enter the Financial Period for report extraction:")
    
   ' USER WARNING BOX
    MsgBox "Macro will run now. PLEASE DO NOT TOUCH THE MOUSE OR KEYBOARD WHILE MACRO IS RUNNING! To stop Macro, press ESC 2x"

    ' Path to Chrome executable
    chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
    

    ' Open Chrome and navigate to ERP
    Shell chromePath & " -url http://192.168.2.8:9080/TH6/index.xhtml", vbNormalFocus

    ' Wait for Chrome to open
    startTime = Timer
    Do While Timer < startTime + 3 ' Wait for 3 seconds
        DoEvents
    Loop
    
     ' Input username & Password
    SendKeys username, True
    SendKeys "{TAB}", True
    SendKeys password, True
    SendKeys "{TAB}{down}", True
    SendKeys "{ENTER}", True

    ' Wait for the Login to be ready
    startTime = Timer
    Do While Timer < startTime + 3 ' Wait for 3 seconds
        DoEvents
    Loop
    
    ' Click Operations > General Ledger > Trial Balance In Home Ccy (E2,TH5) > DL Report
    detailsButtonX = 160 ' Replace with actual X coordinate
    detailsButtonY = 185 ' Replace with actual Y coordinate
    SetCursorPos detailsButtonX, detailsButtonY
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    
     startTime = Timer
    Do While Timer < startTime + 2 ' Wait for 2 seconds
        DoEvents
    Loop
        
    detailsButtonX = 190 ' Replace with actual X coordinate
    detailsButtonY = 325 ' Replace with actual Y coordinate
    SetCursorPos detailsButtonX, detailsButtonY
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        
      startTime = Timer
    Do While Timer < startTime + 2 ' Wait for 2 seconds
        DoEvents
    Loop
        
    detailsButtonX = 1000  ' Replace with actual X coordinate
    detailsButtonY = 640 ' Replace with actual Y coordinate
    SetCursorPos detailsButtonX, detailsButtonY
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        
      startTime = Timer
    Do While Timer < startTime + 2 ' Wait for 2 seconds
        DoEvents
    Loop
    
    SendKeys "{TAB 5}{down 3}", True
    SendKeys "{TAB 3}", True
    SendKeys FinancialYear, True
    SendKeys "{TAB}", True
    SendKeys FinancialPeriod, True
    
    detailsButtonX = 550  ' Replace with actual X coordinate
    detailsButtonY = 240 ' Replace with actual Y coordinate
    SetCursorPos detailsButtonX, detailsButtonY
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        
      startTime = Timer
    Do While Timer < startTime + 2 ' Wait for 2 seconds
        DoEvents
    Loop
```
