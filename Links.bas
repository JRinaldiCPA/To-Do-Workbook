Attribute VB_Name = "Links"
Option Explicit
Sub l_1_Open_Pomodoro_Timer()

' Purpose: To open the Pomodoro Timer chrome link.
' Trigger: Ribbon Icon - GTD Macros > Support > Start Pomodoro Timer
' Updated: 7/12/2023
' Reviewd: 12/1/2023

' Change Log:
'       2/17/2022:  Initial Creation
'       7/12/2023:  Updated to use the PomoFocus Chrome App

' ***********************************************************************************************************************************

    'Call Shell("C:\Program Files\Google\Chrome\Application\chrome.exe -url https://pomofocus.io/app", vbMaximizedFocus)

    Call Shell("C:\Program Files\Google\Chrome\Application\chrome_proxy.exe  --profile-directory=Default --app-id=glhjejmflhdjpaimbkdnhfpbbgdgjkoh")

End Sub
Sub l_2_Open_Daily_Reset_PDF()

' Purpose: To open my Daily Reset via PDF.
' Trigger: Called by Workbook_Open event
' Updated: 2/17/2022
' Reviewd: 12/1/2023

' Change Log:
'       2/17/2022:  Initial Creation

' ***********************************************************************************************************************************

    ThisWorkbook.FollowHyperlink ("C:\U Drive\Support\Daily Reset.pdf")

End Sub
Sub l_3_Open_RSA_VPN()

' Purpose: To open the RSA VPN login window.
' Trigger: Called by Workbook_Open event
' Updated: 12/7/2023
' Reviewd:

' Change Log:
'       12/7/2023:  Initial Creation

' ***********************************************************************************************************************************

    ThisWorkbook.FollowHyperlink ("https://connect.websterbank.com")

End Sub
