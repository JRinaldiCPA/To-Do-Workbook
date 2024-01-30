Attribute VB_Name = "myPrivateMacros"
Option Explicit
Sub DisableForEfficiency()

' -----------
' Turns off functionality to speed up Excel
' -----------

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False

End Sub
Sub DisableForEfficiencyOff()

' -----------
' Turns functionality back on
' -----------

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True

End Sub
Private Sub Export_Macros_v2()
    
' Purpose: To get my To Do ready to email myself the current version of the template.
' Trigger: Manual
' Updated: 12/6/2021

' Change Log:
'       9/17/2019: Initial Creation
'       6/2/2021:  Updated to replace Shell with FollowHyperlink
'       12/6/2021: Added ws_Questions to be cleared, and moved the hyperlink opening
'                  Updated the ranges to clear

' ***********************************************************************************************************************************

Call myPrivateMacros.DisableForEfficiency
    
' -----------
' Declare your variables
' -----------
    
    Dim ws As Worksheet
    
' -----------
' Run your code
' -----------
    
        For Each ws In ActiveWorkbook.Worksheets
            ws.Visible = xlSheetVisible
            If ws.AutoFilterMode = True Then ws.AutoFilter.ShowAllData
            ws.Rows.EntireRow.Hidden = False
        Next ws
    
        For Each ws In Worksheets
            
            If ws.Name = "Projects" Then
                ws.Range("A2:Z9999").ClearContents
                ws.Rows("7:9999").Delete
            ElseIf ws.Name = "Tasks" Then
                ws.Range("A2:Z9999").ClearContents
                ws.Rows("7:9999").Delete
            ElseIf ws.Name = "Waiting" Then
                ws.Range("A2:Z9999").ClearContents
                ws.Rows("7:9999").Delete
            ElseIf ws.Name = "Questions" Then
                ws.Range("A2:Z9999").ClearContents
                ws.Rows("7:9999").Delete
            ElseIf ws.Name = "Recurring" Then
                ws.Range("A2:Z9999").ClearContents
                ws.Rows("7:9999").Delete
            ElseIf ws.Name = "Temp" Then
                ws.Range("A2:Z9999").ClearContents
                ws.Rows("7:9999").Delete
            ElseIf ws.Name = "Daily" Then
                ws.Range("A3:Z99999").ClearContents
                ws.Rows("7:9999").Delete
            End If
        Next ws

    ThisWorkbook.FollowHyperlink ("C:\U Drive\Current")
    
    ThisWorkbook.SaveAs Filename:="C:\U Drive\Current\To Do (MACROS) - " & Format(Now, "yyyy-mm-dd"), FileFormat:=xlOpenXMLWorkbookMacroEnabled
        ThisWorkbook.Close

Call myPrivateMacros.DisableForEfficiencyOff
    
End Sub
Private Sub FixHyperlinks()
    
' Purpose: To fix the issue with the links that break when I update my To Do template at home.
' Trigger: Manual
' Updated: 2/26/2021
'
' Change Log:
'       2/26/2021:  Updated to differentiate between ws_Tasks and ws_Projects
'       9/29/2023:  Manually added the code to fix after saving a backup version of my To DO and the links being updated to Roaming

' ***********************************************************************************************************************************
    
' -----------------
' Declare Variables
' -----------------
    
    Dim ws As Worksheet
        Set ws = ActiveWorkbook.ActiveSheet
    
    Dim hl As Hyperlink
    
    Dim strOldLink As String
        strOldLink = "C:\Users\JRina\Downloads\"
    
    Dim strNewLink As String
        If ws.Name = "Projects" Then
            strNewLink = "D:\D Documents\Projects\"
        ElseIf ws.Name = "Tasks" Then
            strNewLink = "D:\D Documents\Tasks\"
        Else
            MsgBox "ERROR - Active sheet isn't ws_Tasks or ws_Projects"
        End If

' ------------------
' Fix the Hyperlinks
' ------------------

    For Each hl In ws.Hyperlinks
        'hl.Address = Replace(hl.Address, strOldLink, strNewLink)
        If ws.Name = "Tasks" Then
            hl.Address = Replace(hl.Address, "C:\Users\Jrinaldi\AppData\Roaming\Microsoft\Excel\Tasks", "Tasks")
        ElseIf ws.Name = "Projects" Then
            hl.Address = Replace(hl.Address, "..\Users\Jrinaldi\AppData\Roaming\Microsoft\Excel\Projects", "Projects")
        End If
    Next hl
    
End Sub
Private Sub Create_Tickler_Folders()
  
' Purpose: To create the tickler folders for the year
' Trigger: Manual
' Updated: 2/17/2022

' Change Log:
'       1/9/2020:  Added the functionality to copy the shortcut for my monthly review and quarterly review
'       2/17/2022: Added code to open the folder when the process is complete

' ***********************************************************************************************************************************
    
 
' -----------------
' Declare Variables
' -----------------
    
    Dim ParentFolder As String
        If Dir("C:\U Drive\Current\Tickler Folders", vbDirectory) = "" Then
            MkDir "C:\U Drive\Current\Tickler Folders"
        End If
        ParentFolder = "C:\U Drive\Current\Tickler Folders"
    
    Dim SubFolder As String
    
    Dim NewPath As String
    
    Dim DaysTillYE As Long
        DaysTillYE = DateValue("12/31/2022") - Date
    
    Dim Day_String As String
    
    Dim i As Long
    
    Dim fso As Object
        Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    
' ------------------
' Create the folders
' ------------------
    
    'Loop through each business day
        For i = 0 To DaysTillYE
            
            'Only create a folder if it is a business day
            Day_String = Format(DateAdd("d", i, Date), "dddd")
                If Day_String <> "Saturday" And Day_String <> "Sunday" Then
            
                    SubFolder = Format(DateAdd("d", i, Date), "mm.dd.yyyy")
                    NewPath = ParentFolder & "\" & SubFolder
            
                    If Dir(NewPath) = "" Then
                        MkDir (NewPath)
                    End If
                    
                    'Copy in my Weekly Review for Thursdays
                    If Day_String = "Thursday" Then
                        Call fso.CopyFile("C:\U Drive\Reference\Templates, Scripts & Batch Files\VB Scripts\(ARCHIVE)\Print Weekly Review.lnk", NewPath & "\")
                    End If

                    'Copy in my Monthly Review for the last business day, and the Quarterly Review
                    
                    If DateAdd("d", i, Date) = WorksheetFunction.WorkDay(WorksheetFunction.EoMonth(DateAdd("d", i, Date), 0) + 1, -1) Then
                        Call fso.CopyFile("C:\U Drive\Support\GTD Reviews\(ARCHIVE)\2. Monthly Review.lnk", NewPath & "\")
                        
                        If Month(DateAdd("d", i, Date)) = 3 Or Month(DateAdd("d", i, Date)) = 6 Or Month(DateAdd("d", i, Date)) = 9 Or Month(DateAdd("d", i, Date)) = 12 Then
                            Call fso.CopyFile("C:\U Drive\Support\GTD Reviews\(ARCHIVE)\3. Quarterly Review.lnk", NewPath & "\")
                        End If
                        
                    End If

                End If
            
        Next i
        
    ' Open the folder
    Call Shell("explorer.exe" & " " & ParentFolder, vbNormalFocus)
       
 End Sub
Sub Save_Me_Email()
Attribute Save_Me_Email.VB_ProcData.VB_Invoke_Func = "H\n14"

' Purpose:  To save myself when I am being held verbally hostage.
' Trigger:  Ctrl + Shift + H
' Updated:  11/23/2021

' Change Log:
'       1/1/2018: Intial Creation
'       11/23/2021: Cleaned up some of the code

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim OutApp As Object
    Dim OutMail As Object
    Dim strBody As String

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

' ----------------------
' Send the Save Me email
' ----------------------

    strBody = "Kim," & vbNewLine & vbNewLine & _
        "This is an automated call for help. Please come save me in 5 minutes." & vbNewLine & vbNewLine & _
        "Thanks," & vbNewLine & _
        "James"

    With OutMail
        .To = "KKacani@WebsterBank.com"
        '.To = "JRinaldi@WebsterBank.com"
        .Importance = 2     '2 flags the email as high importance
        .CC = ""
        .BCC = ""
        .Subject = "Please save me, I am being held verbally hostage"
        .Body = strBody
        .Display
            'Application.Wait (Now + TimeValue("0:00:00.5"))
            Application.SendKeys "%s"
        
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub
Sub Timer_Code()

'My Timer:

Dim sTime As Double

    'Start Timer
    sTime = Timer
    
    '****************** CODE HERE **************************
    
    Debug.Print "Code took: " & (Round(Timer - sTime, 3)) & " seconds"

End Sub

