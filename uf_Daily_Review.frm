VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_Daily_Review 
   Caption         =   "My Daily Review"
   ClientHeight    =   7464
   ClientLeft      =   -12
   ClientTop       =   96
   ClientWidth     =   21144
   OleObjectBlob   =   "uf_Daily_Review.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_Daily_Review"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Declare Enumerations
Enum Columns

    Col_21_Most_Valuable_Work = 8 + 1
    Col_22_Improve_n_Learn = 8 + 2
    Col_23_Gratitude = 8 + 3
    Col_24_Help_People = 8 + 4
    Col_25_Went_Right = 8 + 5
    Col_26_Went_Wrong = 8 + 6
    
    Col_31_Reality = Col_26_Went_Wrong + 1
    Col_32_Focus_Improve = Col_26_Went_Wrong + 2
    Col_33_Exp_Friction = Col_26_Went_Wrong + 3
    Col_34_Live_Today_Again = Col_26_Went_Wrong + 4
    
End Enum

' Declare Worksheets
Dim wbDailyReview As Workbook
Dim wsData As Worksheet

' Declare Integers
Dim int_LastRow As Long
Dim int_CurRow As Long

' Declare Strings
Dim strToday As String
Private Sub UserForm_Initialize()
    
' Purpose: To initialize the userform, including adding in the data from the arrays.
' Trigger: Event: UserForm_Initialize
' Updated: 6/13/2023

' Change Log:
'       12/31/2020: Added the code to abort if I already completed the day's review
'       4/19/2021:  Updated how the int_CurRow is determined
'       4/20/2021:  Added the new enumerations for the Reasssess section
'       5/23/2021:  Refreshed some of the enumerations to match the current order of the fields.
'       6/6/2021:   Added the code to close the Daily Review wb if I abort the process.
'       6/7/2021:   Cleaned up some unused code in o_31_Create_Daily_To_Do_txt
'       6/24/2021:  Added the "Start Review" button and process to start with a smaller UserForm
'       7/22/2021:  Purged the code related to creating a daily Current Work file
'       12/24/2021: Added the adustments for Daily Routines and adjusted the height of frm_Review.
'       7/27/2022:  Added questions #2 and #3 under Reassess
'       10/27/2022: Added the process goal for no email < 10AM
'       11/25/2022: Removed the 'cmd_DailyToDo' button as these now get created as part of my Weekly process
'       5/19/2023:  Updated to remove the 'frm_DailyRoutines', as I now track these daily and in my monthly Current Events file
'                   Pulled in the content from the Start Review button to enlarge the form
'       6/13/2023:  Removed all details around the frm_DailyRoutines and the manipulation of the form height / width

' ***********************************************************************************************************************************
    
'------------------
' Declare Variables
'------------------

    Dim strFileLoc As String
        strFileLoc = "C:\U Drive\Support\Daily Review.xlsx"
        
    ' Dim Workbooks / Worksheets
    
        Set wbDailyReview = Workbooks.Open(strFileLoc)
        Set wsData = wbDailyReview.Worksheets("Daily Review")

    ' Dim Integers
           
        int_LastRow = wsData.Cells(Rows.count, "A").End(xlUp).Row
        int_CurRow = wsData.Range("A:A").Find("").Row

    ' Dim Strings
        
        strToday = Date & " (" & Format(Date, "DDDD") & ")"

'----------------
' Update UserForm
'----------------

    Me.Top = Application.Top + (Application.UsableHeight / 1.35) - (Me.Height / 2) 'Open near the bottom of the screen
    Me.Left = Application.Left + (Application.UsableWidth / 2) - (Me.Width / 2)
   
    ' Add data from To Do - Daily
    Call Me.o_1_Add_Data_From_wsDaily
    
End Sub
Private Sub UserForm_Terminate()

    ' If the workbook is still open (aka you aborted) then close it
        If Not wbDailyReview Is Nothing Then wbDailyReview.Close SaveChanges:=True

End Sub
Private Sub cmd_Done_Click()

    Call Me.o_2_Update_Daily_Review_Workbook
    
Application.EnableEvents = True
    
    Unload Me

End Sub
Private Sub cmd_Cancel_Click()
    
Application.EnableEvents = False
    
    Unload Me
    
End Sub
Sub o_1_Add_Data_From_wsDaily()

' Purpose: To pull in any applicable data from my To Do - Daily ws.
' Trigger: uf_Daily_Review > Initialize
' Updated: 4/20/2021

' Change Log:
'       2/24/2021: Initial Creation
'       4/20/2021: Added a 2nd '+ Chr(10)' to add a gap
'       4/20/2021: Added code to "strikethrough" the old values, to ignore them for the following day

' ***********************************************************************************************************************************

Call myPrivateMacros.DisableForEfficiency

' -----------------
' Declare Variables
' -----------------

    ' Dim Worksheets
    
    Dim ws_Daily As Worksheet
        Set ws_Daily = ThisWorkbook.Sheets("Daily")

    ' Dim Integers
    
    Dim intDataStart As Long
        intDataStart = ws_Daily.Cells(Rows.count, "A").End(xlUp).Row + 3
    
    Dim intDataEnd As Long
        intDataEnd = ws_Daily.Cells(Rows.count, "B").End(xlUp).Row + 8
        
    ' Dim Strings
        
    Dim strCurrCategory As String
    
    Dim strImprove_n_Learn As String
    Dim strWentRight As String
    Dim strWentWrong As String
    Dim strPositiveExp As String
        
    ' Dim Ranges
        
    Dim rngSearchData As Range
    Set rngSearchData = ws_Daily.Range(ws_Daily.Cells(intDataStart, "B"), ws_Daily.Cells(intDataEnd, "B"))
    
    Dim cell As Variant

' ------------------------------------------------
' Create the strings for the Daily Review UserForm
' ------------------------------------------------
    
    For Each cell In rngSearchData
        
        ' Determine what category you are currently in
        If cell.Value2 = "Improved / Learned" Then
            strCurrCategory = "Improved / Learned"
        ElseIf cell.Value2 = "Start / Continue" Then
            strCurrCategory = "Start / Continue"
        ElseIf cell.Value2 = "Stop / Change" Then
            strCurrCategory = "Stop / Change"
        ElseIf cell.Value2 = "Positive Experiences" Then
            strCurrCategory = "Positive Experiences"
        End If
    
        ' If it's actual data punt it to the UserForm
        If cell.Value2 <> strCurrCategory And cell.Value2 <> "" And cell.Font.Strikethrough = False Then
            If strCurrCategory = "Improved / Learned" Then
                strImprove_n_Learn = cell.Value2 + Chr(10) + Chr(10) & strImprove_n_Learn
            ElseIf strCurrCategory = "Start / Continue" Then
                strWentRight = cell.Value2 + Chr(10) + Chr(10) & strWentRight
            ElseIf strCurrCategory = "Stop / Change" Then
                strWentWrong = cell.Value2 + Chr(10) + Chr(10) & strWentWrong
            ElseIf strCurrCategory = "Positive Experiences" Then
                strPositiveExp = cell.Value2 + Chr(10) + Chr(10) & strPositiveExp
            End If
            
            cell.Font.Strikethrough = True
        End If
    
    Next cell
    
' ----------------------------------------------
' Output the values to the Daily Review UserForm
' ----------------------------------------------
    
    Me.txt_12_Improve_n_Learn = strImprove_n_Learn
    Me.txt_14_WentRight = Trim(strWentRight)
    Me.txt_15_WentWrong = strWentWrong
        
Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_2_Update_Daily_Review_Workbook()

' Purpose: To update the Daily Review workbook.
' Trigger: uf_Daily_Review > cmd_Done_Click
' Updated: 6/13/2023

' Change Log:
'       6/10/2020:  Initial creation
'       8/31/2020:  Complete overhaul, converted to Excel
'       4/20/2021:  Updated to include the Reassess section
'       7/27/2022:  Added additional text boxes under Reassess and refreshed the #s
'       10/27/2022: Added the process goal for no email < 10AM
'       2/27/2023:  Updated so that only the Date is output in the first field
'       6/13/2023:  Removed the code related to the Daily Process Goals

' ***********************************************************************************************************************************

'-------------------------------------
' Output the results from the UserForm
'-------------------------------------

    ' Add the current date
    wsData.Cells(int_CurRow, 1).Value2 = Date

    With wsData
    
        ' Reflect
        .Cells(int_CurRow, Columns.Col_21_Most_Valuable_Work) = Me.txt_11_MostValuableWork
        .Cells(int_CurRow, Columns.Col_22_Improve_n_Learn) = Me.txt_12_Improve_n_Learn
        .Cells(int_CurRow, Columns.Col_23_Gratitude) = Me.txt_13_Gratitude
        .Cells(int_CurRow, Columns.Col_25_Went_Right) = Me.txt_14_WentRight
        .Cells(int_CurRow, Columns.Col_26_Went_Wrong) = Me.txt_15_WentWrong
    
        ' Reassess
        .Cells(int_CurRow, Columns.Col_31_Reality) = Me.txt_21_Reality
        .Cells(int_CurRow, Columns.Col_32_Focus_Improve) = Me.txt_22_Focus_Improve
        .Cells(int_CurRow, Columns.Col_33_Exp_Friction) = Me.txt_23_Exp_Friction
        .Cells(int_CurRow, Columns.Col_34_Live_Today_Again) = Me.txt_24_Live_Today_Again
    
    End With

End Sub
