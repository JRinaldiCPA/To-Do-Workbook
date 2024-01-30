VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_New_NextAction 
   Caption         =   "Next Action Input Form"
   ClientHeight    =   8688.001
   ClientLeft      =   2952
   ClientTop       =   684
   ClientWidth     =   8808.001
   OleObjectBlob   =   "uf_New_NextAction.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_New_NextAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare Sheets
    Dim ws_Projects As Worksheet
    Dim ws_Tasks As Worksheet
    Dim ws_Waiting As Worksheet
    Dim ws_Questions As Worksheet
    
'Declare Strings
    Dim strProjName As String
    Dim strProjArea As String
    Dim strNextActionType As String
    Dim strTaskSource As String

'Declare Integers
    Dim int_CurRow As Long
    Dim int_LastCol As Long
    Dim int_LastCol_wsProjects As Long
    
'Declare ws_Projects Cell References
    Dim arry_Header_wsProjects() As Variant
    
    Dim col_Area_wsProjects As Long
    Dim col_Project_wsProjects As Long
    Dim col_Status_wsProjects As Long

'Declare Common Cell References
    Dim arry_Header() As Variant
    
    Dim col_Ref As Long
    Dim col_Entered As Long
    Dim col_Priority As Long
    Dim col_Project As Long
    Dim col_NextAction As Long
    Dim col_Completed As Long
    Dim col_Notes As Long
    
'Declare wsTask Cell References
    Dim col_Start As Long
    Dim col_Context As Long
    Dim col_Time As Long
    Dim col_Area As Long
    
    Dim col_Component As Long
    Dim col_ActiveProject As Long
    Dim col_ActiveTask As Long
    
'Declare ws_Waiting Cell References
    Dim col_WaitingFor As Long
            
'Declare ws_Questions Cell References
    Dim col_WhoCanAnswer As Long
            
'Declare Dictionaries
    Dim dict_ProjectNames As Scripting.Dictionary

'Dim Booleans
    Dim bol_EnableEvents As Boolean

Private Sub frm_TaskDesc_Click()

End Sub

Private Sub UserForm_Initialize()

' Purpose: To initialize the userform, including adding in the data from the arrays and pulling data from the current row if a Project is selected.
' Trigger: Event: UserForm_Initialize
' Updated: 10/11/2023
' Reviewd: 10/11/2023

' Change Log:
'       2/27/2020:  Updated all of the worksheet references to be Global
'       2/27/2020:  Updated to pull the data from clipboard by default, otherwise whatever is in the cell
'       2/28/2020:  Updated to include an error handler when the content cant be copied from the clipboard
'       12/16/2020: Removed old code related to the clipboard that wasn't applicable
'       3/8/2021:   Setup to pull in the cell contents if creating a task from Temp
'       6/2/2021:   Resolved an error if a blank row is selected when pulling in the project details
'       7/1/2021:   Added 'Me.lst_Area.AddItem .Range("C" & selection.Row).Value' to hopefully resolve the N/A Areas issue
'       7/19/2021:  Added code to remove the bullet if I am copying in from my Daily > Next Actions
'       8/8/2021:   Added the code for Task Notes
'       8/12/2021:  Added the option to fill in the Completed date
'       8/15/2021:  Added code to pre-fill the txt_Waiting_For value
'       8/18/2021:  Switched to Len for checking if there is a value in the Task
'       9/4/2021:   Added code to default the chk_Continuous = True when selecting Area = Continuous
'       9/4/2021:   Made a number of small tweaks to the formatting and size of TxtBoxes
'       9/5/2021:   Added the Zoom Adjust related code
'       9/10/2021:  Added the code to include sub lines to the Notes when copying in a Task from ws_Temp
'       10/1/2021:  Added code to automatically mark a Task / Waiting as complete if Strikethrough = True
'       10/2/2021:  Added code to pull the applicable project if I have one selected in Tasks
'       10/2/2021:  Removed the code related to strTempProjName, now that I don't have to add then select values in my ListBoxes
'       10/3/2021:  Added code to look for the applicable project, and auto select it
'       10/4/2021:  Added code to indicate if pulling value from Notes
'       10/4/2021:  Moved some code to make it more clear what should be pulled in when I'm importing from Temp
'       10/4/2021:  Added clear sections for pulling in data from Projects / Tasks / Temp
'       10/25/2021: Added the code to reset the Area ListBox when I Double Click
'       10/31/2021: Updated so the 'To Do Only' button is the default selection
'       11/1/2021:  Removed the code related to the 'opt_PXXX_Only' and 'opt_Both' and 'opt_To_Do_Only'
'       11/1/2021:  Removed the code related to 'intProjRow' and 'intProjNum'
'       11/16/2021: Updated the reset if a Continuous project to add the applicable Project, not just update the Area
'       11/23/2021: Added the code to send Next Actions to ws_Questions, and replaced the opt Ifs w/ the strNextActionType to be consistent
'       12/8/2021:  Added the bolProjectSelected bolean
'                   Moved some of the code to o_03_Import_Initial_Data
'       12/11/2021: Removed the variety of int_CurRow
'       12/24/2021: Added the 'lst_NextAction' related code
'       12/26/2021: Eliminated the 'strNextActionType' code, no longer used
'       1/26/2022:  Added the code to remove the (?) from questions
'       2/11/2022:  Added the code to remove the ? from the beginning of a question
'       4/28/2022:  Removed the @ACTION and @WAITING options from the Context list
'       5/23/2022:  Updated to my new symbols for my Next Actions
'       11/1/2022:  Removed the "*" bullet when pulling in a Task from my Daily To Do
'       11/9/2022:  Removed the [¤] bullet when pulling in a Task from my Daily To Do
'       3/20/2023:  Updated so if I am in Tasks to copy the active cell task
'       5/9/2023:   Pulled in the dynamic project selector from the uf_Project_Selector UserForm
'       6/28/2023:  Updated to adjust the height of the Areas list for Personal vs Professional
'       9/5/2023:   Created the 'intActiveRow' code and pulled in more fields if Tasks was the Active sheets
'       9/24/2023:  Updated to pull the ActiveCell value for the desc. if not Tasks
'                   Updated so that if I already have a Waiting For to not pull from Next Action desc.
'       10/11/2023: Added the code and related references for o_022_Declare_wsProjects_Variables
'                   Moved alot of the content to 'o_03_Import_Initial_Data'

' ***********************************************************************************************************************************

Call Me.o_021_Declare_Global_Variables

' -----------------
' Declare Variables
' -----------------

    'Declare Other
                   
    Dim i As Long
    
    Dim dblZoomAdjust As Double
        dblZoomAdjust = 1.05
        
' ---------------------
' Initialize the values
' ---------------------

    'Set EnableEvents to True
    bol_EnableEvents = True

    'Assign the Value for strTaskSource
    If ActiveSheet.Name = "Temp" And ActiveCell.Value <> "" Then
        strTaskSource = "Temp"
    
    ElseIf ActiveSheet.Name = "Projects" Then
        strTaskSource = "Projects"
        Call Me.o_022_Declare_wsProjects_Variables
    
    ElseIf ActiveSheet.Name = "Tasks" Then
        strTaskSource = "Tasks"
        Call Me.o_023_Declare_wsTasks_Variables
    
    ElseIf ActiveSheet.Name = "Waiting" Then
        strTaskSource = "Waiting"
        Call Me.o_024_Declare_wsWaiting_Variables

    End If

    'Add Value for Entered TextBox
        txt_Entered.Value = Date
        
    'Add items to Time ListBox
        lst_Time.AddItem "<30 mins"
        lst_Time.AddItem "30m - 2hr"
        lst_Time.AddItem ">2 hours"

    'Add items to Priority ListBox
        lst_Priority.AddItem "High"
        lst_Priority.AddItem "Medium"
        lst_Priority.AddItem "Low"
        
    'Add the values for the Area ListBox
        lst_Area.List = GetAreaArray
    
    'Add the values for the Context ListBox
        lst_Context.List = GetContextArray
        
    'Add the values for the Next Actions ListBox
        lst_NextAction.AddItem "Task"
        lst_NextAction.AddItem "Waiting"
        lst_NextAction.AddItem "Question"

' --------------------------------------
' Import additional data to the UserForm
' --------------------------------------

    Call Me.o_03_Import_Initial_Data

' ----------------------------------------
' Adjust the size of lst_Area for Personal
' ----------------------------------------

    #If Personal = 1 Then
        Me.lst_Area.Height = Me.lst_Area.Height + 16
        Me.chk_Continuous.Top = Me.chk_Continuous.Top + 16
        Me.chk_Pending.Top = Me.chk_Pending.Top + 16
    #End If

' -------------------------------
' Adjust the size of the UserForm
' -------------------------------
        
    Me.Zoom = Me.Zoom * dblZoomAdjust
    Me.Height = Me.Height * dblZoomAdjust
    Me.Width = Me.Width * dblZoomAdjust
    
    Me.cmb_DynamicSearch.Top = 8 '10/12/23 -> Trying to force this, it keeps reseting for some reason

End Sub
Private Sub cmb_DynamicSearch_Change()

' Purpose: To replace the values in the Projects ListBox based on what is typed.
' Trigger: Called by uf_Project_Selector
' Updated: 10/11/2023
' Reviewd: 10/11/2023

' Change Log:
'       4/12/2020:  Removed any reference to the D/A Worksheet, removed with my move to CR&A
'       12/10/2021: Removed the code related to the P.XXX name, now that I include that in the Project Name
'       10/11/2023: Overhauled code to match my Dynamic Selectors for Files / Folders / RefX
'                   Updated to clear the Area list if it isn't already

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim val As Variant
    
' ---------------------------------------------------
' Copy the values from the dictionary to the list box
' ---------------------------------------------------
       
    Me.lst_Project.Clear
        
    For Each val In dict_ProjectNames
        If InStr(1, val, cmb_DynamicSearch.Value, vbTextCompare) Then
            Me.lst_Project.AddItem val 'If the name is similar then add to the list
        End If
    Next val
       
    If Me.lst_Area.Value <> "" Then Me.lst_Area.Clear
    
' ------------------------------
' If only one remains, select it
' ------------------------------

    If lst_Project.ListCount = 1 Then
        lst_Project.Selected(0) = True
    End If

End Sub
Private Sub lst_Area_Click()

If bol_EnableEvents = False Then Exit Sub

    Call Me.o_04_Create_Project_List
    
End Sub
Private Sub lst_Area_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
' Purpose: To reset the Area ListBox when I double click on it.
' Updated: 10/12/2023
' Reviewd: 10/12/2023

' Change Log:
'       10/12/2023: Initial Documentation
'                   Added the code to clear the lst_Project

' ***********************************************************************************************************************************
    
    ' Reset the Area ListBox when I Double Click
    
    Me.lst_Area.Clear
    Me.lst_Area.List = GetAreaArray
    
    Me.lst_Project.Clear

End Sub
Private Sub lst_Project_Click()

' Purpose: To select a project and fill in the Area if it is blank.
' Updated: 10/12/2023
' Reviewd: 10/12/2023

' Change Log:
'       10/11/2023: Initial Documentation
'                   Added the bol_NoSelectedArea to handle when I pick a project from the Dynamic Search
'       10/12/2023: Added code for bol_EnableEvents
'                   Removed the bol_NoSelectedArea, it was having issues when I already picked an Area but then picked a different Project

' ***********************************************************************************************************************************

If bol_EnableEvents = False Then Exit Sub

' -----------------
' Declare Variables
' -----------------

    'Declare Booleans
    Dim bol_NoSelectedArea As Boolean
        
    If Me.lst_Area.ListIndex = -1 Then
        bol_NoSelectedArea = True
    End If
    
    'Declare Integers
    Dim int_ProjectRow As Long
        
    'Declare Strings
    Dim str_SelectedProjectArea As String
    
    'Assign Variables
        int_ProjectRow = fx_Find_CurRow(ws:=ws_Projects, strTargetFieldName:="Project", strTarget:=lst_Project.Value)
        str_SelectedProjectArea = ws_Projects.Cells(int_ProjectRow, col_Area_wsProjects)

' ------------------------------------------------------------------------------------------
' If I used Dynamic Search repopulate the Area List and select based on the selected project
' ------------------------------------------------------------------------------------------

bol_EnableEvents = False

    Me.lst_Area.List = GetAreaArray
    Me.lst_Area.Value = str_SelectedProjectArea
    
bol_EnableEvents = True

End Sub
Private Sub lst_Project_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    
' Purpose: To copy the selected Project if I'm updating my Notes, or to clear the list.
' Updated: 6/27/2023

' Change Log:
'       4/16/2022:  Initial Creation
'       2/3/2023:   Moved the code so it only applies the formatting if I am in the "Project:" cell
'                   Added the code so that if I am NOT in the "Project:" cell it opens the folder
'       6/27/2023:  Overhauled this code as I had pulled it in from my Project Selector in May, and didn't fully test

' ***********************************************************************************************************************************

    ' Determine if I am upating my meeting notes, if not then copy to the clipboard and clear the project lists
    
    If ActiveCell.Value = "Project: " Then
        ActiveCell.Value = ActiveCell.Value & Me.lst_Project.Value
        
        ActiveCell.Font.Bold = False
        ActiveCell.Characters(1, Len("Project:")).Font.Bold = True
        
        Unload Me
    Else
        fx_Copy_to_Clipboard (lst_Project.Value)
        Me.lst_Project.Clear
    End If
    
End Sub
Private Sub lst_Project_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

' Purpose: To copy the selected Project into the Clipboard.
' Updated: 6/27/2023

' Change Log:
'       4/16/2022:  Initial Creation
'       6/27/2023:  Overhauled this code as I had pulled it in from my Project Selector in May, and didn't fully test
'       7/23/2023:  Updated to move the Unload Me to the If statement, so if I hit a button accidently it doesn't unload

' ***********************************************************************************************************************************

    If KeyCode = vbKeyReturn And Me.lst_Project.Value <> "" Then
        fx_Copy_to_Clipboard (lst_Project.Value)
        Unload Me
    End If

End Sub
Private Sub lst_Context_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

' Purpose: To clear the list of Contexts when I double click
' Updated: 4/28/2022

' Change Log:
'       4/28/2022:  Initial Creation

' ***********************************************************************************************************************************

    ' Reset the List
    Me.lst_Context.Clear
    Me.lst_Context.List = GetContextArray

End Sub
Private Sub lst_NextAction_Click()

' Purpose: To udpate the UserForm to just the fields required for the selected Next Action.
' Updated: 12/24/2021

' Change Log:
'       12/24/2021: Initial Creation, based on Next Action Option buttons
'       12/27/2021: Converted to using strNextActionType for when I select the NextAction by code and it doesn't update the ListBox value

' ***********************************************************************************************************************************
        
    'Set the strNextActionType in case it is being pulled in from a keyboard shortcut
    strNextActionType = Me.lst_NextAction.List(Me.lst_NextAction.ListIndex)

    If strNextActionType = "Task" Then
        Call Me.o_31_Next_Action_is_Task
    ElseIf strNextActionType = "Waiting" Then
        Call Me.o_32_Next_Action_is_Waiting
    ElseIf strNextActionType = "Question" Then
        Call Me.o_33_Next_Action_is_Question
    End If

End Sub
Private Sub txt_NextAction_Desc_Change()

    Me.cmd_AddNextAction.Default = True

End Sub
Private Sub chk_Completed_Click()

' Purpose: To popup the txt_Completed object and auto-fill w/ today's date.
' Updated: 8/19/2023

' Change Log:
'       11/23/2021: Added the code to turn the txt_Completed off when
'       12/24/2021: Added to convert the label to "Completed"
'       8/19/2023:  Updated so if I check Completed it uses the Entered date as the Completion Date instead of Today

' ***********************************************************************************************************************************

If Me.chk_Completed.Value = True Then

    Me.txt_Completed.Visible = True
    Me.txt_Completed.Value = Me.txt_Entered.Value
    
    Me.txt_Completed.SetFocus
    
ElseIf Me.chk_Completed.Value = False Then

    Me.txt_Completed.Visible = False
    Me.txt_Completed.Value = ""

End If

End Sub
Private Sub cmd_AddNextAction_Click()

' Purpose: To determine if a task should be added to ws_Tasks, ws_Waiting, or ws_Questions.
' Updated: 11/23/2021

' Change Log:
'       3/20/2020: Updated to only add to Both if a project was actually selected
'       12/16/2020: Added the reset code when adding a new Task or Waiting to my To Do
'       11/1/2021: Removed the code related to the 'opt_PXXX_Only' and 'opt_Both' and 'opt_To_Do_Only'
'       11/23/2021: Added the code for Questions
'       12/26/2021: Switched from strNextActionType to 'Me.lst_NextAction'
'       12/27/2021: Switched back to strNextActionType to account for when I select the NextAction by code and it doesn't update the ListBox value
'       10/12/2023: Removed the code to re-run o_021_Declare_Global_Variables

' ***********************************************************************************************************************************

'Call me.o_021_Declare_Global_Variables

'Check if a task was added
If txt_NextAction_Desc.Value = "" Then
    MsgBox "The Task description text box was blank, a new Task was not created"
    Unload Me
End If

' -----------------------------------------------------------
' Add the values to the ws_Tasks or ws_Waiting your variables
' -----------------------------------------------------------

    If strNextActionType = "Task" Then
        Call Me.o_11_Add_New_Task
    ElseIf strNextActionType = "Waiting" Then
        Call Me.o_12_Add_New_Waiting
    ElseIf strNextActionType = "Question" Then
        Call Me.o_13_Add_New_Question
    Else
        MsgBox "Something went awry, no Next action type was selected in the ListBox"
        Stop
    End If

' ---------------------
' Add the @TASKS Folder
' ---------------------
    
    If Me.lst_Context.Value = CStr("@ TASKS") Then
        Call Me.o_2_Create_Task_Folder
    End If
 
    Unload Me

End Sub
Private Sub chk_Pending_Click()

Call Me.o_04_Create_Project_List

End Sub
Private Sub chk_Continuous_Click()

Call Me.o_04_Create_Project_List

End Sub
Private Sub cmd_Cancel_Click()

    Unload Me

End Sub
Sub o_021_Declare_Global_Variables()

' Purpose: To assign the Private Variables that were Declared above the line.
' Trigger: Event: UserForm_Initialize
' Updated: 12/26/2021

' Change Log:
'       4/25/2020:  Made the lookup for project number more dynamic.
'       5/14/2021:  Moved the calculation for the Task / Waiting # to be feed by the applicable workbook.
'       10/22/2021: Added code to handle the "N/A" project, aka when one doesn't exist
'       11/1/2021:  Removed the code related to 'intProjRow' and 'intProjNum'
'       11/2/2021:  Moved the code to Declare/Assign variables for ws_Tasks and ws_Waiting into o_021_Declare_Global_Variables
'       11/23/2021: Updated to use the 'strNextActionType' instead of the Option button
'       11/23/2021: Updated 'col_Task' to 'col_NextAction'
'       12/8/2021:  Added the code for arrySelectedProject and updated a slew of related code
'                   Broke out the assigning of each worksheets variables
'       12/11/2021: Moved the 'Assign Worksheets' from the Initialize
'       12/14/2021: Added the code to abort if the project name is "N/A"
'       12/26/2021: Updated the intRowCurProject to use 'fx_Find_CurRow'
'       10/12/2023: Removed all the code related to the array for a selected project

' ***********************************************************************************************************************************

' ----------------
' Assign Variables
' ----------------
        
    ' Assign Worksheets
    
    Set ws_Projects = ThisWorkbook.Sheets("Projects")
    Set ws_Tasks = ThisWorkbook.Sheets("Tasks")
    Set ws_Waiting = ThisWorkbook.Sheets("Waiting")
    Set ws_Questions = ThisWorkbook.Sheets("Questions")

    ' Assign ws_Projects Cell References
    
    int_LastCol_wsProjects = ws_Projects.Cells(1, ws_Projects.Columns.count).End(xlToLeft).Column
    
    arry_Header_wsProjects = Application.Transpose(ws_Projects.Range(ws_Projects.Cells(1, 1), ws_Projects.Cells(1, int_LastCol_wsProjects)))

    col_Area_wsProjects = fx_Create_Headers("Area", arry_Header_wsProjects)
    col_Project_wsProjects = fx_Create_Headers("Project", arry_Header_wsProjects)
    col_Status_wsProjects = fx_Create_Headers("Status", arry_Header_wsProjects)

End Sub
Sub o_022_Declare_wsProjects_Variables()

' Purpose: To assign the Private Variables that were Declared above the line.
' Trigger: Event: UserForm_Initialize
' Updated: 10/11/2023

' Change Log:
'       10/11/2023:  Initial Creation, based on o_023_Declare_wsTasks_Variables

' ***********************************************************************************************************************************

' -----------------------------
' Declare ws_Projects Variables
' -----------------------------

    ' Declare Integers
    
    int_LastCol = ws_Projects.Cells(1, ws_Projects.Columns.count).End(xlToLeft).Column

    ' Assign Cell References
    
    arry_Header = Application.Transpose(ws_Projects.Range(ws_Projects.Cells(1, 1), ws_Projects.Cells(1, int_LastCol)))

    col_Area = fx_Create_Headers("Area", arry_Header)
    col_Project = fx_Create_Headers("Project", arry_Header)
    
End Sub
Sub o_023_Declare_wsTasks_Variables()

' Purpose: To assign the Private Variables that were Declared above the line.
' Trigger: Event: UserForm_Initialize
' Updated: 12/8/2021

' Change Log:
'       12/8/2021:  Initial Creation, broke out the assigning of each worksheets variables

' ***********************************************************************************************************************************

' -------------------------
' Declare ws_Tasks Variables
' -------------------------

    ' Declare Integers
    
    int_CurRow = [MATCH(TRUE,INDEX(ISBLANK('[To Do.xlsm]Tasks'!A:A),0),0)]
    int_LastCol = ws_Tasks.Cells(1, ws_Tasks.Columns.count).End(xlToLeft).Column

    ' Assign Cell References
    
    arry_Header = Application.Transpose(ws_Tasks.Range(ws_Tasks.Cells(1, 1), ws_Tasks.Cells(1, int_LastCol)))

    col_Ref = fx_Create_Headers("#", arry_Header)
    col_Entered = fx_Create_Headers("Entered", arry_Header)
    col_Start = fx_Create_Headers("Start", arry_Header)
    
    col_Context = fx_Create_Headers("Context", arry_Header)
    col_Priority = fx_Create_Headers("Priority", arry_Header)
    col_Time = fx_Create_Headers("Time", arry_Header)
    col_Area = fx_Create_Headers("Area", arry_Header)
    col_Project = fx_Create_Headers("Project", arry_Header)
    col_Component = fx_Create_Headers("Component", arry_Header)
    col_NextAction = fx_Create_Headers("Task", arry_Header)
    
    col_ActiveProject = fx_Create_Headers("Active Proj.", arry_Header)
    col_ActiveTask = fx_Create_Headers("Active Task", arry_Header)
    
    col_Completed = fx_Create_Headers("Completed", arry_Header)
    col_Notes = fx_Create_Headers("Notes", arry_Header)

End Sub
Sub o_024_Declare_wsWaiting_Variables()

' Purpose: To assign the Private Variables that were Declared above the line.
' Trigger: Event: UserForm_Initialize
' Updated: 12/8/2021

' Change Log:
'       12/8/2021:  Initial Creation, broke out the assigning of each worksheets variables

' ***********************************************************************************************************************************

' ---------------------------
' Declare ws_Waiting Variables
' ---------------------------

    ' Declare Integers

    int_CurRow = [MATCH(TRUE,INDEX(ISBLANK('[To Do.xlsm]Waiting'!A:A),0),0)]
    int_LastCol = ws_Waiting.Cells(1, ws_Waiting.Columns.count).End(xlToLeft).Column
    
    ' Assign Cell References
    
    arry_Header = Application.Transpose(ws_Waiting.Range(ws_Waiting.Cells(1, 1), ws_Waiting.Cells(1, int_LastCol)))

    col_Ref = fx_Create_Headers("#", arry_Header)
    col_Entered = fx_Create_Headers("Entered", arry_Header)
    col_Priority = fx_Create_Headers("Priority", arry_Header)
    col_Project = fx_Create_Headers("Project", arry_Header)
    col_NextAction = fx_Create_Headers("Task", arry_Header)
    col_WaitingFor = fx_Create_Headers("Waiting For", arry_Header)
    col_Completed = fx_Create_Headers("Completed", arry_Header)
    col_Notes = fx_Create_Headers("Notes", arry_Header)

End Sub
Sub o_025_Declare_wsQuestions_Variables()

' Purpose: To assign the Private Variables that were Declared above the line.
' Trigger: Event: UserForm_Initialize
' Updated: 12/8/2021

' Change Log:
'       12/8/2021:  Initial Creation, broke out the assigning of each worksheets variables

' ***********************************************************************************************************************************

' ---------------------------
' Declare ws_Questions Variables
' ---------------------------

    ' Declare Integers

    int_CurRow = [MATCH(TRUE,INDEX(ISBLANK('[To Do.xlsm]Questions'!A:A),0),0)]
    int_LastCol = ws_Questions.Cells(1, ws_Questions.Columns.count).End(xlToLeft).Column
    
    ' Assign Cell References
    
    arry_Header = Application.Transpose(ws_Questions.Range(ws_Questions.Cells(1, 1), ws_Questions.Cells(1, int_LastCol)))

    col_Ref = fx_Create_Headers("#", arry_Header)
    col_Entered = fx_Create_Headers("Entered", arry_Header)
    col_Priority = fx_Create_Headers("Priority", arry_Header)
    col_Project = fx_Create_Headers("Project", arry_Header)
    col_NextAction = fx_Create_Headers("Pending Question", arry_Header)
    col_WhoCanAnswer = fx_Create_Headers("Who Can Answer", arry_Header)
    col_Completed = fx_Create_Headers("Completed", arry_Header)
    col_Notes = fx_Create_Headers("Notes", arry_Header)

End Sub
Sub o_03_Import_Initial_Data()

' Purpose: To import the initial data from my To Do.
' Trigger: Event: UserForm_Initialize
' Updated: 10/23/2023
' Reviewd: 10/11/2023

' Change Log:
'       12/8/2021:  Initial Creation
'       12/10/2021: Removed the code related to the P.XXX name, and bolDARequest, now that I include that in the Project Name
'       5/26/2022:  Updated to replace "  >" with ">" for new notes
'       10/11/2023: Pulled in content from UserForm_Initialize and overhauled to simplify
'                   Added the content related to the ProjectNames Dictionary
'       10/12/2023: Updated to breakout the variables from setting the values
'                   Reduced and simplified / normalized the code
'       10/23/2023: Added the code to determine the Entered date automatically for new Next Actions

' ***********************************************************************************************************************************

On Error Resume Next

' -----------------
' Declare Variables
' -----------------

    ' Declare Integers
    
    Dim intActiveRow As Long
        intActiveRow = ActiveCell.Row
        
    Dim intLastRow_wsProjects As Long
        intLastRow_wsProjects = fx_Find_LastRow(ws_Target:=ws_Projects, bolIncludeSpecialCells:=True, bolIncludeUsedRange:=True)
        
    Dim i As Long

    ' Declare Project Related Variables
        
    Dim str_Applicable_Area As String
    
    Dim str_Applicable_Project As String
    
    Dim str_Applicable_Component As String
    
    Dim str_Applicable_NextAction As String
    
    Dim str_Applicable_WaitingFor As String
                   
    ' Declare Ranges
    
    Dim rngNotes As Range
    Set rngNotes = Selection.CurrentRegion
    
    ' Assign Dictionaries
    
    Set dict_ProjectNames = New Scripting.Dictionary
        
' ----------------------------
' Pull in values from my To Do
' ----------------------------

    'Assign the Variables
    If strTaskSource = "Projects" Then
        str_Applicable_Project = ws_Projects.Cells(intActiveRow, col_Project)
        str_Applicable_NextAction = fx_Copy_from_Clipboard
    
    ElseIf strTaskSource = "Tasks" Then
        str_Applicable_Project = ws_Tasks.Cells(intActiveRow, col_Project)
        str_Applicable_Component = ws_Tasks.Cells(intActiveRow, col_Component)
        str_Applicable_NextAction = ws_Tasks.Cells(intActiveRow, col_NextAction)
    
    ElseIf strTaskSource = "Waiting" Then
        str_Applicable_Project = ws_Waiting.Cells(intActiveRow, col_Project)
        str_Applicable_NextAction = ws_Waiting.Cells(intActiveRow, col_NextAction)
        str_Applicable_WaitingFor = ws_Waiting.Cells(intActiveRow, col_WaitingFor)
    
    ElseIf strTaskSource = "Temp" Then
        str_Applicable_Project = Mid(CStr(rngNotes.Find("Project:").Value2), (7 + 3), 99)
        str_Applicable_NextAction = ActiveCell.Value
        Me.txt_Entered = Mid(CStr(rngNotes.Find("Meeting Date:").Value2), (12 + 3), 10) ' 10/23/23: Added temporarily
    Else
        str_Applicable_NextAction = fx_Copy_from_Clipboard
    End If
                
    'Assign Area
    str_Applicable_Area = ws_Projects.Cells(fx_Find_CurRow(ws:=ws_Projects, strTargetFieldName:="Project", strTarget:=str_Applicable_Project), col_Area_wsProjects)
                
' ---------------------------------------------------
' Pull in values from my To Do > Temp that are unique
' ---------------------------------------------------
    
    If strTaskSource = "Temp" Then
        
        'Add the detail to the Task Notes if I selected multiple lines
        If Selection.Rows.count > 1 And Selection.Rows.count < 6 Then
            For i = 2 To Selection.Rows.count
                Me.txt_NextAction_Notes = Me.txt_NextAction_Notes & Replace(Selection.Rows(i).Value2, "  >", ">")
                If i < Selection.Rows.count Then Me.txt_NextAction_Notes = Me.txt_NextAction_Notes & Chr(10)
            Next i
        End If
    
        'Mark as Completed if the ActiveCell has the Strikethrough formatting
        If ActiveCell.Font.Strikethrough = True Then
            Me.chk_Completed = True
            Me.lst_Priority = "High"
        End If
        
    End If
                
' ----------------------------------------------
' Assign the Values based on the above Variables
' ----------------------------------------------
                
bol_EnableEvents = False
                        
        ' Area
        If str_Applicable_Area <> "" Then
            Me.lst_Area.Clear
            Me.lst_Area.AddItem str_Applicable_Area
            Me.lst_Area.Value = str_Applicable_Area
        End If
        
        ' Project
        If str_Applicable_Project <> "" Then
            Me.lst_Project.Clear
            Me.lst_Project.AddItem str_Applicable_Project
            Me.lst_Project.Value = str_Applicable_Project
        End If

        ' Other
        Me.cmb_Component.Value = str_Applicable_Component
        Me.txt_NextAction_Desc.Value = str_Applicable_NextAction
        Me.txt_Waiting_For.Value = str_Applicable_WaitingFor
        
' ----------------------
' Modify the Next Action
' ----------------------

    'Remove the bullet
    If Left(Me.txt_NextAction_Desc.Value, 3) = "[•]" Then
        Me.txt_NextAction_Desc.Value = Right(Me.txt_NextAction_Desc.Value, Len(Me.txt_NextAction_Desc.Value) - 4)
    ElseIf Left(Me.txt_NextAction_Desc.Value, 3) = "[o]" Then
        Me.txt_NextAction_Desc.Value = Right(Me.txt_NextAction_Desc.Value, Len(Me.txt_NextAction_Desc.Value) - 4)
    ElseIf Left(Me.txt_NextAction_Desc.Value, 3) = "[¿]" Then
        Me.txt_NextAction_Desc.Value = Right(Me.txt_NextAction_Desc.Value, Len(Me.txt_NextAction_Desc.Value) - 4)
    ElseIf Left(Me.txt_NextAction_Desc.Value, 3) = "[¤]" Then
        Me.txt_NextAction_Desc.Value = Right(Me.txt_NextAction_Desc.Value, Len(Me.txt_NextAction_Desc.Value) - 4)
    
    ElseIf Left(Me.txt_NextAction_Desc.Value, 1) = "*" Then
        Me.txt_NextAction_Desc.Value = Right(Me.txt_NextAction_Desc.Value, Len(Me.txt_NextAction_Desc.Value) - 2)
    ElseIf Left(Me.txt_NextAction_Desc.Value, 1) = "•" Then
        Me.txt_NextAction_Desc.Value = Right(Me.txt_NextAction_Desc.Value, Len(Me.txt_NextAction_Desc.Value) - 2)
    End If
        
' -------------------
' Add additional data
' -------------------
        
    'Use what's in the Task Desc. for Waiting For, if Waiting For is blank
    If strTaskSource = "Waiting" Then
        If Len(Me.txt_NextAction_Desc) > 5 And InStr(Me.txt_NextAction_Desc, " ") And Me.txt_Waiting_For = "" Then
            Me.txt_Waiting_For = StrConv(Replace(Left(Me.txt_NextAction_Desc, InStr(Me.txt_NextAction_Desc, " ") - 1), ":", ""), vbProperCase)
        End If
    End If

    'If Area = Continuous check the Continuous box
    If str_Applicable_Area = "Continuous" Then chk_Continuous = True

    'Fill the Project Names Dictionary
    For i = 2 To intLastRow_wsProjects
        dict_ProjectNames.Add key:=ws_Projects.Cells(i, col_Project_wsProjects), Item:=ws_Projects.Cells(i, col_Project_wsProjects)
    Next i

bol_EnableEvents = True

On Error GoTo 0

End Sub
Sub o_04_Create_Project_List()
   
' Purpose: To create the list of projects based on the selected Area.
' Trigger: Click List Area
' Updated: 6/28/2023

' Change Log:
'       4/12/2020:  Removed any reference to the D/A Worksheet, removed with my move to CR&A
'       9/3/2021:   Added code to default to just Active Projects, but include Pending / Continuous if they are checked
'       9/3/2021:   Added code to select the only Project if only one is in the list
'       9/4/2021:   Added code to default the chk_Continuous = True when selecting Area = Continuous
'       12/10/2021: Removed the code related to the P.XXX name, now that I include that in the Project Name
'       6/28/2023:  Added the 'Recurring' option with 'Continuous'
'       10/17/2023: Added code to abort if EnableEvents is false, so that I don't wipe the project list

' ***********************************************************************************************************************************
        
If bol_EnableEvents = False Then Exit Sub
        
' ----------------------
' Declare your variables
' ----------------------
   
    Dim strProjArea As String
    
    Dim strProjName As String
    
    Dim strProjStatus As String
    
    'ListBox variables
    Dim strArea As String
        If IsNull(lst_Area.Value) Then
            Else: strArea = lst_Area.Value
        End If
    
    Dim x As Long
        x = 2
    
' -----------------------------------------------
' If Area = "Continuous" then chk_Continuous = True
' -----------------------------------------------
    
    If Me.lst_Area.Value = "Continuous" Then chk_Continuous = True
    
' ------------
' Run the loop
' ------------
   
    lst_Project.Clear

    With ws_Projects
        Do While .Cells(x, col_Project_wsProjects).Value2 <> ""
            
            ' Assign the Variables
            strProjArea = .Cells(x, col_Area_wsProjects).Value2
            strProjName = .Cells(x, col_Project_wsProjects).Value2
            strProjStatus = .Cells(x, col_Status_wsProjects).Value2
    
            ' Add the values to the ListBox
            If strProjArea = strArea Then
            
                If strProjStatus = "Active" Or _
                (strProjStatus = "Pending" And chk_Pending = True) Or _
                ((strProjStatus = "Continuous" Or strProjStatus = "Recurring") And chk_Continuous = True) Then
                    lst_Project.AddItem (strProjName)
                End If
            
            End If
            
            x = x + 1
        
        Loop
    End With

' ------------------------------
' If only one remains, select it
' ------------------------------

    If lst_Project.ListCount = 1 Then
        lst_Project.Selected(0) = True
    End If

End Sub
Sub o_11_Add_New_Task()

' Purpose: To allow me to quickly create a new Task  in my To Do
' Trigger: Called: uf_New_NextAction
' Updated: 10/12/2023

' Change Log:
'       9/16/2019:  Rewrote while refactoring
'       6/20/2020:  Turned the center alignment off for the project
'       12/15/2020: Overhauled to use my fx_Create_Headers
'       12/15/2020: Removed the code for Due Date
'       5/14/2021:  Moved the calculation for the Task / Waiting # to be feed by the applicable workbook.
'       8/8/2021:   Added code to import the Task Note and updated Task Notes to no longer shrink to fit
'       8/12/2021:  Added the code for txt_Completed, to output the given data if I filled it in
'       10/1/2021:  Added the code to default the Time field to "<30 mins" if nothing was passed from the form
'       10/13/2021: Added the code to determine the Area if it's currently N/
'       11/2/2021:  Moved the code to Declare/Assign variables for ws_Tasks into o_021_Declare_Global_Variables
'       11/12/2021: Added code for the Component, Active Task, and Active Project fields
'       12/10/2021: Removed the code related to the P.XXX name, and bolDARequest / fx_Selected_is_DA_Request, now that I include that in the Project Name
'       1/11/2022:  Added the code for capturing the Component
'       1/19/2022:  Replaced the formating with fx_Steal_First_Row_Formating
'       1/26/2022:  Added the code so that if the project is N/A it isn't marked as Active
'       6/26/2022:  Converted 'DisableForEfficiency' to just turning ScreenUpdating on and off
'       7/10/2022:  Added the code to unhide all the columns, apply formatting, then rehide
'       9/30/2022:  Added the code to explicitly hide the blank columns
'       10/12/2023: Added the code for Active Project and Active Task and updated the Column references

' ***********************************************************************************************************************************
    
On Error Resume Next
    
Application.ScreenUpdating = False
Application.EnableEvents = False
bol_EnableEvents = False

Call Me.o_023_Declare_wsTasks_Variables

' ----------------
' Add the new Task
' ----------------
    
With ws_Tasks
    
    'Column A: Task #
        .Cells(int_CurRow, col_Ref).Value = [MAX('[To Do.xlsm]Tasks'!A:A) + 1]
       
    'Column B: Entered Date
        .Cells(int_CurRow, col_Entered).Value = Me.txt_Entered.Value
       
    'Column C: Start Date
        .Cells(int_CurRow, col_Start).Value = Me.txt_Start.Value
        
    'Column D: Context
        .Cells(int_CurRow, col_Context).Value = "'" & Me.lst_Context.Value
                        
    'Column E: Priority
        .Cells(int_CurRow, col_Priority).Value = Me.lst_Priority.Value

    'Column F: Time
        .Cells(int_CurRow, col_Time).Value = Me.lst_Time.Value
            If Me.lst_Time.Value = "" Then .Cells(int_CurRow, col_Time).Value = "<30 mins"
    
    'Column G: Area
        .Cells(int_CurRow, col_Area).Value = Me.lst_Area.Value
            If lst_Area.Value = "" Or lst_Area.Value = Null Then
                If Left(Me.lst_Project.Value, 2) = "DA" Then
                    .Cells(int_CurRow, col_Area).Value = "D/A Requests"
                Else
                    .Cells(int_CurRow, col_Area).Value = "Projects"
                End If
            End If
    
    'Column H: Project
        .Cells(int_CurRow, col_Project).Value = Me.lst_Project.Value
            If .Cells(int_CurRow, col_Project).Value = "" Then .Cells(int_CurRow, col_Project).Value = "N/A"
        
    'Column I: Component
        .Cells(int_CurRow, col_Component).Value = Me.cmb_Component.Value
            If .Cells(int_CurRow, col_Component).Value = "" Then .Cells(int_CurRow, col_Component).Value = "N/A"
        
    'Column K: Active Project
        If "Active" = ws_Projects.Cells(fx_Find_CurRow(ws:=ws_Projects, strTargetFieldName:="Project", strTarget:=Me.lst_Project.Value), 7) Then 'Col G = Status
            .Cells(int_CurRow, col_ActiveProject).Value2 = "X"
        End If
        
    'Column J: Task Description
        .Cells(int_CurRow, col_NextAction).Value = Me.txt_NextAction_Desc.Value
            
        
    'Column M: Active Task
        If Me.lst_Priority.Value = "High" Then .Cells(int_CurRow, col_ActiveTask).Value2 = "X"
            
    'Column N: Completed
        If Me.chk_Completed.Value = True Then
            If Me.txt_Completed.Value <> "" Then
                .Cells(int_CurRow, col_Completed).Value = CDate(Me.txt_Completed.Value)
            Else
                .Cells(int_CurRow, col_Completed).Value = Date
            End If
            
        End If
          
    'Column O: Task Notes
        .Cells(int_CurRow, col_Notes).Value = Me.txt_NextAction_Notes.Value
          
End With 'ws_Tasks
          
' -------------------------------
' Apply formatting to the new row
' -------------------------------

    ws_Tasks.Columns.EntireColumn.Hidden = False

    Call fx_Steal_First_Row_Formating(ws:=ws_Tasks, intSingleRow:=int_CurRow)
    
    Call u_Toggle_ActiveFields_wsTasks
    
    'ws_Tasks.Columns(col_Notes & ":" & ws_Tasks.Columns.count).EntireColumn.Hidden = True
     
     ws_Tasks.Columns(col_Notes + 1).Resize(, ws_Tasks.Columns.count - col_Notes).EntireColumn.Hidden = True
    
Application.ScreenUpdating = True
Application.EnableEvents = True
bol_EnableEvents = True

On Error GoTo 0

End Sub
Sub o_12_Add_New_Waiting()

' Purpose: To allow me to quickly create a new Waiting in my To Do
' Trigger: Called: uf_New_NextAction
' Updated: 5/19/2023

' Change Log:
'       9/16/2019:  Rewrote while refactoring
'       6/20/2020:  Turned the center alignment off for the project
'       12/15/2020: Overhauled to use my fx_Create_Headers
'       5/14/2021:  Moved the calculation for the Task / Waiting # to be feed by the applicable workbook.
'       8/8/2021:   Added code to import the Task Note and updated Task Notes to no longer shrink to fit
'       9/11/2021:  Added code to standardize people's names after being added to "Waiting For"
'       11/2/2021:  Added the field for Priority
'       11/2/2021:  Moved the code to Declare/Assign variables for ws_Waiting into o_021_Declare_Global_Variables
'       11/18/2021: Added to center the Priority data
'       1/19/2022:  Replaced the formating with fx_Steal_First_Row_Formating
'       6/26/2022:  Converted 'DisableForEfficiency' to just turning ScreenUpdating on and off
'       7/24/2022:  Updated the expansion of names to use len to determine how long the name is
'                   Added Scott R and Robin T as options
'       10/28/2022: Added Naomi S to the list
'       12/3/2022:  Replaced "Waiting For" code with fx_Name_TextExpander
'       5/12/2023:  Added the code so that if I fill out Completed it keeps the value, not default to Date
'       5/19/2023:  Added code to handle if the Completed Date is "N/A"

' ***********************************************************************************************************************************

Application.ScreenUpdating = False
Application.EnableEvents = False

Call Me.o_024_Declare_wsWaiting_Variables
    
' -------------------
' Add the new Waiting
' -------------------
    
    With ws_Waiting

    'Column A: Waiting #
        .Cells(int_CurRow, col_Ref).Value = [MAX('[To Do.xlsm]Waiting'!A:A) + 1]
        
    'Column B: Input the current date in Entered
        .Cells(int_CurRow, col_Entered).Value = Me.txt_Entered.Value
    
    'Column C: Priority
        .Cells(int_CurRow, col_Priority).Value = Me.lst_Priority.Value
        .Cells(int_CurRow, col_Priority).Validation.Add Type:=xlValidateList, Formula1:="High, Medium, Low"
    
    'Column D: Project / Area
        .Cells(int_CurRow, col_Project).Value = Me.lst_Project.Value
            If .Cells(int_CurRow, col_Project).Value = "" Then .Cells(int_CurRow, col_Project).Value = "N/A"
        
    'Column E: Task Description
        .Cells(int_CurRow, col_NextAction).Value = Me.txt_NextAction_Desc.Value
        
    'Column F: Waiting For
        .Cells(int_CurRow, col_WaitingFor).Value = myFunctions_ToDo.fx_Name_TextExpander(Me.txt_Waiting_For.Value)
            
    'Column G: Completed
        If Me.chk_Completed.Value = True Then
            If Me.txt_Completed.Value = "N/A" Then
                .Cells(int_CurRow, col_Completed).Value = "N/A"
            ElseIf Me.txt_Completed.Value <> "" Then
                .Cells(int_CurRow, col_Completed).Value = CDate(Me.txt_Completed.Value)
            Else
                .Cells(int_CurRow, col_Completed).Value = Date
            End If
            
        End If
    
    'Column H: Task Notes
        .Cells(int_CurRow, col_Notes).Value = Me.txt_NextAction_Notes.Value

  End With 'ws_Waiting
    
' -------------------------------
' Apply formatting to the new row
' -------------------------------

    Call fx_Steal_First_Row_Formating(ws:=ws_Waiting, intSingleRow:=int_CurRow)
    
Application.ScreenUpdating = True
Application.EnableEvents = True

End Sub
Sub o_13_Add_New_Question()

' Purpose: To allow me to quickly create a new Question in my To Do
' Trigger: Called: uf_New_NextAction
' Updated: 6/27/2023

' Change Log:
'       11/23/2021: Initial Creation, based on 'o_12_Add_New_Waiting'
'       1/19/2022:  Replaced the formating with fx_Steal_First_Row_Formating
'       6/26/2022:  Converted 'DisableForEfficiency' to just turning ScreenUpdating on and off
'       12/3/2022:  Replaced "Who Can Answer" code with fx_Name_TextExpander
'       6/27/2023:  Updated to exit sub if this is my Personal To Do

' ***********************************************************************************************************************************
                
#If Personal = 1 Then
    Exit Sub
#End If
    
Application.ScreenUpdating = False
Application.EnableEvents = False

Call Me.o_025_Declare_wsQuestions_Variables
    
' -------------------
' Add the new Question
' -------------------
    
On Error Resume Next
    
    With ws_Questions

    'Column A: Question #
        .Cells(int_CurRow, col_Ref).Value = [MAX('[To Do.xlsm]Questions'!A:A) + 1]
        
    'Column B: Input the current date in Entered
        .Cells(int_CurRow, col_Entered).Value = Me.txt_Entered.Value
    
    'Column C: Priority
        .Cells(int_CurRow, col_Priority).Value = Me.lst_Priority.Value
        .Cells(int_CurRow, col_Priority).Validation.Add Type:=xlValidateList, Formula1:="High, Medium, Low"
    
    'Column D: Project / Area
        .Cells(int_CurRow, col_Project).Value = Me.lst_Project.Value
            If .Cells(int_CurRow, col_Project).Value = "" Then .Cells(int_CurRow, col_Project).Value = "N/A"
        
    'Column E: Pending Qeustion
        .Cells(int_CurRow, col_NextAction).Value = Me.txt_NextAction_Desc.Value
        
    'Column F:  Who Can Answer
        .Cells(int_CurRow, col_WhoCanAnswer).Value = myFunctions_ToDo.fx_Name_TextExpander(Me.txt_Who_Can_Answer.Value)
            
    'Column G: Completed
        If Me.chk_Completed.Value = True Then .Cells(int_CurRow, col_Completed).Value = Date
    
    'Column H: Task Notes
        .Cells(int_CurRow, col_Notes).Value = Me.txt_NextAction_Notes.Value
    
  End With 'ws_Questions
  
On Error GoTo 0

' -------------------------------
' Apply formatting to the new row
' -------------------------------

    Call fx_Steal_First_Row_Formating(ws:=ws_Questions, intSingleRow:=int_CurRow)
    
Application.ScreenUpdating = True
Application.EnableEvents = False

End Sub
Sub o_2_Create_Task_Folder()

' Purpose: To create the folder that is used to house my Task support.
' Trigger: Called: uf_New_NextAction
' Updated: 6/26/2022

' Change Log:
'       9/16/2019:  Initial Creation
'       6/1/2021:   Updated to remove the Input Box, just use the name I give
'       6/2/2021:   Updated to replace Shell with FollowHyperlink
'       10/13/2021: Added the ':' in the Replace for the Folder name, was thrown off by "1:1"
'       1/26/2022:  Updated the Task Desk to be Col K
'       6/26/2022:  Converted 'DisableForEfficiency' to just turning ScreenUpdating on and off

' ***********************************************************************************************************************************

Application.ScreenUpdating = False

' ----------------------
' Declare your variables
' ----------------------

    'Dimension Values

        Dim Task_nbr As Long
            Task_nbr = ws_Tasks.Cells(int_CurRow, 1).Value

        Dim Task_desc As String
            Task_desc = ws_Tasks.Cells(int_CurRow, 10).Value
        
        Dim Hyperlink_Loc As Range
            Set Hyperlink_Loc = ws_Tasks.Cells(int_CurRow, 4) ' Task Context

    'Dimension Strings / Paths
        Dim Task_Fldr_Name As String
            'Task_Fldr_Name = InputBox(Prompt:="What would you like to call your task name", Title:="Task Folder Name", Default:=Task_desc) ' 6/1/21: Commented out
            Task_Fldr_Name = Task_desc
                Task_Fldr_Name = Replace(Replace(Replace(Replace(Replace(Task_Fldr_Name, "'", ""), Chr(13), ""), Chr(10), ""), "/", "."), ":", "")
                   
            If Task_Fldr_Name = vbNullString Then
                Application.ScreenUpdating = True
                Exit Sub
            End If
            
            Task_desc = Task_Fldr_Name
        
        Dim Task_Folder As String
            #If Personal = 0 Then
                Task_Folder = "C:\U Drive\Tasks\" & Task_nbr & ". " & Task_desc
            #ElseIf Personal = 1 Then
                Task_Folder = "D:\D Documents\Tasks\" & Task_nbr & ". " & Task_desc
            #End If

    ' -----------------------------------------------------------
    ' Create the new folder for the project based on the Project #
    ' -----------------------------------------------------------
    
        If Dir(Task_Folder) = "" Then MkDir (Task_Folder)
                         
    ' ------------------
    ' Add the hyperlinks
    ' ------------------
    
    ThisWorkbook.Sheets("Tasks").Hyperlinks.Add Anchor:=Hyperlink_Loc, Address:=Task_Folder
    
    'This fixes the issue of hyperlinks being blue and underlined by making all font black, no underline, and Size 11 Calibri
    
        With Hyperlink_Loc.Font
            '.ColorIndex = xlAutomatic
            .Color = RGB(200, 200, 200)
            .Underline = xlUnderlineStyleNone
            .Name = "Calibri"
            .Size = 11
        End With
        
    'Open the applicable folder
        ThisWorkbook.FollowHyperlink (Task_Folder)
        
Application.ScreenUpdating = True

End Sub
Sub o_31_Next_Action_is_Task()

' Purpose: To udpate the UserForm to just the fields required for a new Task.
' Trigger: Event: lst_NextAction_Click > Task
' Updated: 9/24/2023

' Change Log:
'       11/1/2021: Removed the code related to Priority, since I now use it for Waitings.
'       11/23/2021: Added the code for Questions
'       12/24/2021: Moved code from the Next Action Option button click
'       1/11/2022:  Added the code for the Component ComboBox
'       9/24/2023:  If Tasks then show the lst_Area as it is used

' ***********************************************************************************************************************************
   
' ---------------------------------
' Turn on the Tasks related objects
' ---------------------------------
      
strNextActionType = "Task"

    'Turn ON the Task related functionality
    Me.lbl_Start.Visible = True
    Me.txt_Start.Visible = True
    
    Me.lbl_Time.Visible = True
    Me.lst_Time.Visible = True
    
    Me.lbl_Context.Visible = True
    Me.lst_Context.Visible = True
    
    Me.lbl_Component.Visible = True
    Me.cmb_Component.Visible = True
    
    Me.lst_Area.Visible = True
      
' -------------------------------------
' Turn off the <> Tasks related objects
' -------------------------------------
      
    'Turn OFF the Waiting related functionality
    Me.lbl_WaitingFor.Visible = False
    Me.txt_Waiting_For.Visible = False

    'Turn OFF the Questions related functionality
    Me.lbl_WhoCanAnswer.Visible = False
    Me.txt_Who_Can_Answer.Visible = False

End Sub
Sub o_32_Next_Action_is_Waiting()

' Purpose: To udpate the UserForm to just the fields required for a new Waiting.
' Trigger: Event: lst_NextAction_Click > Waiting
' Updated: 9/24/2023

' Change Log:
'       11/1/2021: Removed the code related to Priority, since I now use it for Waitings.
'       11/23/2021: Added the code for Questions
'       12/24/2021: Moved code from the Next Action Option button click
'       1/11/2022:  Added the code for the Component ComboBox
'       9/24/2023:  If Waiting then hide the lst_Area as it isn't used

' ***********************************************************************************************************************************

' -----------------------------------
' Turn on the Waiting related objects
' -----------------------------------

strNextActionType = "Waiting"

    'Turn ON the Waiting related functionality
    Me.lbl_WaitingFor.Visible = True
    Me.txt_Waiting_For.Visible = True

' ---------------------------------------
' Turn off the <> Waiting related objects
' ---------------------------------------

    'Turn OFF the Task related functionality
    Me.lbl_Start.Visible = False
    Me.txt_Start.Visible = False
    
    Me.lbl_Time.Visible = False
    Me.lst_Time.Visible = False
    
    Me.lbl_Context.Visible = False
    Me.lst_Context.Visible = False
    
    Me.lbl_Component.Visible = False
    Me.cmb_Component.Visible = False

    'Turn OFF the Questions related functionality
    Me.lbl_WhoCanAnswer.Visible = False
    Me.txt_Who_Can_Answer.Visible = False
    
    Me.lst_Area.Visible = False

End Sub
Sub o_33_Next_Action_is_Question()

' Purpose: To udpate the UserForm to just the fields required for a new Question.
' Trigger: Event: lst_NextAction_Click > Question
' Updated: 9/24/2023

' Change Log:
'       11/23/2021: Initial Creation
'       12/24/2021: Moved code from the Next Action Option button click
'       1/11/2022:  Added the code for the Component ComboBox
'       9/24/2023:  If Question then hide the lst_Area as it isn't used

' ***********************************************************************************************************************************

' -------------------------------------
' Turn on the Questions related objects
' -------------------------------------

strNextActionType = "Question"

'Turn ON the Questions related functionality
    Me.lbl_WhoCanAnswer.Visible = True
    Me.txt_Who_Can_Answer.Visible = True

' -----------------------------------------
' Turn off the <> Questions related objects
' -----------------------------------------

'Turn OFF the Task related functionality
    Me.lbl_Start.Visible = False
    Me.txt_Start.Visible = False
    
    Me.lbl_Time.Visible = False
    Me.lst_Time.Visible = False
    
    Me.lbl_Context.Visible = False
    Me.lst_Context.Visible = False

    Me.lbl_Component.Visible = False
    Me.cmb_Component.Visible = False

'Turn OFF the Waiting related functionality
    Me.lbl_WaitingFor.Visible = False
    Me.txt_Waiting_For.Visible = False
    
    Me.lst_Area.Visible = False

End Sub

