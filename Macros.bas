Attribute VB_Name = "Macros"
' Declare Worksheets
Dim ws_Projects As Worksheet
Dim ws_Tasks As Worksheet
Dim ws_Waiting As Worksheet
Dim ws_Questions As Worksheet
Dim ws_Recurring As Worksheet

Dim ws_Daily As Worksheet
Dim ws_Temp As Worksheet

' Declare Integers
Dim int_CurRow As Long
Dim int_LastRow As Long
Dim int_LastCol As Long

' Declare Common Cell References
Dim arry_Header() As Variant
       
Dim col_Priority As Long
Dim col_Area As Long
Dim col_Project As Long
Dim col_Task As Long
Dim col_Status As Long
Dim col_Completed As Long

' Declare ws_Projects Cell References
Dim col_Updated_wsProjects As Long

' Declare ws_Tasks Cell References
Dim col_Start_wsTasks As Long
Dim col_Component_wsTasks As Long

Dim col_ActiveProject_wsTasks As Long
Dim col_ActiveComponent_wsTasks As Long
Dim col_ActiveTask_wsTasks As Long

' Declare ws_Waiting Cell References
Dim col_WaitingFor_wsWaiting As Long

' Declare ws_Recurring Cell References
Dim col_LastCompleted_ws_Recurring As Long

Option Explicit
Sub o_02_Assign_Private_Variables()

' Purpose: To assign the Private Variables that were Declared above the line.
' Trigger: Called by various procedures
' Updated: 11/21/2021
' Reviewd: 5/11/2023

' Change Log:
'       12/17/2020: Intial Creation
'       11/21/2021: Added ws_Questions

' ***********************************************************************************************************************************

' ----------------
' Assign Variables
' ----------------

    Set ws_Projects = ThisWorkbook.Sheets("Projects")
    Set ws_Tasks = ThisWorkbook.Sheets("Tasks")
    Set ws_Waiting = ThisWorkbook.Sheets("Waiting")
    Set ws_Questions = ThisWorkbook.Sheets("Questions")
    Set ws_Recurring = ThisWorkbook.Sheets("Recurring")
    Set ws_Daily = ThisWorkbook.Sheets("Daily")
    Set ws_Temp = ThisWorkbook.Sheets("Temp")

End Sub
Sub o_03_Assign_Private_Variables_wsProjects()

' Purpose: To assign the Private Variables that were Declared above the line related to ws_Projects.
' Trigger: Called by various procedures
' Updated: 11/1/2021
' Reviewd: 5/11/2023

' Change Log:
'       11/1/2021: Intial Creation

' ***********************************************************************************************************************************

' ----------------
' Assign Variables
' ----------------
    
    ' Assign Integers
    int_LastCol = ws_Projects.Cells(1, ws_Projects.Columns.count).End(xlToLeft).Column

    ' Assign Cell References
    arry_Header = Application.Transpose(ws_Projects.Range(ws_Projects.Cells(1, 1), ws_Projects.Cells(1, int_LastCol)))
    
    col_Area = fx_Create_Headers("Area", arry_Header)
    col_Project = fx_Create_Headers("Project", arry_Header)
    col_Updated_wsProjects = fx_Create_Headers("Updated", arry_Header)
    col_Status = fx_Create_Headers("Status", arry_Header)

End Sub
Sub o_04_Assign_Private_Variables_wsTasks()

' Purpose: To assign the Private Variables that were Declared above the line related to ws_Tasks.
' Trigger: Called by various procedures
' Updated: 7/3/2022
' Reviewd: 5/11/2023

' Change Log:
'       12/28/2020: Intial Creation
'       11/2/2021:  Added the field for Active Project, Active Component, and Active Task
'       7/3/2022:   Updated the int_CurRow to use the Find_CurRow function

' ***********************************************************************************************************************************

' ----------------
' Assign Variables
' ----------------
    
    ' Assign Integers
    int_CurRow = fx_Find_CurRow(ws:=ws_Tasks, strTargetFieldName:="Task", strTarget:="")
    int_LastCol = ws_Tasks.Cells(1, ws_Tasks.Columns.count).End(xlToLeft).Column

    ' Assign Cell References
    arry_Header = Application.Transpose(ws_Tasks.Range(ws_Tasks.Cells(1, 1), ws_Tasks.Cells(1, int_LastCol)))
       
    col_Start_wsTasks = fx_Create_Headers("Start", arry_Header)
    col_Priority = fx_Create_Headers("Priority", arry_Header)
    col_Area = fx_Create_Headers("Area", arry_Header)
    col_Project = fx_Create_Headers("Project", arry_Header)
    col_Task = fx_Create_Headers("Task", arry_Header)
    col_Completed = fx_Create_Headers("Completed", arry_Header)
    
    col_Component_wsTasks = fx_Create_Headers("Component", arry_Header)
    col_ActiveProject_wsTasks = fx_Create_Headers("Active Proj.", arry_Header)
    col_ActiveComponent_wsTasks = fx_Create_Headers("Active Comp.", arry_Header)
    col_ActiveTask_wsTasks = fx_Create_Headers("Active Task", arry_Header)

End Sub
Sub o_05_Assign_Private_Variables_wsWaiting()

' Purpose: To assign the Private Variables that were Declared above the line related to ws_Waiting.
' Trigger: Called by various procedures
' Updated: 12/28/2020
' Reviewd: 5/11/2023

' Change Log:
'       12/28/2020: Intial Creation

' ***********************************************************************************************************************************

' ----------------
' Assign Variables
' ----------------
    
    ' Assign Integers
    int_LastCol = ws_Waiting.Cells(1, ws_Waiting.Columns.count).End(xlToLeft).Column

    ' Assign Cell References
    arry_Header = Application.Transpose(ws_Waiting.Range(ws_Waiting.Cells(1, 1), ws_Waiting.Cells(1, int_LastCol)))
       
    col_Priority = fx_Create_Headers("Priority", arry_Header)
    col_Project = fx_Create_Headers("Project", arry_Header)
    col_Task = fx_Create_Headers("Task", arry_Header)
    col_WaitingFor_wsWaiting = fx_Create_Headers("Waiting For", arry_Header)
    col_Completed = fx_Create_Headers("Completed", arry_Header)

End Sub
Sub o_06_Assign_Private_Variables_wsQuestions()

' Purpose: To assign the Private Variables that were Declared above the line related to ws_Questions.
' Trigger: Called by various procedures
' Updated: 6/27/2023
' Reviewd: 5/11/2023

' Change Log:
'       11/21/2021: Intial Creation
'       6/27/2023:  Updated to exit sub if this is my Personal To Do

' ***********************************************************************************************************************************

#If Personal = 1 Then
    Exit Sub
#End If

' ----------------
' Assign Variables
' ----------------
    
    ' Assign Integers
    int_LastCol = ws_Questions.Cells(1, ws_Questions.Columns.count).End(xlToLeft).Column

    ' Assign Cell References
    arry_Header = Application.Transpose(ws_Questions.Range(ws_Questions.Cells(1, 1), ws_Questions.Cells(1, int_LastCol)))
       
    col_Priority = fx_Create_Headers("Priority", arry_Header)
    col_Project = fx_Create_Headers("Project", arry_Header)
    col_Completed = fx_Create_Headers("Completed", arry_Header)

End Sub
Sub o_07_Assign_Private_Variables_ws_Recurring()

' Purpose: To assign the Private Variables that were Declared above the line related to ws_Recurring.
' Trigger: Called by various procedures
' Updated: 3/11/2022
' Reviewd: 5/11/2023

' Change Log:
'       11/1/2021:  Intial Creation
'       3/11/2022:  Added Priority

' ***********************************************************************************************************************************

' ----------------
' Assign Variables
' ----------------
    
    ' Assign Integers
    int_LastCol = ws_Recurring.Cells(1, ws_Recurring.Columns.count).End(xlToLeft).Column

    ' Assign Cell References
    arry_Header = Application.Transpose(ws_Recurring.Range(ws_Recurring.Cells(1, 1), ws_Recurring.Cells(1, int_LastCol)))
       
    col_Task = fx_Create_Headers("Task", arry_Header)
    col_Priority = fx_Create_Headers("Priority", arry_Header)
    col_Status = fx_Create_Headers("Status", arry_Header)
    col_LastCompleted_ws_Recurring = fx_Create_Headers("Last Completed", arry_Header)

End Sub
Sub o_11_Add_a_New_Project()
Attribute o_11_Add_a_New_Project.VB_ProcData.VB_Invoke_Func = "P\n14"

' Purpose: To allow me to quickly input a new Project into my To Do.
' Trigger: Keyboard Shortcut - Ctrl + Shift + P
' Updated: 12/16/2020
' Reviewd: 5/11/2023

' Change Log
'       3/31/2018:  Initial Creation was sometime in Q1 2018
'       12/16/2020: Added the code to SetFocus on Project on initialization

' ***********************************************************************************************************************************
        
' -----------------
' Load the UserForm
' -----------------
        
    uf_New_Project.Show vbModeless

    ' Force the Task Descrption object to take Focus
    
    uf_New_Project.frm_Project.Enabled = False
        uf_New_Project.frm_Project.Enabled = True
        
    uf_New_Project.txt_Project.SetFocus

End Sub
Sub o_12_Add_a_New_DA_Request()
Attribute o_12_Add_a_New_DA_Request.VB_ProcData.VB_Invoke_Func = "D\n14"

' Purpose: To allow me to quickly create a new Data Analytics request for my To Do Excel via UserForm
' Trigger: Keyboard Shortcut - Ctrl + Shift + D
' Updated: 12/22/2021
' Reviewd: 5/11/2023

' Change Log:
'       10/26/2020: Updated to not trigger if I am on my personal computer
'       12/16/2020: Added the conditional compiler constant to abort if this is my Personal To Do
'       12/16/2020: Added the code to SetFocus on the Requestor on initialization
'       12/22/2021: Removed the 'o_13_Filter_wsProjects' code

' ***********************************************************************************************************************************
    
#If Personal = 1 Then
    Exit Sub
#End If
    
' -----------------
' Load the UserForm
' -----------------
        
    uf_New_DARequest.Show vbModeless

    ' Force the Requestor field to take Focus
    uf_New_DARequest.txt_Requestor.Enabled = False
        uf_New_DARequest.txt_Requestor.Enabled = True
    
    uf_New_DARequest.txt_Requestor.SetFocus
    
End Sub
Sub o_13_Filter_wsProjects()

' Purpose: To allow me to filter down to just the Active Projects for my To Do
' Trigger: Called: o_51_Reset_To_Do
' Updated: 8/29/2023
' Reviewd: 6/29/2023

' Change Log:
'       9/30/2020:  Updated to remove some redundant code
'       12/28/2020: Moved the sort code from the To Do Reset
'       6/15/2021:  Changed the sort to make it match the Weekly Reset
'       7/27/2021:  Removed 'D/A Strategy'
'       11/1/2021:  Removed Assigning the Public Variables, and Disabling For Efficiency, it was redundant
'       11/9/2021:  Updated to reflect new Personal Areas of Focus
'       5/7/2022:   Converted "House / Yard" and "Financial" to "Household"
'       6/29/2023:  Updated the sorts to match new Areas of Focus
'       8/29/2023:  Updated to include the new 'Infrastructure' and 'Strategy' options for my new role

' ***********************************************************************************************************************************
    
Call Macros.o_03_Assign_Private_Variables_wsProjects
    
' --------------------------
' Filter and sort ws_Projects
' --------------------------
    
    With ws_Projects

    ' Filter down to only incomplete Projects
        If .AutoFilterMode = True Then .AutoFilter.ShowAllData
            int_LastRow = .Cells(Rows.count, "A").End(xlUp).Row
        
        .Range("A1").Sort _
            Key1:=.Cells(1, col_Area), Order1:=xlAscending, _
            Key2:=.Cells(1, col_Project), Order2:=xlAscending, _
            Header:=xlYes
            
    ' Sort the data
        
        .Sort.SortFields.Clear

        #If Personal <> 1 Then
            .Sort.SortFields.Add key:=Range("C:C"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
                "D/A Requests, Projects, Infrastructure, Strategy, Recurring, Continuous, Personal", DataOption:=xlSortNormal
        #Else
            .Sort.SortFields.Add key:=Range("C:C"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
                "Family, Household, Yard, Finances, Personal, Continuous", DataOption:=xlSortNormal
        #End If

        .Sort.SetRange Rows("1:" & int_LastRow)
        .Sort.Header = xlYes
        .Sort.Apply
                                                   
    ' Apply the filter for Active only
        .Range("A1").AutoFilter Field:=col_Status, Criteria1:=Array("Active", "="), Operator:=xlFilterValues
                                                   
    ' Call the Apply Area Formatting Macro to apply the color formatting
        Call Macros.o_15_Apply_Projects_Formatting
            
    ' Hide the unused rows in ws_Projects
        '.Rows((int_LastRow + 1) & ":" & (int_LastRow + 1)).Hidden = False ' Disabled on 5/19/23
        .Rows((int_LastRow + 1) & ":" & (Rows.count)).Hidden = True ' Renabled on 9/1/23
    
    ' Wipe out the selection
        Application.GoTo .Range("A1"), False
        
    End With
              
End Sub
Sub o_14_Open_Project_Folder()
Attribute o_14_Open_Project_Folder.VB_ProcData.VB_Invoke_Func = "O\n14"

' Purpose: To allow me to open project support on the fly, or filter my To Do to only tasks related to the selected project.
' Trigger: Keyboard Shortcut - Ctrl + Shift + O
' Updated: 2/3/2022
' Reviewd: 5/11/2023

' Change Log:
'       2/28/2020: Switched from modeless to modal to allow me to SetFocus
'       10/8/2021: Switched back to modal, using the code to force focus
'       2/3/2022:  Switched to having the Dynamic Lookup take focus, so I can just start typing a project name

' ***********************************************************************************************************************************

' -----------------
' Load the UserForm
' -----------------

     uf_Project_Selector.Show vbModeless

    ' Force the Dynamic Search object to take Focus
    
    uf_Project_Selector.frm_Project_Lookup.Enabled = False
    uf_Project_Selector.frm_Project_Lookup.Enabled = True
    
    uf_Project_Selector.cmb_DynamicSearch.SetFocus
    
End Sub
Sub o_15_Apply_Projects_Formatting()

' Purpose: To apply the conditional formatting to each Project.
' Trigger: Called: o_51_Reset_To_Do
' Updated: 9/1/2023
' Reviewd: 5/11/2023

' Change Log:
'       12/18/2020: Turned off the "DisableForEfficiency" as this is called by the Reset and was messing things up.
'       5/11/2023:  Removed the redundant code for ws_Projects, replaced with ws_Projects
'       6/28/2023:  Updated to steal the approach used in o_22_Create_Project_Folder, replacing parts of that code
'                   Moved some of the formatting code out of the IF Statement
'       9/1/2023:   Updated to include the 'Strategy' and 'Infrastructure' options
'                   Updated to break the colors into seperate variables

' ***********************************************************************************************************************************

On Error Resume Next

' -----------------
' Declare Variables
' -----------------
    
'   Dim int_LastRow As Long
        int_LastRow = ws_Projects.Cells(Rows.count, "B").End(xlUp).Row
    
    Dim curRow As Long

    Dim strProjectArea As String
    
    'Declare Colors
    
    Dim clrGreenLight As Long
        clrGreenLight = RGB(234, 245, 211)
    
    Dim clrBlueLight As Long
        clrBlueLight = RGB(237, 242, 249)
    
    Dim clrBlueDark As Long
        clrBlueDark = RGB(225, 233, 243)
    
    Dim clrRedSalmon As Long
        clrRedSalmon = RGB(245, 228, 227)
    
    Dim clrPurpleLight As Long
        clrPurpleLight = RGB(240, 235, 248)
        
' ----------------------------------------
' Apply the formatting to each visible row
' ----------------------------------------

With ws_Projects

    For curRow = 2 To int_LastRow
    
        If Not .Rows(curRow).Hidden Then
            If .Cells(curRow, "B").Interior.Color = xlNone Or .Cells(curRow, "B").Interior.Color = RGB(255, 255, 255) Then
                
                strProjectArea = .Cells(curRow, "C").Value2
         
                If (strProjectArea = "D/A Requests" Or strProjectArea = "Family") Then
                    .Range(.Cells(curRow, "B"), .Cells(curRow, "H")).Interior.Color = clrGreenLight
                
                ElseIf (strProjectArea = "Projects" Or strProjectArea = "Household" Or strProjectArea = "Yard") Then
                    .Range(.Cells(curRow, "B"), .Cells(curRow, "H")).Interior.Color = clrBlueLight
                        
                ElseIf (strProjectArea = "Finances" Or strProjectArea = "Strategy" Or strProjectArea = "Infrastructure") Then
                    .Range(.Cells(curRow, "B"), .Cells(curRow, "H")).Interior.Color = clrBlueDark
                                                
                ElseIf strProjectArea = "Personal" Then
                    .Range(.Cells(curRow, "B"), .Cells(curRow, "H")).Interior.Color = clrPurpleLight
                        
                ElseIf (strProjectArea = "Continuous" Or strProjectArea = "Recurring") Then
                    .Range(.Cells(curRow, "B"), .Cells(curRow, "H")).Interior.Color = clrRedSalmon
                    
                End If
            
                'Adjust the row height
                    .Rows(curRow).RowHeight = 45
                    
            End If
    
        End If
                
        'This fixes the issue of hyperlinks being blue and underlined by making all font black, no underline, and Size 11 Calibri
        With .Range(.Cells(curRow, "B"), .Cells(curRow, "H")).Font
            .ColorIndex = xlAutomatic
            .Underline = xlUnderlineStyleNone
            .Name = "Calibri"
            .Size = 11
        End With
    
        'This applies the grey horizontal lines
        With .Range(.Cells(curRow, "B"), .Cells(curRow, "H"))
             .Borders(xlEdgeBottom).Color = RGB(217, 217, 217)
             .Borders(xlInsideHorizontal).Color = RGB(217, 217, 217)
             .Borders(xlEdgeTop).Color = RGB(217, 217, 217)
             .Font.Name = "Calibri"
             .Font.Size = 11
        End With

    Next curRow

End With
    
End Sub
Sub o_21_Add_a_New_Task()
Attribute o_21_Add_a_New_Task.VB_ProcData.VB_Invoke_Func = "T\n14"

' Purpose: To allow me to quickly create a new Task for my To Do Excel via UserForm
' Trigger: Keyboard Shortcut - Ctrl + Shift + T
' Updated: 6/26/2022
' Reviewd: 5/11/2023

' Change Log:
'       3/31/2018:  Initial Creation was sometime in Q1 2018
'       12/16/2020: Added the code to SetFocus on the Task Description on initialization
'       6/20/2021:  Added the code to reset the Autofilter, not run the o_22_Filter_wsTasks
'       11/23/2021: Removed the AutoFilter code, no longer needed
'       12/24/2021: Updated to replace the option button w/ lst_NextAction
'       5/23/2022:  Updated to include the unique code for 'Daily' to apply the New Action formatting
'       6/26/2022:  Removed 'DisableForEfficiency' as it was redundant w/ code in the UserForm

' ***********************************************************************************************************************************
    
If ActiveSheet.Name = "Daily" Then GoTo NextActionFormatting
    
' -----------------
' Declare Variables
' -----------------
    
    Dim CurrLocation As Range
    Set CurrLocation = Selection  'Set this initially
        
' ----------------------------------------------------
' If ws_Daily apply Next Action formatting to selection
' ----------------------------------------------------

NextActionFormatting:
    
    If ActiveSheet.Name = "Daily" Then
        Call Macros.o_63_Update_Meeting_Notes_Type("Task")
        Exit Sub
    End If
    
' -----------------
' Load the UserForm
' -----------------
            
    uf_New_NextAction.Show vbModeless
        uf_New_NextAction.lst_NextAction.Value = "Task"
    
    ' Force the Task Descrption object to take Focus
    
    uf_New_NextAction.frm_TaskDesc.Enabled = False
        uf_New_NextAction.frm_TaskDesc.Enabled = True
        
    uf_New_NextAction.txt_NextAction_Desc.SetFocus
    
Application.GoTo CurrLocation, False
    
End Sub
Sub o_22_Filter_wsTasks()

' Purpose: To allow me to filter down to just the incomplete Tasks for my To Do
' Trigger: Called: o_51_Reset_To_Do
' Updated: 8/29/2023
' Reviewd: 5/11/2023

' Change Log:
'       12/22/2019: I removed the "Waiting On" criteria for Context as I no longer use that context
'       9/30/2020:  Updated to remove some redundant code
'       12/15/2020: Added in the filter to include just High or Medium priority tasks
'       12/27/2020: Added the code to unhide the row after the last
'       12/28/2020: Moved the sort code from the To Do Reset
'       7/27/2021:  Removed 'D/A Strategy'
'       11/1/2021:  Updated to explicitly reference the Columns
'       11/1/2021:  Removed Assigning the Public Variables, and Disabling For Efficiency, it was redundant
'       11/2/2021:  Added the code to set the Active Task field
'       11/2/2021:  Updated the Autofilter to filter based on Active Task instead of Priority and Active Project
'       11/9/2021:  Updated to reflect new Personal Areas of Focus
'       11/12/2021: Updated to NOT flag if the task is complete
'       11/22/2021: Added the ability to mark a task as <> Active if the Active Component is No
'       12/22/2021: Switched the loop to use int_LastRow not int_CurRow (was causing it to abort)
'       1/6/2022:   Added the sort for col_Component_wsTasks
'       1/10/2022:  Added the arry(col_Priority <> "Low") to remove the lows from Active
'       5/7/2022:   Converted "House / Yard" and "Financial" to "Household"
'       11/13/2022: Updated to remove tasks that are not part of an Active Project
'       5/11/2023:  Updated the name of arry => arryAllTasks
'       6/13/2023:  Switched so filtering the data occurs after the sorting, so my non-Active Tasks get sorted correctly
'       6/29/2023:  Updated the sorts to match new Areas of Focus
'       8/29/2023:  Updated to include the new 'Infrastructure' and 'Strategy' options for my new role

' ***********************************************************************************************************************************
    
Call Macros.o_04_Assign_Private_Variables_wsTasks
    
With ws_Tasks
    
' --------------------
' Reset the Autofilter
' --------------------
        
    If .AutoFilterMode = True Then .AutoFilter.ShowAllData
    
' -----------------
' Declare Variables
' -----------------
    
    Dim i As Long
    
    Dim arryAllTasks() As Variant
        arryAllTasks = Application.Transpose(.UsedRange)
        
    Dim arryActiveTask() As Variant
        arryActiveTask = Application.Transpose(.Range(.Cells(1, col_ActiveTask_wsTasks), .Cells(int_CurRow, col_ActiveTask_wsTasks)))
    
    int_LastRow = .Cells(Rows.count, "A").End(xlUp).Row
    
' ------------------------------------------------------------------------------------------
' Determine the Active Task field, based on High Priority Tasks or tasks for Active Projects
' ------------------------------------------------------------------------------------------
    
    For i = 2 To int_LastRow
    
        If arryAllTasks(col_Completed, i) = "" Then
        
            If arryAllTasks(col_Priority, i) = "High" Then
                arryActiveTask(i) = "X"
            ElseIf arryAllTasks(col_ActiveProject_wsTasks, i) = "X" And arryAllTasks(col_ActiveComponent_wsTasks, i) <> "N" And arryAllTasks(col_Priority, i) <> "Low" Then
                arryActiveTask(i) = "X"
            Else
                arryActiveTask(i) = ""
            End If
        Else
            arryActiveTask(i) = "" ' Remove the flag if the task was completed
        End If
    
    Next i
    
    .Range(.Cells(1, col_ActiveTask_wsTasks), .Cells(int_CurRow, col_ActiveTask_wsTasks)) = Application.Transpose(arryActiveTask)
    
' ---------------
' Filter my Tasks
' ---------------

    ' Sort the data on Project Area then Project Name

        .Sort.SortFields.Clear ' Need to include this so that you don't just keep adding sorts
                
        #If Personal <> 1 Then
            .Sort.SortFields.Add key:=Cells(1, col_Area), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
                "D/A Requests, Projects, Infrastructure, Strategy, Recurring, Continuous, Personal", DataOption:=xlSortNormal
        #Else
            .Sort.SortFields.Add key:=Cells(1, col_Area), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
                "Family, Household, Yard, Finances, Personal, Continuous", DataOption:=xlSortNormal
        #End If
                
        .Sort.SortFields.Add key:=Cells(1, col_Project), Order:=xlAscending
        .Sort.SortFields.Add key:=Cells(1, col_Component_wsTasks), Order:=xlAscending
        
        .Sort.SetRange Rows("1:" & int_LastRow)
        .Sort.Header = xlYes
        .Sort.Apply
            
    ' Filter the data to just Active Tasks
        
        With .Range("A1")
            .AutoFilter Field:=col_Start_wsTasks, Criteria1:="<" & Date, Operator:=xlOr, Criteria2:="="
            .AutoFilter Field:=col_ActiveTask_wsTasks, Criteria1:="X"
            .AutoFilter Field:=col_Completed, Criteria1:=""
        End With
                        
    ' Hide the unused rows in the Tasks
        .Rows((int_LastRow + 1) & ":" & (int_LastRow + 1)).Hidden = False
        .Rows((int_LastRow + 2) & ":" & (Rows.count)).Hidden = True

    ' Wipe out the selection
        Application.GoTo .Range("A1"), False

End With

End Sub
Sub o_31_Add_a_New_Waiting()
Attribute o_31_Add_a_New_Waiting.VB_ProcData.VB_Invoke_Func = "W\n14"

' Purpose: To allow me to quickly create a new Task for my To Do Excel via UserForm
' Trigger: Keyboard Shortcut - Ctrl + Shift + W
' Updated: 6/26/2022
' Reviewd: 5/11/2023

' Change Log:
'       3/31/2018:  Initial Creation was sometime in Q1 2018
'       12/16/2020: Added the code to SetFocus on the Task Description on initialization
'       8/15/2021:  Updated to force the Waiting For to take focus
'       11/1/2021:  Added the step to apply the AutoFilter before starting the form
'       11/23/2021: Removed the AutoFilter code, no longer needed
'       12/24/2021: Updated to replace the option button w/ lst_NextAction
'       5/23/2022:  Updated to include the unique code for 'Daily' to apply the New Action formatting
'       6/26/2022:  Removed 'DisableForEfficiency' as it was redundant w/ code in the UserForm

' ***********************************************************************************************************************************
    
If ActiveSheet.Name = "Daily" Then GoTo NextActionFormatting
    
' -----------------
' Declare Variables
' -----------------
    
    Dim CurrLocation As Range
    Set CurrLocation = Selection  'Set this initially
    
' ----------------------------------------------------
' If ws_Daily apply Next Action formatting to selection
' ----------------------------------------------------

NextActionFormatting:
    
    If ActiveSheet.Name = "Daily" Then
        Call Macros.o_63_Update_Meeting_Notes_Type("Waiting")
        Exit Sub
    End If
    
' -----------------
' Load the UserForm
' -----------------
    
    uf_New_NextAction.Show vbModeless
        uf_New_NextAction.lst_NextAction.Value = "Waiting"
    
    ' Force the Waiting For to take Focus
    
    uf_New_NextAction.txt_Waiting_For.Enabled = False
        uf_New_NextAction.txt_Waiting_For.Enabled = True
        
    uf_New_NextAction.txt_Waiting_For.SetFocus

Application.GoTo CurrLocation, False

End Sub
Sub o_32_Filter_wsWaiting()

' Purpose: To allow me to filter down to just the incomplete Waiting for my To Do
' Trigger: Called: o_51_Reset_To_Do
' Updated: 10/25/2022
' Reviewd: 5/11/2023

' Change Log:
'       12/27/2020: Added the code to unhide the row after the last
'       12/28/2020: Updated to sort by Project
'       4/26/2021:  Switched to sort by who I am waiting on, not based on project
'       11/1/2021:  Updated to explicitly reference the Columns
'       11/1/2021:  Removed Assigning the Public Variables, and Disabling For Efficiency, it was redundant
'       11/13/2021: Updated the Sort and AutoFilter to use "A1"
'       10/25/2022: Switched to sorting based on 'Project' instead of 'Waiting For'

' ***********************************************************************************************************************************
    
Call Macros.o_05_Assign_Private_Variables_wsWaiting

With ws_Waiting

' -------------------------
' Filter and sort ws_Waiting
' -------------------------
        
    ' Filter down to only incomplete Waiting
        If .AutoFilterMode = True Then .AutoFilter.ShowAllData
            int_LastRow = .Cells(Rows.count, "A").End(xlUp).Row
        
        With .Range("A1")
            .AutoFilter Field:=col_Completed, Criteria1:="="
            .AutoFilter Field:=col_Priority, Criteria1:=Array("High", "Medium", "="), Operator:=xlFilterValues
        End With

    ' Sort the data on Project
        .Range("A1").Sort Key1:=.Cells(1, col_Project), Order1:=xlAscending, Header:=xlYes
            
    ' Hide the unused rows in the Next Actions tab
        .Rows((int_LastRow + 1) & ":" & (int_LastRow + 1)).Hidden = False
        .Rows((int_LastRow + 2) & ":" & (Rows.count)).Hidden = True
              
    ' Wipe out the selection
        Application.GoTo .Range("A1"), False
              
End With
    
End Sub
Sub o_41_Add_a_New_Question()
Attribute o_41_Add_a_New_Question.VB_ProcData.VB_Invoke_Func = "N\n14"

' Purpose: To allow me to quickly create a new Question for my To Do Excel via UserForm
' Trigger: Keyboard Shortcut - Ctrl + Shift + N
' Updated: 6/27/2023
' Reviewd: 5/11/2023

' Change Log:
'       11/23/2021: Initial Creation, based on o_21_Add_a_New_Task
'       12/24/2021: Updated to replace the option button w/ lst_NextAction
'       5/23/2022:  Updated to include the unique code for 'Daily' to apply the New Action formatting
'       6/26/2022:  Removed 'DisableForEfficiency' as it was redundant w/ code in the UserForm
'       6/27/2023:  Updated to exit sub if this is my Personal To Do

' ***********************************************************************************************************************************

#If Personal = 1 Then
    Exit Sub
#End If
    
If ActiveSheet.Name = "Daily" Then GoTo NextActionFormatting
    
' -----------------
' Declare Variables
' -----------------
    
    Dim CurrLocation As Range
    Set CurrLocation = Selection  'Set this initially
    
' ----------------------------------------------------
' If ws_Daily apply Next Action formatting to selection
' ----------------------------------------------------
        
NextActionFormatting:
        
    If ActiveSheet.Name = "Daily" Then
        Call Macros.o_63_Update_Meeting_Notes_Type("Question")
        Exit Sub
    End If
        
' -----------------
' Load the UserForm
' -----------------

    uf_New_NextAction.Show vbModeless
        uf_New_NextAction.lst_NextAction.Value = "Question"
    
    ' Force the Task Descrption object to take Focus
    
    uf_New_NextAction.frm_TaskDesc.Enabled = False
        uf_New_NextAction.frm_TaskDesc.Enabled = True
        
    uf_New_NextAction.txt_NextAction_Desc.SetFocus
    
Application.GoTo CurrLocation, False
    
End Sub
Sub o_42_Filter_wsQuestions()

' Purpose: To allow me to filter down to just the incomplete Questions for my To Do
' Trigger: Called: o_51_Reset_To_Do
' Updated: 6/27/2023
' Reviewd: 5/11/2023

' Change Log:
'       11/21/2020: Initial Creation, based on o_32_Filter_wsWaiting
'       10/25/2022: Started sorting based on 'Project'
'       6/27/2023:  Updated to exit sub if this is my Personal To Do

' ***********************************************************************************************************************************
    
#If Personal = 1 Then
    Exit Sub
#End If
    
Call Macros.o_06_Assign_Private_Variables_wsQuestions

' -------------------------
' Filter and sort ws_Waiting
' -------------------------
    
    With ws_Questions
        
    ' Filter down to only incomplete Waiting
        If .AutoFilterMode = True Then .AutoFilter.ShowAllData
            int_LastRow = .Cells(Rows.count, "A").End(xlUp).Row
        
        With .Range("A1")
            .AutoFilter Field:=col_Completed, Criteria1:="="
            .AutoFilter Field:=col_Priority, Criteria1:=Array("High", "Medium", "="), Operator:=xlFilterValues
        End With
            
    ' Sort the data on Project
        .Range("A1").Sort Key1:=.Cells(1, col_Project), Order1:=xlAscending, Header:=xlYes
            
    ' Hide the unused rows in the Next Actions tab
        .Rows((int_LastRow + 1) & ":" & (int_LastRow + 1)).Hidden = False
        .Rows((int_LastRow + 2) & ":" & (Rows.count)).Hidden = True
              
    ' Wipe out the selection
        Application.GoTo .Range("A1"), False
              
    End With
    
End Sub
Sub o_51_Reset_To_Do()
Attribute o_51_Reset_To_Do.VB_ProcData.VB_Invoke_Func = " \n14"

' Purpose: To reapply my filters, sort, and otherwise reset my To Do excel.
' Trigger: Called: o_72_Support_Reset_Splitter
' Updated: 3/25/2023
' Reviewd: 5/11/2023

' Change Log:
'       6/6/2020:   Updated the sort to sort on Area and Project
'       7/25/2020:  Updated the sort to sort on Area and Project for Tasks
'       12/15/2020: Swithced to "ShowAllData" for the AutoFilter reset
'       12/18/2020: Replaced the Environ code w/ conditional compiler constant to determine if this is my Personal To Do.
'       12/28/2020: Moved the code out to sort ws_Projects and ws_Tasks
'       12/28/2020: Simplified the code for ws_Recurring
'       11/1/2021:  Updated the code related to ws_Recurring to be explicit with the ranges
'       7/31/2022:  Added the code to reset the heading visibilty (fx_Reset_Heading_Visibility_in_ToDo)
'       3/25/2023:  Removed the Someday Maybe related code
   
' ***********************************************************************************************************************************

Call Macros.o_02_Assign_Private_Variables
    
' -----------------
' Declare Variables
' -----------------

    Dim CurrLocation As Range
    Set CurrLocation = Selection

' -----------------
' Reset ws_Recurring
' -----------------
    
Call Macros.o_07_Assign_Private_Variables_ws_Recurring
    
    With ws_Recurring

        If .AutoFilterMode = True Then .AutoFilter.ShowAllData

        .Range("A1").AutoFilter Field:=col_Priority, Criteria1:=Array("Medium", "High", "Reminder"), Operator:=xlFilterValues
        .Range("A1").AutoFilter Field:=col_Status, Criteria1:=Array("Continuous", "Current", "Past Due"), Operator:=xlFilterValues

    End With
                                                    
' -----------------------------------------------------
' Reset ws_Projects, ws_Tasks, ws_Waiting, and ws_Questions
' -----------------------------------------------------
                
    Call o_13_Filter_wsProjects
    Call o_22_Filter_wsTasks
    Call o_32_Filter_wsWaiting
    Call o_42_Filter_wsQuestions

    Call fx_Reset_Heading_Visibility_in_ToDo
    
    Application.EnableEvents = True ' Added 3/5/2023

Application.GoTo CurrLocation, False

End Sub
Sub o_52_Reset_To_Do_Weekly()
Attribute o_52_Reset_To_Do_Weekly.VB_ProcData.VB_Invoke_Func = " \n14"

' Purpose: To reset my To Do excel to prep for the coming week.
' Trigger: Ribbon Icon - GTD Macros > Reset > Weekly ToDo Reset
' Updated: 11/13/2022
' Reviewd: 5/11/2023

' Change Log:
'       2/24/2020:  Removed the Start / Continue and Stop / Change Section
'       2/24/2020:  Added a section for my notes from the Daily Review UserForm
'       5/10/2020:  Brought back the Start / Continue and Stop / Change sections
'       5/10/2020:  Disabled the Daily Goals - Performance section
'       4/30/2021:  Remove/Cleanup old code related to wsCurrent and ws_Daily
'       5/1/2021:   Deleted content related to Personal I no longer use, "This Week"
'       11/22/2021: Added the Prep ws_Tasks code
'       12/5/2021:  Added the autofilter for ws_Tasks based on Priority (High & Medium)
'       11/13/2022: Called the code to reset the ws_Tasks before doing the filtering

' ***********************************************************************************************************************************

Call myPrivateMacros.DisableForEfficiency
Call Macros.o_02_Assign_Private_Variables
Call Macros.o_04_Assign_Private_Variables_wsTasks

' -----------------
' Declare Variables
' -----------------
    
    ' Dim Worksheets
'    Dim wsCurrent As Worksheet => 5/11/23: Temporarly disabled, I no longer use the wsCurrent
'        Set wsCurrent = ThisWorkbook.Sheets("Current")
    
    ' Dim Integers
    Dim int_LastRow_wsDaily As Long
        int_LastRow_wsDaily = Application.Max( _
            ws_Daily.Cells(Rows.count, "B").End(xlUp).Row, _
            ws_Daily.Cells(Rows.count, "C").End(xlUp).Row, _
            ws_Daily.Cells(Rows.count, "D").End(xlUp).Row)
    
    If int_LastRow_wsDaily = ws_Daily.Cells(Rows.count, "B").End(xlUp).Row Then
        int_LastRow_wsDaily = int_LastRow_wsDaily + 7 ' Add a buffer if the day column is used
    End If
    
    Dim intWeekday As Long ' Change the number of days depending on if it's the Professional or Personal To Do
        If Weekday(Date) = 2 And Environ("UserName") <> "JRina" Then
            intWeekday = 1
        ElseIf Weekday(Date) = 1 And Environ("UserName") = "JRina" Then
            intWeekday = 1
        Else
            intWeekday = 8
        End If

    Dim i As Long
    
    Dim int_CurRow_wsProjects As Long
        int_CurRow_wsProjects = fx_Find_CurRow(ws:=ws_Projects, strTargetFieldName:="Project", strTarget:="")
        
    ' Dim Dictionaries
    Dim dict_ActiveProjects As New Scripting.Dictionary
        dict_ActiveProjects.CompareMode = TextCompare

' -----------------------------------
' Perform the manipulation on ws_Daily
' -----------------------------------
    
    ' Add the date row and Day headers in the Daily tab
        
        int_LastRow_wsDaily = int_LastRow_wsDaily + 2 ' Use this code to move the "active" row for the int_LastRow_wsDaily variable
        
        With ws_Daily
        
        .Cells(int_LastRow_wsDaily, "A").FormulaR1C1 = "'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
        'Add the header
            int_LastRow_wsDaily = int_LastRow_wsDaily + 1 ' Use this code to move the "active" row for the int_LastRow_wsDaily variable
        
            With .Cells(int_LastRow_wsDaily, "C")
                
                If Environ("UserName") = "JRina" Then
                    .Value = CStr(Date + intWeekday - Weekday(Date, vbSunday)) & " - " & CStr(Date + 8 + 7 - Weekday(Date, vbSaturday))
                Else
                    .Value = CStr(Date + intWeekday - Weekday(Date, vbMonday)) & " - " & CStr(Date + 8 - Weekday(Date, vbFriday))
                End If

                .Font.Bold = "True"
                .Font.Underline = xlUnderlineStyleSingle
                .HorizontalAlignment = xlCenter
            End With
            
        ' Add Improved / Learned section
        
            int_LastRow_wsDaily = int_LastRow_wsDaily + 2
                        
            .Cells(int_LastRow_wsDaily, "B").Value = "Improved / Learned"
                .Cells(int_LastRow_wsDaily, "B").Font.Underline = xlUnderlineStyleSingle
                .Cells(int_LastRow_wsDaily, "B").Font.Bold = True
            
            .Range("B" & int_LastRow_wsDaily + 1 & ":" & "B" & int_LastRow_wsDaily + 7).IndentLevel = 1
        
        ' Add Start / Continue section

            int_LastRow_wsDaily = int_LastRow_wsDaily + 8

            .Cells(int_LastRow_wsDaily, "B").Value = "Start / Continue"
                .Cells(int_LastRow_wsDaily, "B").Font.Underline = xlUnderlineStyleSingle
                .Cells(int_LastRow_wsDaily, "B").Font.Bold = True

            .Range("B" & int_LastRow_wsDaily + 1 & ":" & "B" & int_LastRow_wsDaily + 7).IndentLevel = 1

        ' Add Stop / Change section

            int_LastRow_wsDaily = int_LastRow_wsDaily + 8

            .Cells(int_LastRow_wsDaily, "B").Value = "Stop / Change"
                .Cells(int_LastRow_wsDaily, "B").Font.Underline = xlUnderlineStyleSingle
                .Cells(int_LastRow_wsDaily, "B").Font.Bold = True

            .Range("B" & int_LastRow_wsDaily + 1 & ":" & "B" & int_LastRow_wsDaily + 7).IndentLevel = 1
        
        ' Add Positive Experiences section
        
            int_LastRow_wsDaily = int_LastRow_wsDaily + 8
                        
            .Cells(int_LastRow_wsDaily, "B").Value = "Positive Experiences"
                .Cells(int_LastRow_wsDaily, "B").Font.Underline = xlUnderlineStyleSingle
                .Cells(int_LastRow_wsDaily, "B").Font.Bold = True
            
            .Range("B" & int_LastRow_wsDaily + 1 & ":" & "B" & int_LastRow_wsDaily + 7).IndentLevel = 1
        
        End With
        
    ' Go to top of the record for the new week
        Application.GoTo ws_Daily.Cells(Rows.count, "A").End(xlUp), True

' -----------------------------
' Apply the Active Project flag
' -----------------------------

    'Fill the Active Projects Dictionary
    
        For i = 2 To int_CurRow_wsProjects
            If ws_Projects.Cells(i, "G") = "Active" Then
                dict_ActiveProjects.Add key:=ws_Projects.Cells(i, "D").Value2, Item:=ws_Projects.Cells(i, "D").Value2
            End If
        Next i

    ' If the project matches, mark Task as Active, otherwise not
    
    With ws_Tasks
        For i = 2 To int_CurRow
            
            If dict_ActiveProjects.Exists(CStr(.Cells(i, col_Project).Value2)) Then
                .Cells(i, col_ActiveProject_wsTasks) = "X"
            Else
                .Cells(i, col_ActiveProject_wsTasks) = ""
            End If
            
        Next i
    End With
    
    Call Macros.o_22_Filter_wsTasks
    
' ------------
' Prep ws_Tasks
' ------------

    With ws_Tasks
        Union(.Columns(col_ActiveTask_wsTasks), .Columns(col_ActiveComponent_wsTasks), .Columns(col_ActiveProject_wsTasks)).EntireColumn.Hidden = False
        .Range("A1").AutoFilter Field:=col_Priority, Criteria1:=Array("High", "Medium", "="), Operator:=xlFilterValues
    End With
    
' ----------------------
' Cleanup the Current ws
' ----------------------

'    With wsCurrent
'        .Range("B6:B8").ClearContents
'    End With
            
    ' Reset the view back to the Tasks ws
        
    Application.GoTo Sheets("Tasks").Range("A1"), True
                    
Call myPrivateMacros.DisableForEfficiencyOff

Application.EnableEvents = True ' Added 3/5/2023

End Sub
Sub o_53_Complete_Daily_Review()

' Purpose: To force me to leave work thinking about something positive each day.
' Trigger: Event: Workbook_BeforeClose
' Updated: 5/3/2022
' Reviewd: 5/11/2023

'Change Log:
'       9/1/2020:   Converted to call the UserForm
'       9/16/2020:  Switched to vbModeless
'       12/31/2020: Added the code to abort if I already completed the day's review
'       5/3/2022:   Renamed from 'u_Daily_Review' to 'o_53_Complete_Daily_Review'

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim strFileLoc As String
        strFileLoc = "C:\U Drive\Support\Daily Review.xlsx"
        
    ' Dim Workbooks / Worksheets
    Dim wbDailyReview As Workbook
    Set wbDailyReview = Workbooks.Open(Filename:=strFileLoc, UpdateLinks:=False, ReadOnly:=False)
    
    Dim wsData As Worksheet
    Set wsData = wbDailyReview.Sheets(1)

    ' Dim Integers
    Dim int_CurRow As Long
        int_CurRow = wsData.Cells(Rows.count, "A").End(xlUp).Row + 1

    ' Dim Strings
    Dim strToday As String
        strToday = Date & " (" & Format(Date, "DDDD") & ")"
    
    Dim strLastDay As String
        strLastDay = wsData.Cells(int_CurRow - 1, 1).Value2

'--------------------------------------------
' Abort if the day has already been completed
'--------------------------------------------

    If strToday = strLastDay Then
        wbDailyReview.Close SaveChanges:=False
        Exit Sub
    Else
        uf_Daily_Review.Show vbModeless
    End If
        
End Sub
Sub o_55_Weekly_GTD_Review()

' Purpose: To apply filters to only look at what I accomplished in the last week, to aid in my Weekly GTD Review.
' Trigger: Ribbon Icon - GTD Macros > Review
' Updated: 1/15/2022
' Reviewd: 5/11/2023

' Change Log:
'       12/9/2019:  Added in the msgbox option to select the week to filter on
'       6/12/2021:  Added a Sort to make it easier to copy the data into my GTD Review word doc.
'       6/12/2021:  Added the 'With' statements to  make it easier to read / update
'       7/19/2021:  Updated to determine the week to review based on the current week day
'       11/13/2021: Updated to use the assigned variables for ws_Projects, ws_Tasks and ws_Waiting
'       11/13/2021: Updated the code to remove the Conditional Compilation Argument for ws_Recurring
'       1/15/2022:  Updated the order of when the sort is applied, to show all data, then sort, then filter
   
' ***********************************************************************************************************************************

Call myPrivateMacros.DisableForEfficiency
Call Macros.o_02_Assign_Private_Variables
        
' -----------------
' Declare Variables
' -----------------

    Dim int_Week_Filter As Long
        If Format(Date, "dddd") = "Friday" Or Format(Date, "dddd") = "Saturday" Then
            int_Week_Filter = xlFilterThisWeek
        Else
            int_Week_Filter = xlFilterLastWeek
        End If
        
' -----------------------------------------------------------------------
' This portion of the macro filters each tab to just last week's activity
' -----------------------------------------------------------------------
                
    ' Filter Projects
    
    Call o_03_Assign_Private_Variables_wsProjects
    
    With ws_Projects
    
        .AutoFilter.ShowAllData
        
        .Range("A1").Sort _
            Key1:=.Cells(1, col_Area), Order1:=xlAscending, _
            Key2:=.Cells(1, col_Project), Order2:=xlAscending, _
            Header:=xlYes
        
        .Range("A1").AutoFilter Field:=col_Updated_wsProjects, Criteria1:=int_Week_Filter, Operator:=xlFilterDynamic
    
    End With
            
    ' Filter Tasks
    
    Call o_04_Assign_Private_Variables_wsTasks
    
    With ws_Tasks

        .AutoFilter.ShowAllData
    
        .Range("A1").Sort _
            Key1:=.Cells(1, col_Area), Order1:=xlAscending, _
            Key2:=.Cells(1, col_Project), Order2:=xlAscending, _
            Key3:=.Cells(1, col_Task), Order3:=xlAscending, _
            Header:=xlYes
    
        .Range("A1").AutoFilter Field:=col_Completed, Criteria1:=int_Week_Filter, Operator:=xlFilterDynamic
    
    End With
        
    ' Filter Waiting
    
    Call o_05_Assign_Private_Variables_wsWaiting
    
    With ws_Waiting
        
        .AutoFilter.ShowAllData
        
        .Range("A1").Sort _
            Key1:=.Cells(1, col_Project), Order1:=xlAscending, _
            Key2:=.Cells(1, col_WaitingFor_wsWaiting), Order2:=xlAscending, _
            Header:=xlYes
    
        .Range("A1").AutoFilter Field:=col_Completed, Criteria1:=int_Week_Filter, Operator:=xlFilterDynamic
        
    End With
        
    ' Filter Recurring
        
    Call o_07_Assign_Private_Variables_ws_Recurring
    
    With ws_Recurring
        
        .AutoFilter.ShowAllData
            ws_Recurring.Range("A1").AutoFilter Field:=col_LastCompleted_ws_Recurring, Criteria1:=int_Week_Filter, Operator:=xlFilterDynamic

    End With

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_56_Create_Daily_To_Do_txt_For_Upcoming_Week()

' Purpose: To create the "Daily To Do" txt files for the upcoming week.
' Trigger: cmd_DailyToDo_Click
' Updated: 7/27/2023
' Reviewd: 5/20/2023

' Change Log:
'       4/19/2021:  Intial Creation, based on o_93_Open_Daily_To_Do_txt.
'       6/7/2021:   Cleaned up the code to remove the IF statements related to picking a date that isn't tomorrow.
'       6/7/2021:   Updated the code to create the file for Monday, if it's a Friday
'       6/7/2021:   Updated the code to breakout dtNextDay to be easier to maintain
'       7/26/2022:  Updated the hyperlink locations for the new Planning Template folder
'       11/25/2022: Moved this code from my Daily Review uf to Macros, to be triggered as part of my Weekly Review
'                   Broke out the creation of the file code from the opening
'                   Created the intDayCount variable to loop through all five days
'       12/3/2022:  Handled if I run the process on the weekend
'       12/4/2022:  Removed the "Personal" code, as this is now only used for Professional use
'       7/27/2023:  Added Thursday as a day to trigger incase of a long weekend

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    ' Dim Objects
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Dim Folder Location
    Dim strFolderLoc As String
        strFolderLoc = "C:\U Drive\Support\Daily To Do\"

    ' Dim Integers
    Dim intDayCount As Long
        
' --------------------------
' Declare Variables For Loop
' --------------------------
           
For intDayCount = 1 To 5
           
    ' Dim the Next Business Day
    Dim dtNextDay As Date
        If Format(Date, "dddd") = "Friday" Then
            dtNextDay = Date + intDayCount + 2
        ElseIf Format(Date, "dddd") = "Saturday" Then
            dtNextDay = Date + intDayCount + 1
        ElseIf Format(Date, "dddd") = "Sunday" Then
            dtNextDay = Date + intDayCount
        ElseIf Format(Date, "dddd") = "Thursday" Then
            dtNextDay = Date + intDayCount + 3
        Else
            dtNextDay = Date + intDayCount - 1
        End If
        
    ' Dim Template Location
    Dim strFileLoc_DailyTemplate As String
        strFileLoc_DailyTemplate = "C:\U Drive\Support\Planning Templates\XXXX.XX.XX - Daily To Do - Template (" _
            & Format(dtNextDay, "dddd") & ").properties"
                
    If Format(dtNextDay, "dddd") = "Saturday" Then Exit For ' Abort if the next day is going to be a Saturday
    
    ' Dim Dates for Next Day
    
    Dim strNextDay_Date As String
        strNextDay_Date = Format(dtNextDay, "yyyy.mm.dd")
        
    Dim strNextDay_Day As String
        strNextDay_Day = Format(dtNextDay, "DDDD")
        
    Dim strFileLoc As String
        strFileLoc = strFolderLoc & strNextDay_Date & " - " & strNextDay_Day & " - Daily To Do.properties"
        
' --------------------------------------------
' Create the Daily To Do for the upcoming week
' --------------------------------------------
                    
    If objFSO.FileExists(strFileLoc) <> True Then
        objFSO.CopyFile strFileLoc_DailyTemplate, strFileLoc
    End If
        
' ---------------------------------
' Open the next business day's file
' ---------------------------------

    Call Shell("explorer.exe" & " " & strFileLoc)

Next intDayCount

End Sub
Sub o_57_Open_uf_Perspectives()

' Purpose: To open the Perspectives UserForm.
' Trigger: Ribbon
' Updated: 7/19/2021
' Reviewd: 5/20/2023

' Change Log:
'       7/19/2021: Initial Creation

' ***********************************************************************************************************************************

    uf_Perspectives.Show vbModeless

End Sub
Sub o_58_Adjust_For_Laptop_Or_Monitor()

' Purpose: To adjust the zoom of each worksheet in my To Do, based on opening the file in my Laptop or Desktop.
' Trigger: Called on Workbook open
' Updated: 4/3/2023
' Reviewd: 5/20/2023

' Change Log:
'       12/14/2021: Initial Creation
'       12/27/2021: Added the option for 1080p vs 4k
'       4/3/2023:   Updated to reflect my new 4k monitor
   
' ***********************************************************************************************************************************

Call myPrivateMacros.DisableForEfficiency

' -----------------
' Declare Variables
' -----------------
    
    ' Declare Integers
    Dim intScreenHeight As Long
        intScreenHeight = fx_Get_Screen_Resolution(1) ' height in points

    Dim strLaptoporMonitor As String

    If intScreenHeight = 1080 Then
        strLaptoporMonitor = "Z Book - 1080p Laptop"
    ElseIf intScreenHeight = 1200 Then
        strLaptoporMonitor = "Dell U2415H - 16:10 Monitor"
    ElseIf intScreenHeight > 1200 Then
        strLaptoporMonitor = "Dell U3223QE - 4k Monitor"
    End If

    ' Declare Worksheets
    Dim ws As Worksheet

' -----------------------------------------
' Adjust the zoom of each of the worksheets
' -----------------------------------------

    If strLaptoporMonitor = "Z Book - 1080p Laptop" Then ' Apply the zoom for my Win10 Zbook
    
        For Each ws In ThisWorkbook.Sheets
            ws.Activate
            
            Select Case ws.Name
                Case "Current"
                    ActiveWindow.Zoom = 190
                Case "Projects"
                    ActiveWindow.Zoom = 100
                Case "Tasks"
                    ActiveWindow.Zoom = 90
                Case "Waiting"
                    ActiveWindow.Zoom = 90
                Case "Questions"
                    ActiveWindow.Zoom = 90
                Case "Recurring"
                    ActiveWindow.Zoom = 100
                Case "Temp"
                    ActiveWindow.Zoom = 110
                Case "Daily"
                    ActiveWindow.Zoom = 110
            End Select
        Next ws
    
    ElseIf strLaptoporMonitor = "Dell U3223QE - 4k Monitor" Then ' Apply the zoom for my new Dell U3223QE 4k Monitor
    
        For Each ws In ThisWorkbook.Sheets
            ws.Activate
            
            Select Case ws.Name
                Case "Current"
                    ActiveWindow.Zoom = 225
                Case "Projects"
                    ActiveWindow.Zoom = 110
                Case "Tasks"
                    ActiveWindow.Zoom = 100
                Case "Waiting"
                    ActiveWindow.Zoom = 100
                Case "Questions"
                    ActiveWindow.Zoom = 100
                Case "Recurring"
                    ActiveWindow.Zoom = 105
                Case "Temp"
                    ActiveWindow.Zoom = 105
                Case "Daily"
                    ActiveWindow.Zoom = 110
            End Select
        Next ws
    
'    ElseIf strLaptoporMonitor = "Dell U2415H - 16:10 Monitor" Then ' Apply the zoom for my Dell monitor
'
'        For Each ws In ThisWorkbook.Sheets
'            ws.Activate
'
'            Select Case ws.Name
'                Case "Current"
'                    ActiveWindow.Zoom = 246
'                Case "Projects"
'                    ActiveWindow.Zoom = 105
'                Case "Tasks"
'                    ActiveWindow.Zoom = 100
'                Case "Waiting"
'                    ActiveWindow.Zoom = 100
'                Case "Questions"
'                    ActiveWindow.Zoom = 100
'                Case "Recurring"
'                    ActiveWindow.Zoom = 115
'                Case "Temp"
'                    ActiveWindow.Zoom = 115
'                Case "Daily"
'                    ActiveWindow.Zoom = 115
'            End Select
'        Next ws
        
    End If

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_61_Create_Meeting_Notes()

' Purpose: To create a template in To Do > Daily for my meeting notes.
' Trigger: Keyboard Shortcut - Ctrl + Shift + Z (Called by o_71_Dynamic_Macro_Splitter)
' Updated: 12/4/2023
' Reviewd: 5/20/2023

' Change Log:
'       7/16/2020:  Initial Creation
'       9/18/2020:  Added some additional details around takeaways
'       6/3/2021:   Updated int_CurRow to select an existing row if it looks like I started to type meeting notes.
'       6/3/2021:   Added the inputbox to prompt the user for the meeting name
'       6/4/2021:   Added the Reference Material section
'       6/8/2021:   Added the section for Pending Questions
'       6/21/2021:  Added a line for Sterling Attendees
'       6/23/2021:  Converted from numbers Next Actions to bullets
'       6/24/2021:  Commented out the Meeting Type
'       6/24/2021:  Replaced the TBD for Takeaways with "¤"
'       7/23/2021:  Reduced the # of bullets for each section to 1, now that I have o_62_Currate_Meeting_Notes
'       8/30/2021:  Removed the line for Sterling Attendees
'       10/25/2021: Updated how the bold works for the header info
'       5/23/2022:  Updated the symbols, and added the Waiting section
'       10/23/2023: Added the code to determine the Entered date automatically for new Next Actions
'       12/4/2023:  Updated to move the meeting date / time to it's own row instead of the title

' ***********************************************************************************************************************************

Call Macros.o_02_Assign_Private_Variables

' -----------------
' Declare Variables
' -----------------
    
'   Dim int_CurRow As Long
        int_CurRow = ws_Daily.Cells(Rows.count, "D").End(xlUp).Row + 2

    Dim strMeetingName As String

    ' If you are on a cell with text, but the cell below is blank, use that
    If ActiveCell.Value2 <> "" And ActiveCell.Offset(1, 0).Value2 = "" Then
        int_CurRow = ActiveCell.Row
        strMeetingName = ActiveCell.Value2
    End If
    
    ' If the active cell and the next cell isn't blank, abort
    If ActiveCell.Value2 <> "" And ActiveCell.Offset(1, 0).Value2 <> "" Then Exit Sub
    
    ' Assign Strings
        If strMeetingName = "" Then
            strMeetingName = InputBox(Prompt:="What is the name of the  meeting?", Title:="", Default:=strMeetingName)
        End If

' --------------------------
' Create the Meeting Details
' --------------------------

If strMeetingName = "" Then Exit Sub

With ws_Daily
    
    'Create the header info
    .Range("D" & int_CurRow).Value2 = strMeetingName
        .Range("D" & int_CurRow).Font.Bold = True
        int_CurRow = int_CurRow + 1
    
    .Range("D" & int_CurRow).Value2 = "Meeting Date: " & Date & " at " & Format(Time, "h:mm AM/PM") 'Meeting Date
        .Range("D" & int_CurRow).Characters(1, Len("Meeting Date")).Font.Bold = True
        int_CurRow = int_CurRow + 1
    
    .Range("D" & int_CurRow).Value2 = "Attendees:"
        .Range("D" & int_CurRow).Characters(1, Len("Attendees")).Font.Bold = True 'Attendees
        int_CurRow = int_CurRow + 1
    
    .Range("D" & int_CurRow).Value2 = "Meeting Purpose: "
        .Range("D" & int_CurRow).Characters(1, Len("Meeting Purpose")).Font.Bold = True 'Meeting Purpose
        int_CurRow = int_CurRow + 1

    .Range("D" & int_CurRow).Value2 = "Project: TBD"
        .Range("D" & int_CurRow).Characters(1, Len("Project")).Font.Bold = True 'Project
        int_CurRow = int_CurRow + 1
    
    'Create line break
    .Range("D" & int_CurRow).Value2 = "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "

        int_CurRow = int_CurRow + 1
    
' ------------------------------
' Callout the Next Actions: Task
' ------------------------------
    
    .Range("D" & int_CurRow).Value2 = "Next Actions: Task"
        .Range("D" & int_CurRow).Font.Underline = True
        int_CurRow = int_CurRow + 1
    
    'Create line break
    .Range("D" & int_CurRow).Value2 = "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        .Range("D" & int_CurRow).Font.Underline = False
        
        int_CurRow = int_CurRow + 1

' ---------------------------------
' Callout the Next Actions: Waiting
' ---------------------------------
    
    .Range("D" & int_CurRow).Value2 = "Next Actions: Waiting"
        .Range("D" & int_CurRow).Font.Underline = True
        int_CurRow = int_CurRow + 1
        
    'Create line break
    .Range("D" & int_CurRow).Value2 = "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        .Range("D" & int_CurRow).Font.Underline = False
        
        int_CurRow = int_CurRow + 1
        
' -----------------------------
' Callout the Pending Questions
' -----------------------------

    .Range("D" & int_CurRow).Value2 = "Pending Questions"
        .Range("D" & int_CurRow).Font.Underline = True
        int_CurRow = int_CurRow + 1

    'Create line break
    .Range("D" & int_CurRow).Value2 = "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        .Range("D" & int_CurRow).Font.Underline = False
        
        int_CurRow = int_CurRow + 1
        
' -------------------------
' Callout the Key Takeaways
' -------------------------

    .Range("D" & int_CurRow).Value2 = "Key Takeaways"
        .Range("D" & int_CurRow).Font.Underline = True
        int_CurRow = int_CurRow + 1
        
    'Create line break
    .Range("D" & int_CurRow).Value2 = "- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - "
        .Range("D" & int_CurRow).Font.Underline = False
        
        int_CurRow = int_CurRow + 1

' --------------------------------
' Create the meeting Notes section
' --------------------------------
    
    'Callout the Notes section
    .Range("D" & int_CurRow).Value2 = "Notes"
        .Range("D" & int_CurRow).Font.Underline = True
        int_CurRow = int_CurRow + 1

End With

End Sub
Sub o_62_Currate_Meeting_Notes()

' Purpose: To curate my meeting notes and add notes to the Key Takeaways, Next Actions, and Pending Questions sections.
' Trigger: Keyboard Shortcut - Ctrl + Shift + Z (Called by o_71_Dynamic_Macro_Splitter)
' Updated: 6/20/2022
' Reviewd: 5/20/2023

' Legend:
'   Notes:          "  >"
'   Next Actions:   "[•]"
'   Waiting:        "[o]"
'   Key Takeaways:  "[¤]"
'   Questions:      "[¿]"

' Change Log:
'       7/23/2021:  Initial Creation
'       8/10/2021:  Added the code around the footer row to nest comments, and the rngNotesSection variable
'       8/10/2021:  Added code to nest sub-comments (with a dash -)
'       8/10/2021:  Convereted to using a Do While to handle the counter (thanks Axcel)
'       8/12/2021:  Added the code for 'i = intHeaderRow' to reset i
'       9/2/2021:   Added the code to select the full active range once everything is run
'       10/1/2021:  Updated to apply the Strikethrough formating
'       5/23/2022:  Updated the handle the new approach using o_63_Update_Meeting_Notes_Type
'                   Updated to normalize / simplify each of the Next Action types
'                   Broke the insert new row loop into it's own section
'                   Removed the 'Remove any empty cells' section, no longer applicable
'                   Accounted for multiple notes under one Next Action
'       5/24/2022:  Removed the intNewRowCounter, replaced by updating the int_LastRow variable
'       6/20/2022:  Updated to address an issue where the row for notes was being skipped, and incorrectly aggregating

' ***********************************************************************************************************************************

' Abort if I didn't use my Create Meeting Notes macro
If Selection.Find("Next Actions") Is Nothing Then
    Exit Sub
End If

Call myPrivateMacros.DisableForEfficiency

Call Macros.o_02_Assign_Private_Variables

' -----------------
' Declare Variables
' -----------------
    
    ' Declare Ranges
    
    Dim rngFullNotes As Range
    Set rngFullNotes = Selection
    
    Dim rngNotesSection As Range
    
    ' Declare Integers
    
    Dim intFirstRow As Long
        intFirstRow = rngFullNotes.Row
   
    Dim int_LastRow As Long
        int_LastRow = intFirstRow + rngFullNotes.count - 1
    
    Dim i As Long
    
    ' Declare "Ranges"
    
    Dim intNotesHeaderRow As Long
        intNotesHeaderRow = rngFullNotes.Find("Notes").Row
        
    Dim intSectionHeaderRow As Long
    
    Dim intSectionFooterRow As Long
    
    ' Declare Strings
    
    Dim strRowText As String
    
    Dim strSymbol As String

' ---------------------------------------
' Assign Variables for Next Actions: Task
' ---------------------------------------

Task:

   intSectionHeaderRow = rngFullNotes.Find("Next Actions: Task").Row
    
    strSymbol = "[•] "
    
    GoTo LoopEachSection

' ------------------------------------------
' Assign Variables for Next Actions: Waiting
' ------------------------------------------

Waiting:

   intSectionHeaderRow = rngFullNotes.Find("Next Actions: Waiting").Row
        
    strSymbol = "[o] "
    
    GoTo LoopEachSection

' --------------------------------------
' Assign Variables for Pending Questions
' --------------------------------------

Questions:

    intSectionHeaderRow = rngFullNotes.Find("Pending Questions").Row

    strSymbol = "[¿] "
    
    GoTo LoopEachSection

' ----------------------------------
' Assign Variables for Key Takeaways
' ----------------------------------

KeyTakeaways:

    intSectionHeaderRow = rngFullNotes.Find("Key Takeaways").Row
        
    strSymbol = "[¤] "
    
    GoTo LoopEachSection
    
' -------------------------
' Loop through each section
' -------------------------
    
LoopEachSection:
    
With ws_Daily
    
    ' Assign the Variables
    
    i = intNotesHeaderRow + 1
    
    intSectionFooterRow = intSectionHeaderRow + 1
    
    ' Loop through the values from Notes to the bottom of the data, plus account for inserted rows
    
    Do While i >= intNotesHeaderRow And i <= int_LastRow
        
        strRowText = .Range("D" & i).Value2
        
        ' Handle each Next Action
        If Left(strRowText, 4) = strSymbol Then
            .Rows(intSectionFooterRow).Insert
                i = i + 1: int_LastRow = int_LastRow + 1: intSectionFooterRow = intSectionFooterRow + 1    'Need to adjust for the added row
                .Range("D" & intSectionFooterRow - 1).Value2 = strRowText
                .Range("D" & intSectionFooterRow - 1).Font.Underline = False
                .Range("D" & i).Font.Strikethrough = True
                
                i = i + 1
        
        ' Handle each of the contiguous notes for the Next Action
        ElseIf (Left(strRowText, 3) = "  >" Or Left(strRowText, 3) = " - ") And _
            Left(.Range("D" & i - 1).Value, 4) = strSymbol And _
            .Range("D" & i).Font.Strikethrough = False Then
                                        
            ' Loop through the values to capture multiple notes
            Do Until Left(.Range("D" & i).Value, 3) <> "  >" And Left(.Range("D" & i).Value, 3) <> " - "
            
            strRowText = .Range("D" & i).Value2

            .Rows(intSectionFooterRow).Insert
                i = i + 1: int_LastRow = int_LastRow + 1: intSectionFooterRow = intSectionFooterRow + 1    'Need to adjust for the added row
                .Range("D" & intSectionFooterRow - 1).Value2 = strRowText
                .Range("D" & intSectionFooterRow - 1).Font.Underline = False
                .Range("D" & i).Font.Strikethrough = True
                
                i = i + 1 ' Go to next value for 'Do Until' loop
            Loop

        Else
            i = i + 1
        End If
        
    Loop
        
End With

    If strSymbol = "[•] " Then
        GoTo Waiting
    ElseIf strSymbol = "[o] " Then
        GoTo Questions
    ElseIf strSymbol = "[¿] " Then
        GoTo KeyTakeaways
    ElseIf strSymbol = "[¤] " Then
        ' Do nothing
    End If
    
' ----------------------
' Select all of the text
' ----------------------

Set rngFullNotes = rngFullNotes.CurrentRegion
    rngFullNotes.Select

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_63_Update_Meeting_Notes_Type(strNextActionType As String)

' Purpose: To rotate through the different types of meeting notes.
' Trigger: Triggered by various Macros, via Keyboard Shortcut (ex. Ctrl + Shift + T)
' Updated: 9/20/2023
' Reviewd: 5/20/2023

' Legend:
'   Notes:          "  >"
'   Next Actions:   "[•]"
'   Waiting:        "[o]"
'   Key Takeaways:  "[¤]"
'   Questions:      "[¿]"

' Change Log:
'       5/22/2022:  Initial Creation
'       5/23/2022:  Overhauled to break out each section, and loop through the full range
'       5/24/2022:  Added the detail to include " - " as a Note
'       8/5/2022:   Added Trim to remove blanks when copying in from Notepad++
'                   Added the 'strUpdatedRowText' variable to be able to track the changes
'                   Moved all of the code to the update section to normalize, simplifiying each section
'       9/1/2022:   Added Trim to fix the duplicate space issue
'       9/20/2023:  Updated so that if I trigger the macro after adding a line it applies to the line above, instead of the blank line
'                   Updated to use a Case statement for applying the formatting instead of multipe If statements
'                   Updated so that the Detail (>) is just another type of Next Action, simplifiying the code

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------
    
    Dim rngSelection As Range
    Set rngSelection = Selection
    
    Dim strOriginalRowText As String
    
    Dim strUpdatedRowText As String
        
    Dim cell As Range
    
    Dim intCellCount As Long

' ------------------------------------------------------------------------
' Update the approach if it's a blank cell or multiple cells were selected
' ------------------------------------------------------------------------

For Each cell In rngSelection
    
    intCellCount = intCellCount + 1
    strOriginalRowText = Trim(cell.Value2)
    
    If strOriginalRowText = "" And intCellCount = 1 Then ' If applying to a blank row apply to the row above
        Set cell = cell.Offset(-1, 0)
        strOriginalRowText = Trim(cell.Value2)
    End If
    
    If intCellCount > 1 Then strNextActionType = "Detail" ' For everything beyond the first cell, flag as detail

' ------------------------------------------
' Update the strOriginalRowText to normalize
' ------------------------------------------

    If Left(strOriginalRowText, 3) = "[•]" Or _
       Left(strOriginalRowText, 3) = "[o]" Or _
       Left(strOriginalRowText, 3) = "[¿]" Or _
       Left(strOriginalRowText, 3) = "[¤]" Then
        strUpdatedRowText = Trim(Right(strOriginalRowText, Len(strOriginalRowText) - 3))
    
    ElseIf Left(strOriginalRowText, 2) = "> " Or _
       Left(strOriginalRowText, 2) = "- " Or _
       Left(strOriginalRowText, 2) = "• " Or _
       Left(strOriginalRowText, 2) = "o " Or _
       Left(strOriginalRowText, 2) = "¤ " Then
        strUpdatedRowText = Trim(Right(strOriginalRowText, Len(strOriginalRowText) - 2))
        
    Else
        strUpdatedRowText = strOriginalRowText
    End If
    
' ------------------------------------
' Update based on the Next Action type
' ------------------------------------
    
    Select Case strNextActionType
        Case "Task"
            strUpdatedRowText = "[•] " & strUpdatedRowText
        Case "Waiting"
            strUpdatedRowText = "[o] " & strUpdatedRowText
        Case "Question"
            strUpdatedRowText = "[¿] " & strUpdatedRowText
        Case "Key Takeaway"
            strUpdatedRowText = "[¤] " & strUpdatedRowText
        Case "Detail"
            strUpdatedRowText = "  > " & strUpdatedRowText
    End Select

' ------------------
' Update the Value/s
' ------------------
    
    cell.Value2 = strUpdatedRowText

Next cell ' Continue looping through the selection

End Sub
Sub o_64_Create_Key_Takeaway_in_wsDaily()
Attribute o_64_Create_Key_Takeaway_in_wsDaily.VB_ProcData.VB_Invoke_Func = "K\n14"

' Purpose: To allow me to quickly apply the 'Key Takeaway' formatting in ws_Daily
' Trigger: Keyboard Shortcut - Ctrl + Shift + K
' Updated: 5/23/2022
' Reviewd: 5/20/2023

' Change Log:
'       5/23/2022:  Initial Creation was sometime in Q1 2018

' ***********************************************************************************************************************************
    
' ----------------------------------------------------
' If ws_Daily apply Next Action formatting to selection
' ----------------------------------------------------

    If ActiveSheet.Name = "Daily" Then
        Call Macros.o_63_Update_Meeting_Notes_Type("Key Takeaway")
        Exit Sub
    End If
    
End Sub
Sub o_71_Dynamic_Macro_Splitter()
Attribute o_71_Dynamic_Macro_Splitter.VB_ProcData.VB_Invoke_Func = "Z\n14"

' Purpose: To have a macro that can be ran dependant on the situtation, but with the same keyboard shortcut for multiples.
' Trigger: Keyboard Shortcut - Ctrl + Shift + Z
' Updated: 12/10/2021
' Reviewd: 5/20/2023

' Change Log:
'       1/9/2020:   Added the code to prevent overwtitting existing data
'       7/15/2020:  Added the code to create a new meeting on Daily
'       7/23/2021:  Added the code for o_62_Currate_Meeting_Notes
'       12/10/2021: Commented out the code related to u_Copy_Project_ID_to_Clipboard

' ***********************************************************************************************************************************

' ------------------------
' Prevent overwritten data
' ------------------------
    
Call myPrivateMacros.DisableForEfficiency
    
    If ActiveSheet.Name = "Projects" Then
    
    ElseIf ActiveSheet.Name = "Daily" Then
        If Selection.count <= 1 Then
            Call o_61_Create_Meeting_Notes
        Else
            Call o_62_Currate_Meeting_Notes
        End If
    
    ElseIf ActiveCell.Value <> "" Then
    
        Dim intResponse As Long
            intResponse = MsgBox(Prompt:="Do you want to overwrite the existing line?", Buttons:=vbYesNo + vbQuestion, Title:="Overwrite data?")
                  
        If intResponse = vbNo Then Exit Sub
        
        Call myUtilityMacros_Ribbon.u_Create_Dynamic_Reference_Number
        Call myUtilityMacros_Keyboard.u_Steal_Formatting_Row_Above
    
    Else
    
        Call myUtilityMacros_Ribbon.u_Create_Dynamic_Reference_Number
        Call myUtilityMacros_Keyboard.u_Steal_Formatting_Row_Above
    
    End If
    
Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_72_Support_Reset_Splitter()
Attribute o_72_Support_Reset_Splitter.VB_ProcData.VB_Invoke_Func = "R\n14"
    
' Purpose: To have one process to reset my To Dos, P.XXX Support, DA-XXX Support, or Staff Support
' Trigger: Keyboard Shortcut - Ctrl + Shift + R
' Updated: 7/31/2022
' Reviewd: 5/20/2023

' Change Log:
'       9/18/2019:  Initial Creation
'       6/12/2021:  Removed the code related to Executive Leadership
'       6/12/2021:  Combined the code for P. and DA. project support
'       11/1/2021:  Disabled the code related to my P.XXX / DA.XXX Support workbooks
'       7/31/2022:  Updated to "Show All Data" if reseting a non-ToDo

' ***********************************************************************************************************************************

Call myPrivateMacros.DisableForEfficiency

' -----------------
' Declare Variables
' -----------------
    
    Dim strWBName As String
        strWBName = ActiveWorkbook.Name
    
' ---------------
' Reset the files
' ---------------

On Error GoTo ErrorHandler
    
    If strWBName = "To Do.xlsm" Then
        Call Macros.o_51_Reset_To_Do
    Else
        If ActiveSheet.AutoFilterMode = True Then
            ActiveSheet.AutoFilter.ShowAllData
        Else
            ActiveSheet.Range("1:1").AutoFilter
        End If
    End If
     
Call myPrivateMacros.DisableForEfficiencyOff
     
Exit Sub

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'-------------------------------------------------------------------------------------------
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
     
ErrorHandler:
 
Call myPrivateMacros.DisableForEfficiencyOff

MsgBox "The Support Reset was unable to run for some reason."
     
End Sub
Sub o_81_Access_RefX_Text_Files()
Attribute o_81_Access_RefX_Text_Files.VB_ProcData.VB_Invoke_Func = "M\n14"

' Purpose: To allow me to quickly append to my text files.
' Trigger: Keyboard Shortcut - Ctrl + Shift + M
' Updated: 9/16/2019
' Reviewd: 5/20/2023

' Change Log:

' ***********************************************************************************************************************************

    uf_RefX_and_wsDaily_Update.Show vbModeless

End Sub
Sub o_82_Open_Dynamic_Folder_Search()
Attribute o_82_Open_Dynamic_Folder_Search.VB_ProcData.VB_Invoke_Func = "F\n14"

' Purpose: To allow me to quickly open any folder on my "U Drive"
' Trigger: Keyboard Shortcut: Ctrl + Shift + F OR Called by uf_Project_Selector
' Updated: 2/19/2020
' Reviewd: 5/20/2023

' Change Log:

' ***********************************************************************************************************************************

    uf_Search_Folder.Show

End Sub
Sub o_83_Open_Dynamic_RefX_Search()
Attribute o_83_Open_Dynamic_RefX_Search.VB_ProcData.VB_Invoke_Func = "X\n14"

' Purpose: To allow me to quickly open any RefX file in my Reference folder
' Trigger: Keyboard Shortcut: Ctrl + Shift + X OR Called by uf_Project_Selector
' Updated: 2/19/2020
' Reviewd: 5/20/2023

' Change Log:

' ***********************************************************************************************************************************

    uf_Search_RefX.Show

End Sub
Sub o_84_Open_Dynamic_File_Search()
Attribute o_84_Open_Dynamic_File_Search.VB_ProcData.VB_Invoke_Func = "C\n14"

' Purpose: To allow me to quickly open any file dynamically in my Reference folder
' Trigger: Keyboard Shortcut: Ctrl + Shift + C OR Called by uf_Project_Selector
' Updated: 9/14/2023
' Reviewd: 9/14/2023

' Change Log:

' ***********************************************************************************************************************************

    uf_Search_File.Show

End Sub
Sub o_91_Go_To_Daily()
Attribute o_91_Go_To_Daily.VB_ProcData.VB_Invoke_Func = "Q\n14"

' Purpose: To jump between the last blank cell in my Daily tab, the top of Temp, and the bottom of Temp.
' Trigger: Keyboard Shortcut - Ctrl + Shift + Q
' Updated: 5/30/2020
' Reviewd: 5/20/2023

' Change Log:
'       9/16/2019:  Initial Creation (?)
'       5/30/2020:  Updated to replace Current w/ Projects as I use that more now

' ***********************************************************************************************************************************

Call Macros.o_02_Assign_Private_Variables

' -----------------
' Declare Variables
' -----------------

    Dim lastRow_Temp As Long
        lastRow_Temp = ws_Temp.Cells(Rows.count, "B").End(xlUp).Row + 1

    Dim int_CurRow_Daily As Long
        int_CurRow_Daily = Application.WorksheetFunction.Max( _
        ws_Daily.Cells(Rows.count, "D").End(xlUp).Row, _
        ws_Daily.Cells(Rows.count, "C").End(xlUp).Row) + 1
    
' -------------
' Run your code
' -------------

    If ActiveSheet.Name = "Projects" Then
        Application.GoTo ws_Temp.Cells(lastRow_Temp, "B"), False
    ElseIf ActiveSheet.Name = "Temp" Then
        Application.GoTo ws_Daily.Cells(int_CurRow_Daily, "D"), False
    Else
        Application.GoTo ws_Projects.Range("A2"), False
    End If
    
 End Sub
Sub o_92_Remember_For_Today()
   
' Purpose: To create the pop up to remind me of any tasks to accomplish before leaving for the day.
' Trigger: Event: Workbook_BeforeClose
' Updated: 2/28/2022
' Reviewd: 5/20/2023

' Change Log:
'       9/16/2019:  Initial Creation
'       6/7/2020:   Updated to move the Remember Me in my To Do, use the --- to be the end of the section
'       2/28/2022:  Updated to a nicer looking MsgBox

' ***********************************************************************************************************************************
   
With ThisWorkbook.Sheets("Current")
   
' -----------------
' Declare Variables
' -----------------
   
    Dim intRow_Remember As Long
        intRow_Remember = .Range("B:B").Find("Remember for today").Row
        
    Dim int_LastRow_Remember As Long
        int_LastRow_Remember = .Range("B:B").Find("-----").Row - 1
    
    Dim Message As String
    
    Dim i As Long
    
' ----------------------------------------------
' Create the messagebox, if there is any content
' ----------------------------------------------
    
    For i = intRow_Remember + 1 To int_LastRow_Remember

    If .Range("B" & i) <> "" Then
        Message = Message & .Range("B" & i) & Chr(10)
    End If

    Next i

End With

    If Len(Message) > 0 Then
       Call MsgBox(Title:="Remember for Today", Buttons:=vbInformation, Prompt:=Message)
    End If

End Sub
Sub o_93_Open_Daily_To_Do_txt(Optional bolPassedFromWorkbookOpen As Boolean)

' Purpose: To open my Daily To Do txt file.
' Trigger: Ribbon Icon - GTD Macros > Prep for Day > Open Daily To Do
' Updated: 7/26/2022
' Reviewd: 5/20/2023

' Change Log:
'       5/26/2020:  Updated the code to include an InputBox to pick other dates
'       5/30/2020:  Run the VBScript if the file doesn't exist yet
'       6/8/2020:   Updated to include the day of the week in the file name and allow the creation of future dates
'       9/11/2020:  Updated to open the text file using Notepad++
'       11/11/2020: Removed the option to NOT open the file, if it was missing create it then open it
'       12/16/2020: Added the conditional compiler constant to determine the file location if on my Personal computer.
'       4/14/2021:  Updated the code for Notepad++ on my personal comupter to point to the right spot / version
'       8/19/2021:  Simplified the If Statements
'       8/19/2021:  Added a seperate process to assume if you haven't created today's file to then create it
'       10/26/2021: Added the code related to bolPassedFromWorkbookOpen to bypass the Input Box
'       7/26/2022:  Updated the hyperlink locations for the new Planning Template folder

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    ' Dim Objects
    Dim objFSO As Object
        Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Dim Folder Location
    Dim strFolderLoc As String
        #If Personal <> 1 Then
            strFolderLoc = "C:\U Drive\Support\Daily To Do\"
        #Else
            strFolderLoc = "D:\D Documents\Support\Daily To Do\"
        #End If
    
    Dim strFileLoc_DailyTemplate As String
        #If Personal <> 1 Then
            strFileLoc_DailyTemplate = "C:\U Drive\Support\Planning Templates\XXXX.XX.XX - Daily To Do - Template.properties"
        #Else
            strFileLoc_DailyTemplate = "D:\D Documents\Support\Planning Templates\XXXX.XX.XX - Daily To Do - Template.properties"
        #End If
    
    ' Dim Today File Location
    Dim strTodayDateDay As String
        strTodayDateDay = Format(Date, "DDDD")
        
    Dim strTodayDate As String
        strTodayDate = Format(Date, "yyyy.mm.dd")
    
    Dim strTodayFileLoc As String
        strTodayFileLoc = strFolderLoc & strTodayDate & " - " & strTodayDateDay & " - Daily To Do.properties"
        
    Dim bolTodayFileExists As Boolean
        bolTodayFileExists = fx_File_Exists(strTodayFileLoc)
    
    ' Dim Other Date File Location
    
    If bolTodayFileExists = True Then
    
        Dim strDate As String, strDateDay As String
            If bolPassedFromWorkbookOpen = True Then
                strDate = Date
            Else
                strDate = InputBox(Prompt:="What date would you like to open?", Title:="Daily To Do Date Selection", Default:=Date)
            End If
            If StrPtr(strDate) = 0 Then Exit Sub 'Abort if cancel was pushed
            
            strDateDay = Format(strDate, "DDDD")
            strDate = Format(strDate, "yyyy.mm.dd")
        
        Dim strFileLoc As String
            strFileLoc = strFolderLoc & strDate & " - " & strDateDay & " - Daily To Do.properties"
        
    End If

' -----------------------------------------
' Create Today's file (if it doesn't exist)
' -----------------------------------------
        
    If bolTodayFileExists = False Then

        #If Personal <> 1 Then
            objFSO.CopyFile strFileLoc_DailyTemplate, strTodayFileLoc
            Call Shell("explorer.exe" & " " & strTodayFileLoc)
        #Else
            objFSO.CopyFile strFileLoc_DailyTemplate, strTodayFileLoc
            CreateObject("Shell.Application").Open (strTodayFileLoc)
        #End If
        
        Exit Sub

    End If
        
' ----------------------------
' Open the selected day's file
' ----------------------------

    #If Personal <> 1 Then
        If objFSO.FileExists(strFileLoc) = True Then
            Call Shell("explorer.exe" & " " & strFileLoc)
            'Call Shell("%LOCALAPPDATA%\Microsoft\AppV\Client\Integration\414E1AAA-FDD1-48F6-90E0-49CB392CF789\Root\VFS\ProgramFilesX64\Notepad++\notepad++.exe" & strFileLoc) '
        Else
            objFSO.CopyFile strFileLoc_DailyTemplate, strFileLoc
            Call Shell("explorer.exe" & " " & strFileLoc)
        End If
    #Else
        If objFSO.FileExists(strFileLoc) = True Then
            CreateObject("Shell.Application").Open (strFileLoc)
        Else
            objFSO.CopyFile strFileLoc_DailyTemplate, strFileLoc
            CreateObject("Shell.Application").Open (strFileLoc)
        End If
    #End If

End Sub
Sub o_94_Open_Tickler_Folder()

' Purpose: To have a prompt for a date that will then open the applicable day's tickler folder.
' Trigger: Ribbon > GTD Macros > Support > Open Tickler
' Updated: 6/2/2021
' Reviewd: 5/20/2023

' Change Log:
'       9/16/2019:  Initial Creation
'       6/2/2021:   Switched from Shell to FollowHyperlink

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim strTicklerDate As String
        strTicklerDate = InputBox(Prompt:="Insert date you want to open a folder for in mm/dd/yyyy format", Title:="Tickler Folder Date", Default:=Date)
            strTicklerDate = Format(strTicklerDate, "mm.dd.yyyy")

    Dim strTicklerLoc As String
        #If Personal <> 1 Then
            strTicklerLoc = "C:\U Drive\Reference\Tickler Folders\"
        #Else
            strTicklerLoc = "D:\D Documents\Reference\Tickler Folders\"
        #End If

    Dim fldrPath As String
        fldrPath = strTicklerLoc & strTicklerDate

' -----------------------
' Open the tickler folder
' -----------------------

    If Dir(fldrPath, vbDirectory) <> "" Then
        ThisWorkbook.FollowHyperlink (fldrPath)
    Else
        ThisWorkbook.FollowHyperlink (strTicklerLoc)
    End If
    
End Sub
Sub o_95_Create_Dynamic_File_List_in_wsLists()

' Purpose: To create the dictionary that houses the file names, paths, and update dates for the Dynamic File Search.
' Trigger: Workbook_Open Event
' Updated: 11/4/2023
' Reviewd: 11/4/2023

' Change Log:
'       11/4/2023:  Initial Creation, based on o_1_Create_File_List from hte uf_Search_File
'                   Added the code related to the array and putting in wsLists

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------
    
    ' Declare Worksheets
    
    Dim wsLists As Worksheet
    Set wsLists = ThisWorkbook.Sheets("Lists")
    
    ' Declare Strings

    'Dim str_PathParent As String
    Dim str_PathParent As String
        
        #If Personal <> 1 Then
            str_PathParent = "C:\U Drive\"
        #Else
            str_PathParent = "D:\D Documents\"
        #End If
        
    ' Declare Objects
    
    Dim objFSO As Object
        Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim objDirParent As Object
        Set objDirParent = objFSO.GetFolder(str_PathParent)
    
    Dim obj_SubFolder_Lvl1 As Object
    Dim obj_SubFolder_Lvl2 As Object
    Dim obj_SubFolder_Lvl3 As Object

    ' Declare Dictionaries / Arrays

    Dim dict_FilePath As Scripting.Dictionary
    Set dict_FilePath = New Scripting.Dictionary
    
    Dim arry_Files As Variant
    
    ' Declare Loop Variables
    
    Dim dictVal As Variant
    
    Dim objFile As Variant
    
    Dim i As Long
        i = 1
    
' ----------------------------------
' Load the Files into the dictionary
' ----------------------------------

On Error Resume Next

    'Run it for the main folder
        
    objFile = Dir(str_PathParent & "*.*")
    
    Do While objFile <> ""
        dict_FilePath.Add key:=objFile, Item:=str_PathParent & objFile
        objFile = Dir()
    Loop
                        
    For Each obj_SubFolder_Lvl1 In objDirParent.SubFolders 'Loop through 1st SubFolder
        
        objFile = Dir(obj_SubFolder_Lvl1.path & "\" & "*.*")
            
        Do While objFile <> ""
            dict_FilePath.Add key:=objFile, Item:=obj_SubFolder_Lvl1.path & "\" & objFile
            objFile = Dir()
        Loop
        
        For Each obj_SubFolder_Lvl2 In obj_SubFolder_Lvl1.SubFolders 'Loop through 2nd level SubFolder
                
            objFile = Dir(obj_SubFolder_Lvl2.path & "\" & "*.*")
            
            Do While objFile <> ""
                dict_FilePath.Add key:=objFile, Item:=obj_SubFolder_Lvl2.path & "\" & objFile
                objFile = Dir()
            Loop
            
            For Each obj_SubFolder_Lvl3 In obj_SubFolder_Lvl2.SubFolders 'Loop through 3rd level SubFolder (Inception level sh*t right here)

                objFile = Dir(obj_SubFolder_Lvl3.path & "\" & "*.*")

                Do While objFile <> ""
                    dict_FilePath.Add key:=objFile, Item:=obj_SubFolder_Lvl3.path & "\" & objFile
                    objFile = Dir()
                Loop

            Next obj_SubFolder_Lvl3
        Next obj_SubFolder_Lvl2
    Next obj_SubFolder_Lvl1

On Error GoTo 0

' ------------------------------------------
' Output the values to an array then wsLists
' ------------------------------------------

    ReDim arry_Files(1 To dict_FilePath.count, 1 To 2)
    
    For Each dictVal In dict_FilePath
        arry_Files(i, 1) = dictVal
        arry_Files(i, 2) = dict_FilePath(dictVal)
        i = i + 1
    Next dictVal
    
    wsLists.Range("F2").Resize(UBound(arry_Files, 1), 2).Value2 = arry_Files

End Sub
Sub o_99_Run_Backup_Batch()

' Purpose: To open the batch files related to my daily backup for Personal.
' Trigger: Ribbon Icon - GTD Macros > Personal Macros > TBD
' Updated: 1/14/2024
' Reviewd: 5/20/2023

' Change Log:
'       4/24/2021:  Initial Creation
'       6/2/2021:   Switched for Professional to open the folder, not try to run the batch file (Webster blocks it).
'                   Swithed from using Shell to FollowHyperlink, much faster
'       1/14/2024:  Updated the files for my personal backup

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    'Dim Objects
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    'Dim File Locations
    Dim strFolderLoc As String
        #If Personal <> 1 Then
            strFolderLoc = "C:\U Drive\Reference\Templates, Scripts & Batch Files\Batch Files\"
        #Else
            strFolderLoc = "D:\D Documents\Reference\Templates, Scripts & Batch Files\Batch Files\"
        #End If
                
' --------------------
' Open the batch files
' --------------------

        #If Personal <> 1 Then
            ThisWorkbook.FollowHyperlink (strFolderLoc)
        #Else
            Call Shell(strFolderLoc & "1.1 Local Documents Backup.bat " & strFolderLoc)
            Call Shell(strFolderLoc & "1.2 USB Documents Backup.bat " & strFolderLoc)
            Call Shell(strFolderLoc & "1.3 Backup Documents to Dropbox Folder.bat " & strFolderLoc)
            Call Shell(strFolderLoc & "1.4 Backup Documents to Google Drive Folder.bat " & strFolderLoc)
        #End If
       
End Sub
