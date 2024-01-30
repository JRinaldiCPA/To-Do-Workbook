VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_New_Project 
   Caption         =   "Project Input Form"
   ClientHeight    =   3240
   ClientLeft      =   168
   ClientTop       =   768
   ClientWidth     =   7380
   OleObjectBlob   =   "uf_New_Project.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_New_Project"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ws_Projects As Worksheet

Public curRow As Long
Private Sub UserForm_Initialize()
    
' Purpose: To initialize the userform, including adding in the data from the arrays.
' Trigger: Event: UserForm_Initialize
' Updated: 9/10/2023

' Change Log:
'       9/4/2021:   Made a number of small tweaks to the formatting and size of TxtBoxes
'       9/5/2021:   Added the Zoom Adjust related code
'       6/28/2023:  Updated to adjust the height for Personal vs Professional
'       9/10/2023:  Removed the reference to txt_Status as I now default to Active
'                   Replaced the lst_Area with cmb_Area, removed lst_Entered, removed lst_Updated, removed strRequest_Num

' ***********************************************************************************************************************************

Call myPrivateMacros.DisableForEfficiency

' ----------------------
' Declare your variables
' ----------------------

    'Dim ws_Projects As Worksheet
        Set ws_Projects = ThisWorkbook.Sheets("Projects")
    
    'Dim CurRow as Long
        curRow = [MATCH(TRUE,INDEX(ISBLANK('[To Do.xlsm]Projects'!A:A),0),0)]
                
    Dim i As Long
    
    Dim dblZoomAdjust As Double
        dblZoomAdjust = 1.05

' -----------------------------
' Initialize the initial values
' -----------------------------
    
    'Add the values for the Area ComboBox
        Me.cmb_Area.List = GetAreaArray
        'Me.cmb_Area.RemoveItem (6) 'Added 11/27/23 to remove "D/A Requests" as an option
        
' -------------------------------
' Adjust the size of the UserForm
' -------------------------------
        
    Me.Zoom = Me.Zoom * dblZoomAdjust
    Me.Height = Me.Height * dblZoomAdjust
    Me.Width = Me.Width * dblZoomAdjust

' --------------------------------------------
' Set the current focus to the Project TextBox
' --------------------------------------------
        
    Me.txt_Project.SetFocus
    
Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Private Sub cmd_Add_Project_Click()

' Purpose: To do add the project to my To Do workbook.
' Trigger: Click Add Project
' Updated: 9/16/2019

' Change Log:
'       9/16/2019:  Initial Creation

' ***********************************************************************************************************************************

'Check if a task was added
If Me.txt_Project.Value = "" Then
    MsgBox "The Project description text box was blank, a new Project was not created"
    Unload Me
End If

Call Me.o_21_Add_New_Project
    Call Me.o_22_Create_Project_Folder
        Unload Me

End Sub
Private Sub cmd_Cancel_Click()
    
    Unload Me

End Sub
Sub o_21_Add_New_Project()

' Purpose: To allow me to quickly create a new Project for my To Do Excel via UserForm
' Trigger: Called: uf_New_Project
' Updated: 9/10/2023

' Change Log:
'       12/10/2021: Updated to reflect the P.XXX as part of the project's name
'                   Added in the code from creating a D/A request to include the year
'       9/10/2023:  Updated to default the Status to be Active
'                   Updated to default the Updated and Entered to be the current date
'                   Updated to default the Req # instead of having it be an option to update

' ***********************************************************************************************************************************
    
Call myPrivateMacros.DisableForEfficiency

' -----------------
' Declare Variables
' -----------------

    Dim strRequest_Num As String
        strRequest_Num = Application.WorksheetFunction.Max(ws_Projects.Range("A1:A" & curRow)) + 1
        
        ' Pad the zeros
        If Len(strRequest_Num) = 1 Then
            strRequest_Num = "00" & strRequest_Num
        ElseIf Len(strRequest_Num) = 2 Then
            strRequest_Num = "0" & strRequest_Num
        End If
        
    Dim strYear As String
        strYear = Format(Date, "YY")

' -----------
' Add the new Project
' -----------
    
    With ws_Projects
    
    'Column A: Project #
        .Range("A" & curRow).Value = strRequest_Num
       
    'Column B: Entered Date
        .Range("B" & curRow).Value = Date
       
    'Column C: Area
        .Range("C" & curRow).Value = cmb_Area.Value
            If cmb_Area.Value = "" Then .Range("C" & curRow).Value = "N/A"
        .Range("C" & curRow).Validation.Add Type:=xlValidateList, Formula1:=Join(GetAreaArray, ",")
      
    'Column D: Project Name
        .Range("D" & curRow).Value = "P." & strYear & "." & strRequest_Num & " - " & Me.txt_Project.Value
    
    'Column E: Project Objective
        .Range("E" & curRow).Value = Me.txt_Project_Obj.Value
    
    'Column F: Update Date
        .Range("F" & curRow).Value = Date
    
    'Column G: Status
        .Range("G" & curRow).Value = "Active"
            .Range("G" & curRow).Validation.Add Type:=xlValidateList, Formula1:=Join(GetStatusArray, ",")
          
' -----------
' Apply formatting to the new row
' -----------

    'Apply upper and lower grey line, and correct font/size, to all cells
        With .Range(.Cells(curRow, "A"), .Cells(curRow, "H"))
            .Borders(xlEdgeBottom).Color = RGB(217, 217, 217)
            .Borders(xlEdgeTop).Color = RGB(217, 217, 217)
            .Font.Name = "Calibri"
            .NumberFormat = "General"
            .Font.Size = 11
        End With

    'Apply the formating to the #
        With .Range("A" & curRow)
            .Font.Color = RGB(184, 0, 0)
            .Font.Bold = "True"
            .Interior.Color = RGB(242, 242, 242)
            .NumberFormat = "0"
        End With
    
    'Apply the center alignment
                
        Union(.Range(.Cells(curRow, "A"), .Cells(curRow, "C")), .Range("F" & curRow), .Range("G" & curRow)).HorizontalAlignment = xlCenter

        .Range(.Cells(curRow, "A"), .Cells(curRow, "H")).VerticalAlignment = xlCenter

    'Apply word wrap to the Objective and Status columns
        
        Union(.Range("E" & curRow), .Range("H" & curRow)).WrapText = True
        
    'Apply the custom date formatting
        
        Union(.Range("B" & curRow), .Range("F" & curRow)).NumberFormat = "m.d.yyyy"
           
  End With

Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_22_Create_Project_Folder()

' Purpose: To create the folder that is used to house my project support.
' Trigger: Called: uf_New_Project
' Updated: 6/28/2023

' Change Log:
'       9/16/2019:  Initial Creation
'       11/9/2021:  Updated to reflect new Personal Areas of Focus
'       12/10/2021: Updated to only use rngProjectName for creating the folder
'       5/7/2022:   Converted "House / Yard" and "Financial" to "Household"
'       7/18/2022:  Updated to include the P.XXX Support text file
'                   Removed unused variables
'       9/20/2022:  Removed the 'Why' and 'Success' fields, now that those are tracked in the .properties file
'                   Removed the Support Workbook, now that it has been replaced by the .properties file
'       6/28/2023:  Updated to use RGB color codes and added some Personal Areas of Focus
'                   Moved the updated formatting code to o_15_Apply_Projects_Formatting to reduce redundancy

' ***********************************************************************************************************************************

Call myPrivateMacros.DisableForEfficiency
 
' -----------------
' Declare Variables
' -----------------
    
    Dim objFSO As Object
        Set objFSO = VBA.CreateObject("Scripting.FileSystemObject")
    
    'Declare Workbooks / Sheets
    
    Dim wb_PXXX As Workbook
    
    Dim ws_PXXX_Proj As Worksheet
    
    'Declare Ranges
    
    Dim rngProjectArea As Range
        Set rngProjectArea = ws_Projects.Range("C" & curRow)
    
    Dim rngProjectName As Range
        Set rngProjectName = ws_Projects.Range("D" & curRow)
            
    'Declare Folder Paths
    
    Dim ProjectFolder As String
        #If Personal <> 1 Then
            ProjectFolder = "C:\U Drive\Projects\" & rngProjectName
        #Else
            ProjectFolder = "D:\D Documents\Projects\" & rngProjectName
        #End If
            
    'Declare Text File Paths
        
    Dim str_PXXXUpdatedPath_txt As String
        str_PXXXUpdatedPath_txt = ProjectFolder & "\" & rngProjectName & ".properties"
    
    Dim str_ProjectSupportPath_txt As String
        #If Personal <> 1 Then
            str_ProjectSupportPath_txt = "C:\U Drive\Support\Planning Templates\P.XXX - Support.properties"
        #Else
            str_ProjectSupportPath_txt = "D:\D Documents\Support\Planning Templates\P.XXX - Support.properties"
        #End If
    
    'Declare Integers
    
    Dim NewProject_row As Long
    
    Dim NewProject_num As Long
                
' ------------------------------------------------------------
' Create the new folder for the project based on the Project #
' ------------------------------------------------------------

    If Dir(ProjectFolder) = "" Then MkDir (ProjectFolder)
    
    MkDir (ProjectFolder & "\(ARCHIVE)")
            
' ---------------------------------
' Copy in the Project Support Files
' ---------------------------------
           
    Call objFSO.CopyFile(str_ProjectSupportPath_txt, ProjectFolder & "\")
    
' -------------------------------------------
' Rename the Support Files with the Project #
' -------------------------------------------

    Name ProjectFolder & "\P.XXX - Support.properties" As str_PXXXUpdatedPath_txt

' ------------------
' Add the hyperlinks
' ------------------

    ws_Projects.Hyperlinks.Add Anchor:=rngProjectArea, Address:=ProjectFolder
    ws_Projects.Hyperlinks.Add Anchor:=rngProjectName, Address:=str_PXXXUpdatedPath_txt
    
' ----------------------------
' Apply the project formatting
' ----------------------------
    
    Call Macros.o_15_Apply_Projects_Formatting
    
Call myPrivateMacros.DisableForEfficiencyOff

End Sub
