VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_New_DARequest 
   Caption         =   "New D/A Request"
   ClientHeight    =   5328
   ClientLeft      =   36
   ClientTop       =   252
   ClientWidth     =   7356
   OleObjectBlob   =   "uf_New_DARequest.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_New_DARequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Dim sheets
    Public ws_Projects As Worksheet

'Dim Request Variables
    Public DA_Requestor As String
    Public DA_RequestName As String
    Public RequestLoc1 As Range
    Public RequestLoc2 As Range
    Public DA_Type As String
    Public DA_Objective As String
      
'Dim integers
    Public int_CurRow As Long

Private Sub UserForm_Initialize()

' Purpose: To initialize the userform, including adding in the data from the arrays.
' Trigger: Event: UserForm_Initialize
' Updated: 4/3/2023

' Change Log:
'       9/5/2021: Added the Zoom Adjust related code
'       4/3/2023:   Converted the COUNTIF to MAXIFS to account for old analytics requests being archived

' ***********************************************************************************************************************************

' ----------------------
' Declare your variables
' ----------------------
   
    'Dim ws_Projects As Worksheet
        Set ws_Projects = ThisWorkbook.Sheets("Projects")
   
   'Dim int_CurRow As Long
        int_CurRow = Evaluate("MATCH(TRUE,INDEX(ISBLANK('[To Do.xlsm]Projects'!A:A),0),0)")
        
    Dim intRequestNum As Long
        'intRequestNum = Evaluate("COUNTIF('[To Do.xlsm]Projects'!C:C,""D/A Requests"")") + 1
        intRequestNum = Evaluate("MAXIFS('[To Do.xlsm]Projects'!A:A,'[To Do.xlsm]Projects'!C:C,""D/A Requests"")") + 1 'Created on 4/3/2023
        
    If ActiveWorkbook.Name Like "DA *" Then
        Dim ws_DA_Intake_Form
            Set ws_DA_Intake_Form = ActiveWorkbook.Sheets(1)
        
        Dim str_Req_Num As String
            str_Req_Num = Right(Left(ws_DA_Intake_Form.Parent.Name, InStr(1, ws_DA_Intake_Form.Parent.Name, " - ") - 1), 3)
        
        Dim str_Requestor As String
            str_Requestor = fx_Reverse_Name(ws_DA_Intake_Form.Range("D5").Value) 'Added on 10/11/2019
            
    End If
        
    Dim dblZoomAdjust As Double
        dblZoomAdjust = 1.05
        
' -----------------------------
' Initialize the initial values
' -----------------------------
    
    'Add value for D/A Request # TextBox
        Me.txt_Req_Num.Value = intRequestNum
        
    'Add today's date for the Entered TextBox
        Me.txt_Requested.Value = Date
    
    'Add today's date for the Updated TextBox
        Me.txt_Updated.Value = Date
    
    'Add today's date for the Status TextBox
        Me.txt_Status.Value = "Active"
        
    'Add the values for the Request Type ListBox
        
        Me.lst_Request_Type.List = GetDARequestTypeArray
                
' -------------------------------
' Adjust the size of the UserForm
' -------------------------------
        
    Me.Zoom = Me.Zoom * dblZoomAdjust
    Me.Height = Me.Height * dblZoomAdjust
    Me.Width = Me.Width * dblZoomAdjust
                
' ----------------------------------------------
' Set the current focus to the Requestor TextBox
' ----------------------------------------------
        
    Me.txt_Requestor.SetFocus

End Sub
Private Sub txt_Title_Enter()

' Purpose: To pre-populate the D/A Request name.
' Trigger: Enter D/A Request Name
' Updated: 12/10/2021

' Change Log:
'       12/10/2021: Cleaned Up
'       3/3/2022:   Removed the Requestor Name as part of the request title

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim strRequest_Num As String
        strRequest_Num = Me.txt_Req_Num.Value
        
        ' Pad the zeros
        If Len(strRequest_Num) = 1 Then
            strRequest_Num = "00" & strRequest_Num
        ElseIf Len(strRequest_Num) = 2 Then
            strRequest_Num = "0" & strRequest_Num
        End If
    
    Dim strYear As String
        strYear = Format(Date, "YY")
    
' -------------------------
' Output the D/A Request ID
' -------------------------

    Me.txt_Title.Value = "DA." & strYear & "." & strRequest_Num & " - "

End Sub
Private Sub cmd_Add_DARequest_Click()

 Call myPrivateMacros.DisableForEfficiency

'Check if a task was added
If Me.txt_Request.Value = "" Then
    MsgBox "The D/A Request text box was blank, a new D\A Requst was not created"
    Unload Me
End If

Call Me.o_11_Add_New_DARequest
Call Me.o_12_Format_New_DARequest
Call Me.o_2_Set_Request_Values
    
    If Me.lst_Request_Type.Value <> "Ad-Hoc - Quick" Then Call Me.o_3_Create_DA_Request_Folder
    
    If ActiveWorkbook.Name Like "DA *" Then
        Call Shell("explorer.exe" & " " & "C:\U Drive\Analytics Requests\" & Me.txt_Title, vbNormalFocus)
    End If
    
    Unload Me

 Call myPrivateMacros.DisableForEfficiencyOff

End Sub
Private Sub cmd_Cancel_Click()

    Unload Me

End Sub
Sub o_11_Add_New_DARequest()

' Purpose: To allow me to quickly create a new D/A Request for my To Do Excel via UserForm
' Trigger: Called: uf_New_DARequest
' Updated: 9/16/2019

' ***********************************************************************************************************************************
   
' -----------
' Add the new Project
' -----------
    
    With ws_Projects
    
    'Column A: Project #
        .Range("A" & int_CurRow).Value = Me.txt_Req_Num.Value
       
    'Column B: Requested
        .Range("B" & int_CurRow).Value = Me.txt_Requested.Value
    
    'Column C: Area
        .Range("C" & int_CurRow).Value = "D/A Requests"
    
    'Column D: D/A Request Title
        .Range("D" & int_CurRow).Value = Me.txt_Title.Value
    
    'Column E: Request
        .Range("E" & int_CurRow).Value = Me.txt_Request.Value
    
    'Column F: Updated Date
        .Range("F" & int_CurRow).Value = Me.txt_Updated.Value
    
    'Column G: Status Desc.
        .Range("G" & int_CurRow).Value = Me.txt_Status.Value
        .Range("G" & int_CurRow).Validation.Add Type:=xlValidateList, Formula1:=Join(GetStatusArray, ",")
          
    End With

End Sub
Sub o_12_Format_New_DARequest()

' -----------
' Apply formatting to the new row
' -----------

    With ws_Projects

    'Apply upper and lower grey line, and correct font/size, to all cells
        With .Range(.Cells(int_CurRow, "A"), .Cells(int_CurRow, "H"))
            .Borders(xlEdgeBottom).Color = RGB(217, 217, 217)
            .Borders(xlEdgeTop).Color = RGB(217, 217, 217)
            .Font.Name = "Calibri"
            .NumberFormat = "General"
            .Font.Size = 11
        End With

    'Apply the formating to the #
        With .Range("A" & int_CurRow)
            .Font.Color = RGB(184, 0, 0)
            .Font.Bold = "True"
            .Interior.Color = RGB(242, 242, 242)
            .NumberFormat = "0"
        End With
    
    'Apply the center alignment
                
        Union(.Range(.Cells(int_CurRow, "A"), .Cells(int_CurRow, "C")), .Range("F" & int_CurRow), .Range("G" & int_CurRow)).HorizontalAlignment = xlCenter

        .Range(.Cells(int_CurRow, "A"), .Cells(int_CurRow, "H")).VerticalAlignment = xlCenter

    'Apply word wrap to the Objective and Status columns
        
        Union(.Range("E" & int_CurRow), .Range("H" & int_CurRow)).WrapText = True
        
    'Apply the custom date formatting
        
        Union(.Range("B" & int_CurRow), .Range("F" & int_CurRow)).NumberFormat = "m.d.yyyy"
        
  End With

End Sub
Sub o_2_Set_Request_Values()

' Purpose: To set the various variables related to the DA Request, based on what was input into the UserForm.
' Trigger: Called: uf_New_DARequest
' Updated: 4/11/2020

' ***********************************************************************************************************************************

' -----------
' Declare your variables
' -----------

    'Dim DA_Requestor As String
        DA_Requestor = Me.txt_Requestor
        
    'Dim DA_RequestName As String
        DA_RequestName = Me.txt_Title
    
    'Dim RequestLoc1 As Range
        Set RequestLoc1 = ws_Projects.Range("C" & int_CurRow) 'Project Area

    'Dim RequestLoc2 As Range
        Set RequestLoc2 = ws_Projects.Range("D" & int_CurRow) ' Project Name

    'Dim DA_Type As String
        DA_Type = Me.lst_Request_Type.Value

    'Dim DA_Objective As String
        DA_Objective = Me.txt_Request

    'Dim DA_RequestedFor As String
        DA_RequestedFor = Me.txt_Requestor

End Sub
Sub o_3_Create_DA_Request_Folder()

' Purpose: To create the folder that is used to house my project support
' Trigger: Called: uf_New_DARequest
' Updated: 7/20/2022
'
' Change Log:
'       12/15/2020: Updated to format the row with the green and height of 45
'       7/20/2022:  Updated to include the DA.XXX Support text file
'                   Updated to mirror the code for the P.XXX Support folder

' ***********************************************************************************************************************************
 
' -----------------
' Declare Variables
' -----------------
    
    'Declare Objects
        
        Dim objFSO As Object
            Set objFSO = VBA.CreateObject("Scripting.FileSystemObject")
        
    'Declare Workbooks / Worksheets
                    
        Dim DA_XXX_wb As Workbook
        
        Dim DA_XXX_Request_ws As Worksheet
            
    'Declare Folder Path
            
        Dim RequestFolder As String
            RequestFolder = "C:\U Drive\Analytics Requests\" & DA_RequestName
    
    'Declare Workbook Paths
    
        Dim str_DARequest_UpdatedPath_wb As String
            str_DARequest_UpdatedPath_wb = RequestFolder & "\" & DA_RequestName & ".xlsx"
        
    'Declare Text File Paths
    
    
' ----------------------------------------------------------------
' Create the new folder for the project based on the D/A Request #
' ----------------------------------------------------------------
    
    If Dir(RequestFolder) = "" Then MkDir (RequestFolder)
    
' -------------------------------------
' Copy in the D/A Request Support Files
' -------------------------------------
            
    Call objFSO.CopyFile("C:\U Drive\Support\DA XXX Support\*", RequestFolder & "\") 'Copy the DA.XXX Support files
    Call objFSO.CopyFolder("C:\U Drive\Support\DA XXX Support\*", RequestFolder & "\") 'Copy the sub folders
        
' --------------------------------------------------
' Rename the Support Files with the D/A Request Name
' --------------------------------------------------
                            
    Name RequestFolder & "\DA XXX Support.xlsx" As RequestFolder & "\" & DA_RequestName & ".xlsx"
        
    Name RequestFolder & "\DA XXX Support.properties" As RequestFolder & "\" & DA_RequestName & ".properties"
                            
' ------------------
' Add the hyperlinks
' ------------------

    ws_Projects.Hyperlinks.Add Anchor:=RequestLoc1, Address:=RequestFolder
    ws_Projects.Hyperlinks.Add Anchor:=RequestLoc2, Address:=str_DARequest_UpdatedPath_wb

' -------------------------------
' Apply the formating in Projects
' -------------------------------

With ws_Projects
    
    .Range(.Cells(int_CurRow, "B"), .Cells(int_CurRow, "H")).Interior.Color = 13891050 ' dark green
        .Rows(int_CurRow).RowHeight = 45

    'This fixes the issue of hyperlinks being blue and underlined by making all font black, no underline, and Size 11 Calibri
        With .Range(.Cells(int_CurRow, "B"), .Cells(int_CurRow, "H")).Font
            .ColorIndex = xlAutomatic
            .Underline = xlUnderlineStyleNone
            .Name = "Calibri"
            .Size = 11
        End With

End With

' ---------------------------------------
' Copy data into the DA XXX Support Excel
' ---------------------------------------

    Set DA_XXX_wb = Workbooks.Open(str_DARequest_UpdatedPath_wb)
    
    Set DA_XXX_Request_ws = DA_XXX_wb.Sheets("Request")

    'Copy in the data
        With DA_XXX_Request_ws
            .Range("D5").Value2 = DA_Requestor
            .Range("D6").Value2 = Date
            .Range("D7").Value2 = DA_RequestName
            .Range("D8").Value2 = DA_Type
            .Range("D9").Value2 = DA_Objective
            
        End With
        
        DA_XXX_wb.Close SaveChanges:=True

End Sub

