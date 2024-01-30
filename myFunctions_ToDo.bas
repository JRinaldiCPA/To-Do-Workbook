Attribute VB_Name = "myFunctions_ToDo"
Option Explicit
Declare PtrSafe Function fx_Get_Screen_Resolution Lib "User32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Public Function myFunc_My_Experience() As String

' Purpose: To easily determine how many years of experience I have.
' Trigger: Called via formula
' Updated: 8/31/2023

' Change Log:
'       9/5/2020:   Added in my Sr. Analyst role
'       6/8/2021:   Updated to reflect my switch in roles.
'       10/6/2022:  Added the Audit and Analytics days
'       8/31/2023:  Updated to be more precise about my Webster time, and include my new ITRM role

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------
    
    ' Job Variables
    Dim intBarnesDays As Long
        intBarnesDays = DateValue("08/25/06") - DateValue("05/15/06") + 21
    
    Dim intKPMGDays As Long
        intKPMGDays = DateValue("11/12/08") - DateValue("09/17/07")
    
    Dim intWebsterDays As Long
        intWebsterDays = (DateValue("06/22/14") - DateValue("07/12/10")) + (Date - DateValue("03/16/15"))
    
    Dim intBoADays As Long
        intBoADays = DateValue("03/15/15") - DateValue("06/23/14")
    
    ' Time Calc Variables
    Dim intAuditDays As Long
        intAuditDays = intBarnesDays + intKPMGDays + intBoADays + (DateValue("06/20/14") - DateValue("07/12/10")) + (DateValue("07/01/17") - DateValue("03/15/15"))
    
    Dim intAnalyticsDays As Long
        intAnalyticsDays = DateValue("07/30/23") - DateValue("07/01/17")
    
    Dim intRiskManagementDays As Long
        intRiskManagementDays = Date - DateValue("07/31/23")
    
    Dim intTotalDays As Long
        intTotalDays = intBarnesDays + intKPMGDays + intWebsterDays + intBoADays

' ------------------
' Create the message
' ------------------

    myFunc_My_Experience = "I have " & Round(intTotalDays / 365, 1) & " years experience, of which " & Round(intWebsterDays / 365, 1) & " years was at Webster.  " & Chr(10) & _
    "I switched to an analytics role in January 2017, took on PP in March 2019, switched to Credit Risk in April 2020, and switched to IT in July 2023.  " & Chr(10) & _
    "I have a total of " & Round(intAuditDays / 365, 1) & " years experience in Audit, " _
     & Round(intAnalyticsDays / 365, 1) & " years experience in Analytics, and " _
     & Round(intRiskManagementDays / 365, 1) & " years experience in IT Risk Management." _

' ***********************************************************************************************************************************

'Title                      Company             Start Date      End Date        Time                Cumulative Time
'MD, IT Risk Management     Webster Bank        07/31/2023      N/A             01 months           14 years 08 months
'MD, Credit Analytics       Webster Bank        10/24/2022      07/28/2023      09 months           14 years 06 months
'Sr. Credit Risk Analyst    Webster Bank        04/09/2020      10/23/2022      31 months           13 years 09 months
'Sr. Mgr. Audit Analytics   Webster Bank        10/25/2018      04/08/2020      13 months           11 years 03 months
'Audit Program Manager      Webster Bank        11/12/2015      10/24/2018      36 months           09 years 09 months 'But switched to analytics semi-FT in January 2017
'Audit Supervisor           Webster Bank        03/16/2015      11/11/2015      08 months           06 years 10 months
'Senior Auditor II          Bank of America     06/23/2014      03/13/2015      09 months           06 years 02 months
'Audit Supervisor           Webster Bank        06/13/2013      06/20/2014      12 months           05 years 05 months
'Audit Senior               Webster Bank        05/05/2011      06/12/2013      26 months           04 years 05 months
'Audit Staff                Webster Bank        07/12/2010      05/04/2011      10 months           02 years 04 months
'Audit Staff                KPMG                09/17/2007      11/12/2008      14 months           01 years 06 months
'Audit Intern               Barnes Group        05/15/2006      08/25/2006      04 months           00 years 04 months '+ 21 days during the winter

End Function
Public Function fx_Create_Project_ID(intProjectNbr As String, ProjectName As String) As String
    
' Purpose: This function creates the ProjectID based on the input given.
' Trigger: Called Function
' Updated: 9/16/2019

' ***********************************************************************************************************************************
    
    Dim ProjectId As String

    ProjectId = "P." & intProjectNbr & " - " & ProjectName

    fx_Create_Project_ID = ProjectId
 
End Function
Public Function fx_Copy_Project_ID_to_Clipboard(Optional strArea As String, Optional strProject As String)

' Purpose: To create the full Project ID for my projects or my staff's projects.
' Trigger: Ribbon > Personal Macros > Functions > Copy ProjectID to Clipboard
' Updated: 12/10/2021

' Change Log:
'       6/8/2020:   Updated the code to account for a DA project in Projects ws
'       7/26/2021:  Stripped out the reference to the Deb / Axcel workbooks
'       10/25/2021: Updated to allow variables to be passed from a UserForm to create the ID
'       12/10/2021: Removed the code related to converting a project to P.XXX
'                   Simplified to remove the "GoTo" now that I don't use the D/A vs P.XXX Code
'       5/3/2022:   Renamed from 'u_Copy_Project_ID_to_Clipboard' to 'fx_Copy_Project_ID_to_Clipboard'

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim objDataObj As New MSForms.DataObject
    
    Dim ws_Projects As Worksheet
    Set ws_Projects = ThisWorkbook.Sheets("Projects")
    
    Dim Project_ID As String
    
' -------------------------------------------------
' Pick the path to follow and create the Project_ID
' -------------------------------------------------

    If ActiveSheet.Name = ws_Projects.Name Then
        Project_ID = ws_Projects.Range("D" & Selection.Row).Value
    Else
        Project_ID = strProject
    End If
    
' --------------------------------------
' Copy the Project ID into the clipboard
' --------------------------------------
    
    objDataObj.SetText Project_ID
        objDataObj.PutInClipboard

End Function
Public Function fx_Copy_to_Clipboard(Optional strTextToCopy As String) As String

' Purpose: This function will copy data into the clipboard, from the selected ragne or .
' Trigger: Called Function
' Updated: 2/3/2022

' Change Log:
'       12/16/2020: Reduced the amount of code related to copying the selection, no longer used
'       12/24/2020: Added back in the copy paste from Seleciton in the Error Handler
'       2/3/2022:   Updated to use the new fx_Copy_from_Clipboard function, and disabled wiping the clipboard
'                   Added the strTextToCopy variable and related code
'                   Updated w/ the HTML code

' ***********************************************************************************************************************************

On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------

    Dim strText As String

    Dim varStrText As Variant
        varStrText = strTextToCopy
    
    Dim rng As Range
        Set rng = Selection
        
    Dim objDataObj As New MSForms.DataObject
        'objDataObj.GetFromClipboard
    
    Dim i As Long

' ----------------------------------------------------------
' If strTextToCopy was passed then put that in the Clipboard
' ----------------------------------------------------------

    If strTextToCopy <> "" Then
        CreateObject("htmlfile").parentWindow.clipboardData.setData "text", varStrText
        Exit Function
    End If

' -------------------------------------------
' Copy a large range of data to the Clipboard
' -------------------------------------------
    
    If rng.Rows.count > 1 Then
    
        For i = rng.Row To rng.Row - 1 + rng.Rows.count
            If Cells(i, rng.Column).Value <> "" Then
                strText = strText & Cells(i, rng.Column).Value & Chr(10) & Chr(10)
            End If
        Next i
        
        fx_Copy_to_Clipboard = Left$(strText, Len(strText) - 2) 'Remove trailing linebreaks
        Exit Function
    
    End If

' -----------------------------------------------
' Copy to the Clipboard if it's not a large range
' -----------------------------------------------
    
    strText = fx_Copy_from_Clipboard
    
    'Application.CutCopyMode = False
    
    If Right$(strText, 2) = vbCrLf Or Right$(strText, 2) = vbNewLine Then 'Remove trailing linebreaks
        strText = Left$(strText, Len(strText) - 2)
    End If
    
    fx_Copy_to_Clipboard = strText

ErrorHandler:
    'Debug.Print "There was an error with the fx_Copy_To_Clpiboard on: " & Date & " at "; Time
    
End Function
Public Function fx_Copy_from_Clipboard() As String

' Purpose: This function will copy the data from the clipboard and pass as a String.
' Trigger: Called Function
' Updated: 2/3/2022

' Change Log:
'       2/3/2022:   Reduced the amount of code related to copying the selection, no longer used

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim objDataObj As New MSForms.DataObject
        objDataObj.GetFromClipboard
    
    On Error Resume Next
    Dim strClipboardContents As String
        strClipboardContents = Trim(objDataObj.GetText(1))
    On Error GoTo 0
' ---------------
' Pass the String
' ---------------
        
    fx_Copy_from_Clipboard = strClipboardContents
    
'    If Right$(strText, 2) = vbCrLf Or Right$(strText, 2) = vbNewLine Then strText = Left$(strText, Len(strText) - 2) 'Remove trailing linebreaks

End Function
Public Function fx_RefX_Length(strTodayFileLoc)

' Purpose: This function determines the length of a text file, and returns the length as an integer.
' Trigger: Called Function
' Updated: 9/16/2019

' ***********************************************************************************************************************************

' -----------
' Declare your variables
' -----------

    Dim intTxtFile As Long
        intTxtFile = FreeFile
        
    Dim FilePath As String
        FilePath = strTodayFileLoc
    
    Dim FileLength As Long
    
' -----------
' Determine the file length
' -----------

    'Open the text file in Read Only mode to pull the current content
        Open FilePath For Input As intTxtFile
            FileLength = LOF(intTxtFile)
            
' -----------
' Output the link to the file
' -----------
  
    fx_RefX_Length = FileLength
            
End Function
Public Sub fx_Backup_to_ARCHIVE()
Attribute fx_Backup_to_ARCHIVE.VB_ProcData.VB_Invoke_Func = "S\n14"

' Purpose: This function backs up the applicable document to the (ARCHIVE) folder.
' Trigger: Called / Ctrl + Shift + S
' Updated: 10/7/2019

' ***********************************************************************************************************************************

On Error GoTo ErrorHandler

' -----------------
' Declare Variables
' -----------------

   Dim objFSO As Object
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
    Dim wbName As String
        wbName = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5)
        
    Dim strDir1 As String
        strDir1 = ActiveWorkbook.path & "\(ARCHIVE)"
        
    Dim strDir2 As String
        strDir2 = ActiveWorkbook.path & "\[ARCHIVE]"
        
    Dim strDir3 As String
        strDir3 = ActiveWorkbook.path
        
    Dim strDirFinal As String
    
    Dim wbExt As String
        wbExt = Right(ActiveWorkbook.Name, 5)
        
' --------------------------------------------------
' Determine if the folder is there, and abort if not
' --------------------------------------------------
    
    If Dir(strDir1, vbDirectory) <> vbNullString Then
        strDirFinal = strDir1 & "\"
    ElseIf Dir(strDir2, vbDirectory) <> vbNullString Then
        strDirFinal = strDir2 & "\"
    Else
        strDirFinal = strDir3 & "\"
    End If
    
    objFSO.CopyFile _
        Source:=ActiveWorkbook.FullName, _
        Destination:=strDirFinal & wbName & " (" & Format(Date, "yyyy.mm.dd") & " - " & Format(Time, "hh.mm") & ")" & wbExt

Exit Sub
    
ErrorHandler:
    MsgBox fx_Error_Handler(Err.Number, Err.Source, Err.Description)
    
End Sub
Public Function fx_Create_Folder(strFullPath As String)

' Purpose: This function will create a folder if it doesn't already exist.
' Trigger: Called
' Updated: 9/18/2019

' ***********************************************************************************************************************************

    If Dir(strFullPath, vbDirectory) = vbNullString Then MkDir (strFullPath)
    
    Dim objFSO As Object
    
    fx_Create_Folder = objFSO.GetFolder(strFullPath)

End Function
Function fx_Create_File_Dictionary(strFolderPath As String) As Dictionary

' Purpose: To create a dictionary of files for a given folder path.
' Trigger: Called
' Updated: 8/9/2021

' Use Example: _
    Set dict_Files = fx_Create_File_Dictionary(strFolderPath:="C:\U Drive\Support\Weekly Plan\")

' Change Log:
'       8/9/2021: Initial Creation

' ***********************************************************************************************************************************

' ----------------------
' Declare your variables
' ----------------------

    ' Declare Objects

    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim objDirParent As Object
    Set objDirParent = objFSO.GetFolder(strFolderPath)
    
    Dim objFile As Object
    
    ' Declare Dictionary
    
    Dim dict_File As Scripting.Dictionary
    Set dict_File = New Scripting.Dictionary

' ----------------------------------
' Load the Files into the dictionary
' ----------------------------------

On Error Resume Next

    For Each objFile In objDirParent.Files
        dict_File.Add key:=objFile.Name, Item:=objFile.path
    Next objFile
                        
On Error GoTo 0

' Output the results
Set fx_Create_File_Dictionary = dict_File

End Function
Function fx_Get_Most_Recent_File_From_Directory_Based_On_File_Name(strFolderPath As String) As String

' Purpose: To loop through a directory and return the most recent file (based on file name) and the path to that file.
' Trigger: Called
' Updated: 8/12/2021

' Use Example: _
    strMostRecentFilePath = fx_Get_Most_Recent_File_From_Directory_Based_On_File_Name(strFolderPath:="C:\U Drive\Support\Weekly Plan\")

' Change Log:
'       8/12/2021: Initial Creation

' ***********************************************************************************************************************************

' ----------------------
' Declare your variables
' ----------------------

    ' Declare Dates
    
    Dim dtCurrFile As Date
    
    ' Declare Dictionaries

    Dim dict_Files As Scripting.Dictionary
    Set dict_Files = fx_Create_File_Dictionary(strFolderPath:=strFolderPath)
        
    Dim dict_FileDates As Scripting.Dictionary
    Set dict_FileDates = New Scripting.Dictionary

    ' Declare Loop Variables
    
    Dim val As Variant

' -----------------------------------------
' Determine the most recent Weekly Review
' ------------------------------------------------

    ' Create the Dictionary of File Dates
    
    For Each val In dict_Files.Keys
        dict_FileDates.Add key:=CDate(Replace(Mid(String:=val, Start:=15, Length:=10), ".", "/")), Item:=val
    Next val
    
    ' Find the most recent date

    For Each val In dict_FileDates.Keys
        If dtCurrFile < val Then
            dtCurrFile = val
        End If
    Next val

    ' Output the Most Recent file from the Directory
    
    fx_Get_Most_Recent_File_From_Directory_Based_On_File_Name = dict_Files(dict_FileDates(dtCurrFile))

End Function
Function fx_Get_Most_Recent_File_From_Directory_Based_On_Modified_Date(strFolderPath As String, strFileExtension As String, Optional strFileName As String) As String

' Purpose: To loop through a directory and return the most recent file (based on Modified Date) and the path to that file.
' Trigger: Called
' Updated: 1/31/2022

' Use Example: _
        strMostRecentFilePath = fx_Get_Most_Recent_File_From_Directory_Based_On_Modified_Date( _
                                strFolderPath:="C:\U Drive\Analytics Requests\DA.21.049 - Soto (Sterling Concentration Analysis)\", _
                                strFileExtension:=".vsd")
                                
' Change Log:
'       9/10/2021:  Initial Creation, based on fx_Get_Most_Recent_File_From_Directory
'       9/10/2021:  Updated to be more dynamic with the file extension
'       1/31/2022:  Added the Like * to be more dynamic for the file extentions
'                   Added the code for strFileName to search for a specfic file

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    ' Declare Dates
    
    Dim dtCurrFile As Date
    
    ' Declare Dictionaries

    Dim dict_Files As Scripting.Dictionary
    Set dict_Files = fx_Create_File_Dictionary(strFolderPath:=strFolderPath)
        
    Dim dict_FileDates As Scripting.Dictionary
    Set dict_FileDates = New Scripting.Dictionary

    ' Declare Loop Variables
    
    Dim val As Variant
    
    ' Declare File Variables
    
    Dim strCurrFileExt As String
    
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim objFile As Object
    
    Dim strModifiedDate As String

' ---------------------------------------
' Determine the most recent Weekly Review
' ---------------------------------------

    ' Create the Dictionary of File Dates
    
    For Each val In dict_Files.Keys
        strCurrFileExt = Right(val, Len(val) - InStrRev(val, ".") + 1)
        
        If strCurrFileExt Like strFileExtension & "*" Then
            If strFileName = "" Then
                
                    ' Set File Properties to get to the DateLastModified
                    Set objFile = objFSO.GetFile(strFolderPath & val)
                        strModifiedDate = objFile.DateLastModified
                    
                    dict_FileDates.Add key:=CDate(strModifiedDate), Item:=val
            Else
                If val Like strFileName Then
                    ' Set File Properties to get to the DateLastModified
                    Set objFile = objFSO.GetFile(strFolderPath & val)
                        strModifiedDate = objFile.DateLastModified
                    
                    dict_FileDates.Add key:=CDate(strModifiedDate), Item:=val
                End If
            End If
        End If
    Next val
    
    ' Find the most recent date

    For Each val In dict_FileDates.Keys
        If dtCurrFile < val Then
            dtCurrFile = val
        End If
    Next val

    ' Output the Most Recent file from the Directory
    
    fx_Get_Most_Recent_File_From_Directory_Based_On_Modified_Date = dict_Files(dict_FileDates(dtCurrFile))

End Function
Public Function fx_Output_Project_Number(strProject As String) As String

' Purpose: This function will output the Project number for the passed project.
' Trigger: Called
' Updated: 10/4/2021

' Change Log:
'       10/4/2021: Initial Creation

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim strProjLeftofDash As String
        strProjLeftofDash = Left(strProject, InStr(1, strProject, " - "))
    
    Dim strProjectNum As String
        strProjectNum = Trim(Right(strProjLeftofDash, Len(strProjLeftofDash) - InStr(1, strProject, ".")))

' --------------
' Pass the value
' --------------

    fx_Output_Project_Number = strProjectNum

End Function
Function fx_Add_Task_To_Daily_ToDo(intTargetRow As Integer)

' Purpose: To allow me to quickly add the selected Task into my Daily To Do
' Trigger: Doublclicking on a RefNum in ws_Tasks
' Updated: 11/29/2022

' Change Log:
'       4/28/2022:  Initial Creation
'       4/29/2022:  Updated to a function and to point to the ws_Tasks to pull in the Project and Task
'                   Updated to pull in the Notes field and use an array to handle multiple notes
'       10/27/2022: Updated to reflect the name of the section to "[TODAY'S WORK]"
'       11/29/2022: Added the code to Open / Save the text file before writing to it, resolving the overwrite issue
'                   Added the code to determine if the project is already in the To Do Daily, if so adds under that project header

' ***********************************************************************************************************************************

'Save the text file before opening it -> Added on 11/29/22
    AppActivate ("Notepad++")
    Application.SendKeys ("^s")

' -----------------------------
' Declare Daily To Do Variables
' -----------------------------

    ' Folder Location
    Dim strFolderLoc As String
    #If Personal <> 1 Then
        strFolderLoc = "C:\U Drive\Support\Daily To Do\"
    #Else
        strFolderLoc = "D:\D Documents\Support\Daily To Do\"
    #End If

    ' Today's Daily To Do
    Dim strTodayDateDay As String
        strTodayDateDay = Format(Date, "DDDD")
        
    Dim strTodayDate As String
        strTodayDate = Format(Date, "yyyy.mm.dd")
    
    Dim strTodayFileLoc As String
        strTodayFileLoc = strFolderLoc & strTodayDate & " - " & strTodayDateDay & " - Daily To Do.properties"

' -----------------
' Declare Variables
' -----------------

    ' Text File Variables

    Dim intTxtFile As Long
        intTxtFile = FreeFile
    
    Dim FileContent As String
    
    Dim strCurrentWorkSection As String
        strCurrentWorkSection = "[TODAY'S WORK]"
    
    Dim intSec1 As Long
    
    Dim intSec2 As Long
    
    ' Task Variables
    
    Dim strSelectedProject As String
        strSelectedProject = ThisWorkbook.Sheets("Tasks").Cells(intTargetRow, "H").Value2
        
    Dim strSelectedTask As String
        strSelectedTask = ThisWorkbook.Sheets("Tasks").Cells(intTargetRow, "J").Value2
        
    Dim strSelectedNotes As String
        strSelectedNotes = Replace(ThisWorkbook.Sheets("Tasks").Cells(intTargetRow, "O").Value2, vbTab, "")

    Dim strSelectedNotes_Updated As String

' -----------------------------
' Create the New Content String
' -----------------------------

    Dim arry_Notes() As String
        arry_Notes = Split(strSelectedNotes, vbLf) ' Split the string into individual values based on the line breaks
    
    Dim i As Integer
    
    For i = 0 To UBound(arry_Notes)
        strSelectedNotes_Updated = strSelectedNotes_Updated & String(3, vbTab) & Replace("> " & arry_Notes(i), "> >", ">") & Chr(10)
    Next i
    
    If Right(strSelectedNotes_Updated, 1) = vbLf Then
        strSelectedNotes_Updated = Left(strSelectedNotes_Updated, Len(strSelectedNotes_Updated) - 1)
    End If
        
    Dim strNewContent As String
    If strSelectedNotes <> "" Then
        strNewContent = String(1, vbTab) & "♦ " & strSelectedProject & Chr(10) & _
                        String(2, vbTab) & "• " & strSelectedTask & Chr(10) & _
                        strSelectedNotes_Updated & Chr(10)
    Else
        strNewContent = String(1, vbTab) & "♦ " & strSelectedProject & Chr(10) & _
                        String(2, vbTab) & "• " & strSelectedTask & Chr(10)
    End If
                    
' --------------------------------------------------
' Add the new content under the Current Work section
' --------------------------------------------------

    'Open the text file in Read Only mode to pull the current content
        Open strTodayFileLoc For Input As intTxtFile
            FileContent = Input(LOF(intTxtFile), intTxtFile)
        Close intTxtFile
        
    'Assign the string variables, based on if there is an existing project or it's new
        
        If InStr(1, FileContent, strSelectedProject, vbTextCompare) > 0 Then
            intSec1 = InStr(1, FileContent, strSelectedProject, vbTextCompare)
                intSec1 = intSec1 - 6
            
            intSec2 = Len(FileContent) - intSec1 - 6 - Len(strSelectedProject)
            
            FileContent = Left(FileContent, intSec1) & strNewContent & Right(FileContent, intSec2)
        Else
            intSec1 = InStr(1, FileContent, strCurrentWorkSection, vbTextCompare) + Len(strCurrentWorkSection) - 1
                intSec1 = intSec1 + Len("--------------") + 4
            
            intSec2 = Len(FileContent) - intSec1
            
            FileContent = Left(FileContent, intSec1) & vbCrLf & strNewContent & Right(FileContent, intSec2)
        End If
      
    'Open the text file in a Write mode to add the new content
        Open strTodayFileLoc For Output As intTxtFile
            Print #intTxtFile, FileContent
        Close intTxtFile
        
    AppActivate ("Excel") ' Switch back to Excel

End Function
Public Function fx_Reset_Heading_Visibility_in_ToDo()

' Purpose: To loop through the worksheets in my To Do and disable the headings.
' Trigger: Called
' Updated: 7/31/2022

' Use Example: _
    Call fx_Reset_Heading_Visibility_in_ToDo

' Change Log:
'       7/31/2022:  Initial Creation

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim ws As Worksheet

' ------------------------
' Loop through the folders
' ------------------------

    For Each ws In ThisWorkbook.Sheets
    
        ws.Activate
    
        Select Case ws.Name
            Case "Current"
                ActiveWindow.DisplayHeadings = False
            Case "Projects"
                ActiveWindow.DisplayHeadings = False
            Case "Tasks"
                ActiveWindow.DisplayHeadings = False
            Case "Waiting"
                ActiveWindow.DisplayHeadings = False
            Case "Questions"
                ActiveWindow.DisplayHeadings = False
            Case "Recurring"
                ActiveWindow.DisplayHeadings = False
            Case "Temp"
                ActiveWindow.DisplayHeadings = True
            Case "Daily"
                ActiveWindow.DisplayHeadings = True
            Case "Lists"
                ActiveWindow.DisplayHeadings = True
        End Select
    
    Next ws
    
End Function
Public Function fx_Name_TextExpander(strShortName As String) As String

' Purpose: To expand someone's name based on a shorthand.
' Trigger: Called
' Updated: 12/8/2023

' Use Example: _
    Call fx_Name_TextExpander(.Cells(int_CurRow, col_WaitingFor).Value)

' Change Log:
'       12/3/2022:  Initial Creation
'       2/27/2023:  Added Lindsey Desai
'       8/29/2023:  Updated to reflect people from my new role and purge those from my old
'       12/8/2023:  Added Selasi

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim intShortNameLength As Long
    
    Dim strLongName As String

' --------------------------------
' Loop through each of the options
' --------------------------------

    If UCase(Left(strShortName, Len("Sam"))) = "SAM" Then
        strLongName = "Sam Yang"
    ElseIf UCase(Left(strShortName, Len("Heather"))) = "HEATHER" Then
        strLongName = "Heather Laberinto"
    ElseIf UCase(Left(strShortName, Len("Heather"))) = "STEPHEN" Then
        strLongName = "Stephen Freyermuth"
    ElseIf UCase(Left(strShortName, Len("Laura"))) = "LAURA" Then
        strLongName = "Laura (Deloitte)"
    ElseIf UCase(Left(strShortName, Len("Nik"))) = "NIK" Then
        strLongName = "Nik Corbaxhi (SOXPO)"
    ElseIf UCase(Left(strShortName, Len("Nik K"))) = "NIK K" Then
        strLongName = "Nik Corbaxhi (SOXPO)"
    ElseIf UCase(Left(strShortName, Len("Selasi"))) = "SELASI" Then
        strLongName = "Selasi Kumekpor"
    Else
        strLongName = strShortName
    End If
                    
    ' Output the value
    fx_Name_TextExpander = strLongName

End Function


