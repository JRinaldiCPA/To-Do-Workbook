Attribute VB_Name = "myFunctions"
Option Explicit
Public Function fx_Banded_Rows(rg As Range)

' Purpose: This function takes the selected rows and creates the banded rows.
' Trigger: Called Function
' Updated: 9/16/2019

' ***********************************************************************************************************************************

Dim x As Long
    
For x = rg.Row To rg.Rows.count
    If x Mod 2 = 1 Then
        rg.Rows(x).Interior.Color = RGB(240, 240, 240)
    End If
Next x

    rg.Borders(xlEdgeBottom).Color = RGB(190, 190, 190)
    rg.Borders(xlInsideHorizontal).Color = RGB(190, 190, 190)
    rg.Borders(xlEdgeTop).Color = RGB(190, 190, 190)

End Function
Public Function fx_Error_Handler(errNum, errSource, errDesc)

' Purpose: This function uses the given error to output the error details in more elegant way.
' Trigger: Called Function
' Updated: 9/16/2019

'XXX - XXX I should set this up so that I have a boolean option for Simple or Complex, simple is what I have now, Complex trys to output a specific error message AND solution

' ***********************************************************************************************************************************

fx_Error_Handler = "Uh oh, something went awry." & Chr(10) & Chr(10) _
    & "You should talk to Axcel or James, but before you do take a screenshot of the following error details:" & Chr(10) & Chr(10) _
    & "    Error Source: " & errSource & Chr(10) _
    & "    Error Number: " & errNum & Chr(10) _
    & "    Error Desc. : " & errDesc

Call myPrivateMacros.DisableForEfficiencyOff

'IF Complex

If errNum = "91" Then
    MsgBox ("Your error is the result of an Object not being set." & Chr(10) & Chr(10) & _
        "This typically happens becuase you didn't qualify a range, check for any ws.range(cells(),cells()) situations")
End If

End Function
Public Function fx_Create_Folder(strFullPath As String)

' Purpose: This function will create a folder if it doesn't already exist.
' Trigger: Called
' Updated: 9/18/2019

' ***********************************************************************************************************************************

    If Dir(strFullPath, vbDirectory) = vbNullString Then MkDir (strFullPath)
    
    Dim objFSO As Object
    
    fx_Create_Folder = objFSO.GetFolder(strFullPath)

End Function
Public Function fx_Reverse_Name(strFullName As String)

' Purpose: This function reverses a users name that was LName, FName to be FName LName, (Ex. Rinaldi, Ethan -> Ethan Rinaldi)
' Trigger: Called
' Updated: 10/11/2019

' ***********************************************************************************************************************************

Dim strFirstName As String
    strFirstName = Right(strFullName, Len(strFullName) - InStrRev(strFullName, ",") - 1)

Dim strLastName As String
    strLastName = Left(strFullName, InStr(1, CStr(strFullName), ",") - 1)

fx_Reverse_Name = strFirstName & " " & strLastName

End Function
Function fx_Create_Headers(str_Target_FieldName As String, arry_Target_Header As Variant)

' Purpose: To determine the column number for a specific title in the header.
' Trigger: Called
' Updated: 12/11/2020

' Change Log:
'       5/1/2020: Intial Creation
'       12/11/2020: Updated to use an array instead of the range, reducing the time to run by 75%.

' --------------------------------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
    intColNum_Source = fx_Create_Headers(strFieldName, arry_Header_Source)

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim i As Long

' --------------------------------------------------
' Loop through the array to find the matching column
' --------------------------------------------------

    For i = LBound(arry_Target_Header) To UBound(arry_Target_Header)
        If arry_Target_Header(i, 1) = str_Target_FieldName Then
            fx_Create_Headers = i
            Exit Function
        End If
    Next i

End Function
Function fx_Sheet_Exists(strWBName As String, strWsName As String) As Boolean

' Purpose: To determine if a sheet exists, to be used in an IF statement.
' Trigger: Called
' Updated: 6/29/2021

' Change Log:
'       6/29/2021: Intial Creation
'       6/29/2021: Added the ErrorHandler for the 2015 #VALUE error when the ws doesn't exist

' Use Example: _
'    Dim bolSupportOpenAlready As Boolean
'    bolSupportOpenAlready = fx_Sheet_Exists( _
        strWbName:=strProjName & ".xlsx", _
        strWsName:="Next Actions")
'   If fx_Sheet_Exists(ThisWorkbook.Name, "VALIDATION") = False Then

' ***********************************************************************************************************************************

On Error GoTo ErrorHandler

    If Evaluate("ISREF('[" & strWBName & "]" & strWsName & "'!A1)") = True Then fx_Sheet_Exists = True
        Exit Function

ErrorHandler:

fx_Sheet_Exists = False

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
Public Function fx_File_Exists(strFullPath As String) As Boolean

' Purpose: This function will determine if a file exists already.
' Trigger: Called
' Updated: 7/18/2022

' Use Example: _
    'bolTodayFileExists = fx_File_Exists(strTodayFileLoc)

' Change Log:
'       8/19/2021:  Initial Creation
'       7/18/2022:  Added the Use Example

' **************************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    'Dim Objects
    Dim objFSO As Object
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        
' ---------------------------
' Check if passed file exists
' ---------------------------
    
    If objFSO.FileExists(strFullPath) = True Then fx_File_Exists = True

    'Release the Object
    Set objFSO = Nothing

End Function
Function fx_Steal_First_Row_Formating(ws As Worksheet, Optional intFirstRow As Long, Optional int_LastCol As Long, Optional int_LastRow As Long, Optional intSingleRow As Long)

' Purpose: To copy the formatting from the first row of data and apply to the rest of the data.
' Trigger: Called
' Updated: 2/3/2022

' Use Example: _
    Call fx_Steal_First_Row_Formating( _
        ws:=wsQCReview, _
        intFirstRow:=2, _
        int_LastRow:=int_LastRow, _
        int_LastCol:=int_LastCol)

' Use Example 2: Call fx_Steal_First_Row_Formating(ws:=wsQCReview)

' Change Log:
'       5/17/2021:  Intial Creation
'       6/16/2021:  Added the 'Application.Goto' to reset the copy paste
'       12/6/2021:  Added the option to pass only a single row
'       12/8/2021:  Added the rngCur so the screen doesn't jump around
'       1/8/2022:   Updated some of the passed variables to be optional
'                   Defaulted intFirstRow to be 2 if not passed
'       2/3/2022:   Updated to keep whatever was in the clipboard before running this

' ***********************************************************************************************************************************
  
' -----------------
' Declare Variables
' -----------------

    ' Declare Current Data
    Dim rngCur As Range
    Set rngCur = ActiveCell

    Dim strCurrentClipboardContents As String
        strCurrentClipboardContents = fx_Copy_from_Clipboard

    ' Declare Integers
    
    If intFirstRow = 0 Then intFirstRow = 2
    
    If int_LastRow = 0 Then
       int_LastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    End If
    
    If int_LastCol = 0 Then
       int_LastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    End If
    
    ' Declare Ranges
    
With ws

    Dim rngFormat As Range
        Set rngFormat = .Range(.Cells(intFirstRow, 1), .Cells(intFirstRow, int_LastCol))
    
    Dim rngTarget As Range
        If intSingleRow <> 0 Then
            Set rngTarget = .Range(.Cells(intSingleRow, 1), .Cells(intSingleRow, int_LastCol))
            'Set rngTarget = .Cells(intSingleRow, 1)
        ElseIf int_LastRow <> 0 Then
            Set rngTarget = .Range(.Cells(intFirstRow + 1, 1), .Cells(int_LastRow, int_LastCol))
        Else
            MsgBox "There was no row passed to the Steal First Row function."
        End If

End With

' ---------------------------------------------------------------------------------------------------
' Copy the formatting from the first row of data (intFirstRow) to the remaining rows (thru int_LastRow)
' ---------------------------------------------------------------------------------------------------
    
    rngFormat.Copy: rngTarget.PasteSpecial xlPasteFormats
    
    ' Go back to where you were before the code
    Application.CutCopyMode = False
    
    Call fx_Copy_to_Clipboard(strTextToCopy:=strCurrentClipboardContents)
    
    Application.GoTo Reference:=rngCur, Scroll:=False
    
End Function
Function fx_Find_CurRow(ws As Worksheet, strTargetFieldName As String, strTarget As String) As Long

' Purpose: To find the target value in the passed column for the passed worksheet.  Replaces the Find function, to account for hidden rows.
' Trigger: Called
' Updated: 12/27/2021

' Use Example: _
    Call fx_Find_CurRow( _
        ws:=ThisWorkbook.Sheets("Projects"), _
        strTargetFieldName:="Project", _
        strTarget:="P.343 - Migrate to Win10")
'
'   intRowCurProject = fx_Find_CurRow(ws:=ws_Projects, strTargetFieldName:="Project", strTarget:=strProjName)

' Change Log:
'       12/26/2021: Initial Creation
'       12/27/2021: Made the int_LastRow more dynamic, and added the 1 to capture a blank row

' ***********************************************************************************************************************************
  
' -----------------
' Declare Variables
' -----------------

    ' Declare Header Variables
    
    Dim arry_Header_Data() As Variant
        arry_Header_Data = Application.Transpose(ws.Range(ws.Cells(1, 1), ws.Cells(1, 99)))
        
    Dim col_Target As Long
        col_Target = fx_Create_Headers(strTargetFieldName, arry_Header_Data)
    
    ' Declare Other Variables

    Dim int_LastRow As Long
        int_LastRow = WorksheetFunction.Max( _
        ws.Cells(Rows.count, col_Target).End(xlUp).Row, _
        ws.UsedRange.Rows(ws.UsedRange.Rows.count).Row) + 1
        
    Dim arryData() As Variant
        arryData = ws.Range(ws.Cells(1, col_Target), ws.Cells(int_LastRow, col_Target))

    Dim dictData As New Scripting.Dictionary
        dictData.CompareMode = TextCompare
        
    Dim i As Long
        
' -------------------
' Fill the Dictionary
' -------------------
    
    For i = 1 To UBound(arryData)
    On Error Resume Next
        dictData.Add key:=arryData(i, 1), Item:=i
    On Error GoTo 0
    Next i
    
' --------------------
' Find the Current Row
' --------------------
    
    fx_Find_CurRow = dictData(strTarget)

End Function
Function fx_Find_Row(ws As Worksheet, str_Target As String, Optional str_TargetFieldName As String, Optional str_TargetCol As String) As Long

' Purpose: To find the target value in the passed column for the passed worksheet.  Replaces the Find function, to account for hidden rows.
' Trigger: Called
' Updated: 3/6/2022

' Change Log:
'       12/26/2021: Initial Creation
'       12/27/2021: Made the int_LastRow more dynamic, and added the 1 to capture a blank row
'       1/19/2022:  Added the code to allow str_TargetCol to be passed
'       3/6/2022:   Added Error Handling around the Dictionary to allow duplicates
'                   Updated to handle situations where str_TargetCol AND str_TargetFieldName are not passed

' Note: Formerly called fx_Find_CurRow

' --------------------------------------------------------------------------------------------------------------------------------------------------------

' Use Example: _
    Call fx_Find_Row( _
        ws:=ThisWorkbook.Sheets("Projects"), _
        str_TargetFieldName:="Project", _
        str_Target:="P.343 - Migrate to Win10")

' Use Example 2: Passing the Target Field Name _
    intRowCurProject = fx_Find_Row(ws:=wsProjects, str_Target:=strProjName, str_TargetFieldName:="Project")

' Use Example 3: Passing the Target Column letter reference _
    intRowCurProject = fx_Find_Row(ws:=wsProjects, str_Target:=strProjName, str_TargetCol:="B")

' ***********************************************************************************************************************************
  
' -----------------
' Declare Variables
' -----------------

    'Declare Header Variables
    
    Dim arry_Header_Data() As Variant
        arry_Header_Data = Application.Transpose(ws.Range(ws.Cells(1, 1), ws.Cells(1, 99)))
        
    Dim col_Target As Long
        If str_TargetCol <> "" Then
            col_Target = ws.Range(str_TargetCol & "1").Column
        ElseIf str_TargetFieldName <> "" Then
            col_Target = fx_Create_Headers(str_TargetFieldName, arry_Header_Data)
        Else
            col_Target = 1
        End If
    
    'Declare Other Variables

    Dim int_LastRow As Long
        int_LastRow = WorksheetFunction.Max( _
        ws.Cells(ws.Rows.count, col_Target).End(xlUp).Row, _
        ws.UsedRange.Rows(ws.UsedRange.Rows.count).Row) + 1
        
    Dim arryData() As Variant
        arryData = ws.Range(ws.Cells(1, col_Target), ws.Cells(int_LastRow, col_Target))

    Dim dictData As New Scripting.Dictionary
        dictData.CompareMode = TextCompare
        
    Dim i As Long
        
' -------------------
' Fill the Dictionary
' -------------------
    
On Error Resume Next
    
    For i = 1 To UBound(arryData)
        dictData.Add key:=arryData(i, 1), Item:=i
    Next i
    
On Error GoTo 0
    
' --------------------
' Find the Current Row
' --------------------
    
    fx_Find_Row = dictData(str_Target)

End Function

Public Function fx_List_Files_In_Folder(strDirParentName As String, Optional bolPrintResults As Boolean) As Dictionary

' Purpose: To loop through the passed folder to add all of the files to a collection.
' Trigger: Called
' Updated: 5/16/2022

' Use Example: _
    Set dict_Files = fx_List_Files_In_Folder("C:\U Drive\Analytics Requests\DA.22.055 - Scorecard Tolerance Breach Tracker")

' Change Log:
'       2/6/2022:   Initial Creation
'       2/13/2022:  Update the code around the Archive folder and looping
'       5/16/2022:  Added the bolPrintResults to allow the user to print the results if they want

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim objDirParent As Object
    Set objDirParent = objFSO.GetFolder(strDirParentName)
    
    Dim objSubFolder As Object
    
    Dim objFile As Object
    
    Dim dictFiles As New Dictionary

' ------------------------
' Loop through the folders
' ------------------------

    For Each objFile In objDirParent.Files
        dictFiles.Add key:=objFile.Name, Item:=objFile.path
    Next objFile
    
    For Each objSubFolder In objDirParent.SubFolders
        If Right(objSubFolder, 9) <> "(ARCHIVE)" Then
            fx_List_Files_In_Folder (objSubFolder.path)
        End If
    Next objSubFolder
    
    Set fx_List_Files_In_Folder = dictFiles
    
' -----------------
' Print the results
' -----------------
    
    If bolPrintResults = True Then
    
        Dim key As Variant
        For Each key In dictFiles.Keys
            Debug.Print key, dictFiles(key)
        Next key
    
    End If
       
' ---------------
' Clear Variables
' ---------------
    
    Set objFile = Nothing
    Set objDirParent = Nothing
    Set objFSO = Nothing
    
End Function
Function fx_Create_Dynamic_Lookup_List(wsDataSource As Worksheet, str_Dynamic_Lookup_Value As String, col_Dynamic_Lookup_Field As Long, Optional col_Criteria_Field As Long, Optional str_Criteria_Match_Value As Variant, Optional col_Target_Field As Long) As Variant

' Purpose: To create the dynamic list of values to be used in the ListBox, based on a change to the cmb_Dynamic_Borrower_Lookup.
' Trigger: Start typing in the Dynamic_Borrower_Lookup combo box (cmb_Dynamic_Borrower_Lookup_Change)
' Updated: 10/7/2022

' Change Log:
'       11/21/2021: Intial Creation for the PAR Agenda, taken from Sageworks Validation Dashboard code
'       1/18/2022:  Updated and converted to a function
'       10/6/2022:  Updated the naming of the fields, and handled the error if there was no LOB Lookup value
'       10/7/2022:  Updated to allow a seperate Dynamic Lookup and Target field

' ********************************************************************************************************************************************************

' Use Example: _
'    Dim arryBorrowersTemp As Variant
'       arryBorrowersTemp = myFunctions.fx_Create_Dynamic_Lookup_List( _
        wsDataSource:=wsData, _
        col_Dynamic_Lookup_Field:=col_Borrower, _
        str_Dynamic_Lookup_Value:=Me.cmb_Dynamic_Borrower.Value, _
        col_Criteria_Field:=col_LOBUpdated, _
        str_Criteria_Match_Value:=Me.lst_LOB.Value)

'    Me.lst_Borrowers.List = arryTargetValuesTemp

' ********************************************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    'Declare Strings
    Dim strLookupValue As String
    Dim strCriteriaValue As String
    Dim strTargetValue As String
    
    'Declare Cell References
    Dim intSourceRow As Long: intSourceRow = 2
    Dim intArryRow As Long: intArryRow = 0
    
    'Declare Arrays
    Dim arryTargetValues As Variant
        ReDim arryTargetValues(1 To 99999)

    If IsMissing(str_Criteria_Match_Value) Then str_Criteria_Match_Value = ""

' -------------------------------------
' Add the borrowers to the lookup Array
' -------------------------------------
       
    With wsDataSource
            
        Do While .Cells(intSourceRow, col_Dynamic_Lookup_Field).Value2 <> ""
            
            ' Set the loop variables
            strLookupValue = .Cells(intSourceRow, col_Dynamic_Lookup_Field).Value2
            If col_Criteria_Field <> 0 Then strCriteriaValue = .Cells(intSourceRow, col_Criteria_Field).Value2
            If col_Target_Field <> 0 Then strTargetValue = .Cells(intSourceRow, col_Target_Field).Value2
                
                'If the data matches add to the array
                If InStr(1, strLookupValue, str_Dynamic_Lookup_Value, vbTextCompare) Then
                    If col_Criteria_Field = 0 Or str_Criteria_Match_Value = strCriteriaValue Then
                        intArryRow = intArryRow + 1
                        If strTargetValue <> "" Then arryTargetValues(intArryRow) = strTargetValue Else arryTargetValues(intArryRow) = strLookupValue
                    End If
                End If
            
            intSourceRow = intSourceRow + 1
        Loop
    End With

    If intArryRow > 0 Then ' If nothing was passed, don't redim
        ReDim Preserve arryTargetValues(1 To intArryRow)
    End If
    
    'Output the results
    fx_Create_Dynamic_Lookup_List = arryTargetValues

End Function
Function fx_Find_LastColumn(ws_Target As Worksheet, Optional bolIncludeSpecialCells As Boolean, Optional bolIncludeUsedRange As Boolean) As Long

' Purpose: To output the the Last Column for the passed ws using multiple options.
' Trigger: Called
' Updated: 11/18/2023
'
' Change Log:
'       3/6/2022:   Initial Creation, based on fx_Find_LastCol
'       11/18/2023: Fixed an unclosed error handler (missing On Error GoTo 0)
'                   Updated to simplify the code and remove an If statement
'
' -----------------------------------------------------------------------------------------------------------------------------------
'
' Use Example: _
'   int_LastCol = fx_Find_LastColumn(wsData)
'
' Legend:
'
'   bolIncludeUsedRange: If this is True then the last Col of the UsedRange will be included in the Max formula
'   bolIncludeSpecialCells: If this is True then the SpecialCells(xlCellTypeLastCell) col will be included in the Max formula
'
' ***********************************************************************************************************************************
  
' -----------------
' Declare Variables
' -----------------

On Error Resume Next

    'Declare the 1st Option based on End(xlToLeft)
    
    Dim int_LastCol_1st As Long
        int_LastCol_1st = ws_Target.Cells(1, ws_Target.Columns.count).End(xlToLeft).Column
        
    'Declare the 2nd Option based on xlCellTypeLastCell
    
    If bolIncludeSpecialCells = True Then
        Dim int_LastCol_2nd As Long
            int_LastCol_2nd = ws_Target.Cells(1, 1).SpecialCells(xlCellTypeLastCell).Column
    End If
    
    'Declare the 3rd Option based on UsedRange.Rows.Count
    
    If bolIncludeUsedRange = True Then
        Dim int_LastCol_3rd As Long
            int_LastCol_3rd = ws_Target.UsedRange.Columns(ws_Target.UsedRange.Columns.count).Column
    End If

    'Declare the Max Integer
    Dim int_LastCol_Max As Long

On Error GoTo 0

' ---------------------------------
' Determine which int_LastCol to use
' ---------------------------------

    int_LastCol_Max = WorksheetFunction.Max(int_LastCol_1st, int_LastCol_2nd, int_LastCol_3rd, 2) 'Don't pass values <2
        
    fx_Find_LastColumn = int_LastCol_Max
        
End Function

Function fx_Find_LastRow(ws_Target As Worksheet, _
Optional int_TargetColumn As Long, Optional bol_MinValue2 As Boolean, Optional bolIncludeSpecialCells As Boolean, Optional bolIncludeUsedRange As Boolean) As Long

' Purpose: To output the the Last Row for the passed ws using multiple options.
' Trigger: Called
' Updated: 11/19/2023
'
' Change Log:
'       11/29/2021: Initial Creation
'       3/6/2022:   Overhauled to include error handling, and the if statements to breakout the determination of the Last Row
'                   Added the fx_Find_Row code as an alternative to handle filtered data
'       11/18/2023: Fixed an unclosed error handler (missing On Error GoTo 0)
'                   Updated to simplify the code and remove an If statement
'       11/19/2023: Added the code for bol_MinValue2 to allow the option to not pass a # less than 2
'
' -----------------------------------------------------------------------------------------------------------------------------------
'
' Use Example: _
'   int_LastRow = fx_Find_LastRow(wsData)
'
' Use Example 2: Using all of the optional variables _
'   int_LastRow = fx_Find_LastRow(ws_Target:=wsTest, int_TargetColumn:=2, bolIncludeSpecialCells:=True, bolIncludeUsedRange:=True)
'
' Legend:
'   bolIncludeUsedRange: If this is True then the last row of the UsedRange will be included in the Max formula
'   bolIncludeSpecialCells: If this is True then the SpecialCells(xlCellTypeLastCell) row will be included in the Max formula
'
' ***********************************************************************************************************************************
  
' -----------------
' Declare Variables
' -----------------

On Error Resume Next
    
    ' Declare the 1st Option based on End(xlUp)
    
    Dim int_LastRow_1st As Long
    
    If int_TargetColumn <> 0 Then
        int_LastRow_1st = ws_Target.Cells(ws_Target.Rows.count, int_TargetColumn).End(xlUp).Row
    Else
        int_LastRow_1st = ws_Target.Cells(ws_Target.Rows.count, "A").End(xlUp).Row
    End If
    
    ' Declare the 2nd Option based on xlCellTypeLastCell
    
    If bolIncludeSpecialCells = True Then
        Dim int_LastRow_2nd As Long
            int_LastRow_2nd = ws_Target.Cells(1, 1).SpecialCells(xlCellTypeLastCell).Row
    End If
    
    ' Declare the 3rd Option based on UsedRange.Rows.Count
    
    If bolIncludeUsedRange = True Then
        Dim int_LastRow_3rd As Long
            int_LastRow_3rd = ws_Target.UsedRange.Rows(ws_Target.UsedRange.Rows.count).Row
    End If
    
    ' Declare the 4th Option of a minimum of 2
    
    If bol_MinValue2 = True Then
        Dim int_LastRow_4th As Long
            int_LastRow_4th = 2
    End If
    
    ' Declare the Max
    
    Dim int_LastRow_Max As Long

On Error GoTo 0

' ---------------------------------
' Determine which int_LastRow to use
' ---------------------------------

    int_LastRow_Max = WorksheetFunction.Max(int_LastRow_1st, int_LastRow_2nd, int_LastRow_3rd, int_LastRow_4th)
        
    fx_Find_LastRow = int_LastRow_Max
        
End Function

Public Function fx_Return_Column_Letter(intColumnNum As Long) As String

' Purpose: This function will return the letter of the passed column.
' Trigger: Called
' Updated: 9/16/2021

' Change Log:
'       9/16/2021: Initial Creation

' ***********************************************************************************************************************************
    
    fx_Return_Column_Letter = Split(Cells(1, lngColumn).Address, "$")(1)

End Function
Function fx_Close_Workbook(objWorkbookToClose As Workbook, _
    Optional bol_KeepWorkbookOpen As Boolean, Optional bol_SaveWorkbook As Boolean, Optional bol_ConvertToXLSX As Boolean, _
    Optional strAddToWorkbookName As String, Optional strNewFilePath As String)
             
' Purpose: This function will close the passed workbook.
' Trigger: Called Function
' Updated: 11/20/2023
'
' Change Log:
'       11/19/2023: Initial Creation, based on Sageworks Validation Dashboard > o_51_Create_a_XLSX_Copy
'                   Added the conflict resolution if the file already exists
'       11/20/2023: Added the strNewFilePath variable and related code
'                   Added the bolConvertToXlsx option
'
' --------------------------------------------------------------------------------------------------------------------------------------------------------
'
' Use Example:
'    Call fx_Close_Workbook( _
'       objWorkbookToClose:=Workbook("(WINDOWS) TEST BLANKS.xlsx"), _
'       bol_SaveWorkbook:=True, _
'       strAddToWorkbookName:="COMBINED")
'
' Legend:
'   strWorkbookToClose:   The workbook to be closed
'   bol_KeepWorkbookOpen: Determine if the workbook should be closed
'   bol_SaveWorkbook:     Should the workbook be saved
'   bol_ConvertToXLSX:    Will convert the file to an .xlsx if selected
'   strAddToWorkbookName: The text to add to the workbook's name if it is being saved
'   strNewFilePath: Allow the user to move the updated files to a new path, only used when saving the workbook
'
' ***********************************************************************************************************************************
             
' -----------------
' Declare Variables
' -----------------

    ' Declare Strings

    Dim str_FullPath As String
        str_FullPath = objWorkbookToClose.FullName
    
    Dim str_FileExtension As String 'Determine the extension for the original file
        str_FileExtension = Left(Right(str_FullPath, 5), 5)
        
        If Left(str_FileExtension, 1) <> "." Then str_FileExtension = Left(Right(str_FullPath, 4), 4)
    
    Dim str_NewFileExtension As String 'Determine the extension for the new file type
    
    If bol_ConvertToXLSX = True Then
        str_NewFileExtension = ".xlsx"
    Else
        str_NewFileExtension = str_FileExtension
    End If
        
    If strAddToWorkbookName <> "" Then
        str_FullPath = Replace(str_FullPath, str_FileExtension, " " & strAddToWorkbookName & str_FileExtension)
    End If
        
    Dim str_UpdatedFullFilePath As String
        str_UpdatedFullFilePath = Replace(str_FullPath, objWorkbookToClose.path, strNewFilePath)
        str_UpdatedFullFilePath = Replace(str_UpdatedFullFilePath, str_FileExtension, str_NewFileExtension)
        
    'Remove special characters
        
        str_UpdatedFullFilePath = Replace(str_UpdatedFullFilePath, "[", "(")
        str_UpdatedFullFilePath = Replace(str_UpdatedFullFilePath, "]", ")")
        
' -----------------
' Save the Workbook
' -----------------

    If bol_SaveWorkbook = True Then
        If strNewFilePath = "" Then
            objWorkbookToClose.SaveAs Filename:=str_FullPath, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
        Else
            If bol_ConvertToXLSX = True Then
                objWorkbookToClose.SaveAs Filename:=str_UpdatedFullFilePath, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges, FileFormat:=xlOpenXMLWorkbook
            Else
                objWorkbookToClose.SaveAs Filename:=str_UpdatedFullFilePath, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
            End If
        End If
    End If

' ------------------------------------------
' Save the Workbook (or not, that's cool to)
' ------------------------------------------
    
    If bol_KeepWorkbookOpen = False Then objWorkbookToClose.Close SaveChanges:=False  'Close the Workbook

End Function

