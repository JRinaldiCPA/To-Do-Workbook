VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_Search_Folder 
   Caption         =   "  --- Dynamic Folder Search ---"
   ClientHeight    =   6552
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   10260
   OleObjectBlob   =   "uf_Search_Folder.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "uf_Search_Folder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str_PathParent As String
Dim str_FolderOpeningType As String

Dim dict_Folders As Scripting.Dictionary

Dim bol_ReverseSortOrder As Boolean

Option Explicit
Private Sub UserForm_Initialize()
    
' Purpose: To allow me to quickly pull up any folder from my U Drive.
' Trigger: Selected from the uf_Project_Selector UserForm
' Updated: 8/18/2023

' Change Log:
'       9/30/2020:  Refreshed to use a Dictionary instead of an Array.
'       5/9/2023:   Updated to remove part of the file path to shorten the output.
'       8/18/2023:  Updated to account for the different lengths of the starting path for Personal vs Profesional

' ***********************************************************************************************************************************
        
    Me.cmb_DynamicSearch.SetFocus

    With Me.lst_Folders 'Used to keep the simple name and full path
        '.ColumnCount = 2
        '.ColumnWidths = Me.lst_Folders.Width - 5 & ";" & "1"
    End With
    
    bol_ReverseSortOrder = False

Call Me.o_1_Create_Folder_List

End Sub
Public Sub cmb_DynamicSearch_Change()

' Purpose: To output the folders for the given value in the combo box.
' Updated: 10/2/2023
' Reviewd: 9/23/2023

' Change Log
'       9/30/2020:  Refreshed to use a Dictionary instead of an Array.
'       5/9/2023:   Updated to remove part of the file path to shorten the output.
'       8/18/2023:  Updated to account for the different lengths of the starting path for Personal vs Profesional
'       9/23/2023:  Removed variables no longer used
'       9/25/2023:  Converted to using a Dictionary instead of Collection
'       10/2/2023:  Added the code to allow the values to be reversed in the listbox

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------

    Dim val As Variant
    
    Dim i As Long
    
    Dim str_dictFoldersValues As Variant
        str_dictFoldersValues = dict_Folders.Items
    
' ---------------------------------------------------
' Copy the values from the dictionary to the list box
' ---------------------------------------------------
    
    Me.lst_Folders.Clear
    
    If bol_ReverseSortOrder = False Then
        
        For Each val In dict_Folders
            If InStr(1, val, cmb_DynamicSearch.Value, vbTextCompare) Then
                Me.lst_Folders.AddItem val 'If the name is similar then add to the list
            End If
        Next val
    
    ElseIf bol_ReverseSortOrder = True Then
    
        For i = UBound(str_dictFoldersValues) To LBound(str_dictFoldersValues) Step -1
            If InStr(1, str_dictFoldersValues(i), cmb_DynamicSearch.Value, vbTextCompare) Then
                Me.lst_Folders.AddItem str_dictFoldersValues(i) 'If the name is similar then add to the list
            End If
        Next i
        
    End If

End Sub
Private Sub cmd_ReverseSortOrder_Click()

' Purpose: To reverse the order of the search results as displayed in the List Box, so that the last alphabetically show up first.
' Note:    This approach helps with any folders with dates in them, so that the latest date appears first in the list.
' Updated: 10/2/2023
' Reviewd: 9/23/2023

' Change Log:
'       10/2/2023:  Original Creation

' ***********************************************************************************************************************************

    ' Update to reverse the sort order
    bol_ReverseSortOrder = Not bol_ReverseSortOrder
    
    ' Recreate the results using what is in the Dynamic Search combobox
    If Me.cmb_DynamicSearch <> "" Then Call Me.cmb_DynamicSearch_Change
    
    ' Set the focus so I can start typing
    Me.cmb_DynamicSearch.SetFocus
    
End Sub
Private Sub lst_Folders_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

' Purpose: To open the folder selected from the ListBox.
' Updated: 9/25/2023
' Reviewd: 9/25/2023

' Change Log:
'       6/2/2021:   Updated to replace Shell with FollowHyperlink
'       5/9/2023:   Updated to add in the missing part of the file path, as a result of shortening the path in the Folders list
'       8/18/2023:  Updated to account for the different lengths of the starting path for Personal vs Profesional
'       9/25/2023:  Moved the opening code to a new sub Me.o_2_Open_Selected_File

' ***********************************************************************************************************************************

    str_FolderOpeningType = "Double Click"
    
    Call Me.o_2_Open_Selected_Folder

End Sub
Private Sub lst_Folders_Enter()

' Purpose: To open the applicable RefX file when entering the list, and only one value is present.
' Trigger: Called: uf_Search_RefX
' Updated: 9/25/2023
' Reviewd: 9/25/2023

' Change Log:
'       9/5/2023:   Original Creation, taken from RefX_Search
'       9/25/2023:  Moved the opening code to a new sub Me.o_2_Open_Selected_File

' ***********************************************************************************************************************************

If Me.lst_Folders.ListCount = 1 Then
    
    str_FolderOpeningType = "On Enter List"
    
    Call Me.o_2_Open_Selected_Folder
    
End If

End Sub
Private Sub lst_Folders_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

' Purpose: To open the applicable RefX file when hitting enter on the applicable value.
' Trigger: Called: uf_Search_RefX
' Updated: 9/25/2023
' Reviewd: 9/25/2023

' Change Log:
'       9/5/2023:   Original Creation, taken from RefX_Search
'       9/25/2023:  Moved the opening code to a new sub Me.o_2_Open_Selected_File

' ***********************************************************************************************************************************

If KeyCode = vbKeyReturn And Me.lst_Folders.Value <> "" Then
    
    str_FolderOpeningType = "Press Enter"
    
    Call Me.o_2_Open_Selected_Folder
    
End If
    
End Sub
Private Sub cmd_Cancel_Click()
    
    Unload Me
    
End Sub
Sub o_1_Create_Folder_List()

' Purpose: To create the initial array that will be used to ID the folders.
' Trigger: Called: uf_Search_Folder
' Updated: 9/25/2023

' Change Log:
'       9/30/2020: Initial Creation
'       10/26/2020: Added in the split for personal vs work computers
'       8/18/2023:  Updated to account for the different lengths of the starting path for Personal vs Profesional
'       9/25/2023:  Converted to using a Dictionary instead of Collection

' ***********************************************************************************************************************************

' -----------------
' Declare Variables
' -----------------
    
    'Declare Strings
    
    #If Personal = 0 Then
        str_PathParent = "C:\U Drive\"
    #ElseIf Personal = 1 Then
        str_PathParent = "D:\D Documents\"
    #End If
    
    Dim str_CurFolder As String
    
    ' Declare Objects
    Dim objFSO As Object
        Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    Dim objDirParent As Object
        Set objDirParent = objFSO.GetFolder(str_PathParent)
    
    Dim objSubFolder As Object
    
    ' Declare Dictionaries
    Set dict_Folders = New Scripting.Dictionary
    
    ' Declare Integers
    Dim int_ParentPathLength As Long
        int_ParentPathLength = Len(str_PathParent)
    
' ------------------------------------
' Load the folders into the dictionary
' ------------------------------------

On Error Resume Next

    Do Until str_CurFolder = ""
        If (GetAttr(str_PathParent & str_CurFolder) And vbDirectory) = vbDirectory Then
            dict_Folders.Add key:=str_CurFolder, Item:=str_PathParent
        End If

        str_CurFolder = Dir()
    Loop

    'Recur through each folder

    For Each objSubFolder In objDirParent.SubFolders

    str_CurFolder = Dir(objSubFolder.path & "\", vbDirectory)

        Do Until str_CurFolder = ""
            If (GetAttr(objSubFolder.path & "\" & str_CurFolder) And vbDirectory) = vbDirectory Then
                dict_Folders.Add key:=Mid(objSubFolder.path & "\" & str_CurFolder, int_ParentPathLength), _
                                 Item:=objSubFolder.path & "\" & str_CurFolder
            End If

            str_CurFolder = Dir()
        Loop

    Next
                       
On Error GoTo 0
                       
End Sub
Sub o_2_Open_Selected_Folder()

' Purpose: To open the selected folder.
' Updated: 10/23/2023
' Reviewd: 9/25/2023

' Change Log:
'       9/25/2023:  Initial Creation
'                   Replaced Shell w/ FollowHyperlink to make opening the folder faster
'       10/13/2023: Added a wait to try to resolve the crashing  'Run-time error '-2147417848(80010108)''

' ***********************************************************************************************************************************

On Error Resume Next

' -----------------
' Declare Variables
' -----------------

    Dim str_FolderName As String
    
    If str_FolderOpeningType = "Double Click" Then
        str_FolderName = Me.lst_Folders.Value
    ElseIf str_FolderOpeningType = "Press Enter" Then
        str_FolderName = Me.lst_Folders.Value
    ElseIf str_FolderOpeningType = "On Enter List" Then
        str_FolderName = Me.lst_Folders.List(0)
    End If
                        
    Dim str_FullPath As String
        str_FullPath = dict_Folders(str_FolderName) ' Pull the full path from the dictionary
        
' ----------------------
' Open the selected file
' ----------------------
    
    Application.Wait (Now + TimeValue("0:00:005")) ' Added a wait on 10/13/23 to try to resolve the crashing
    
    ThisWorkbook.FollowHyperlink (str_FullPath)
    
    Application.Wait (Now + TimeValue("0:00:01")) ' Added a wait on 10/13/23 to try to resolve the crashing (Still having issues on 10/23/23)
    
    Unload Me
                       
End Sub
