Attribute VB_Name = "myArrays"
Dim arry_Area() As String

Dim arry_Context() As String

Dim arry_Status(1 To 5) As String
Dim arry_DA_Request_Type(1 To 5) As String

Option Explicit
Public Function GetAreaArray() As Variant
    
' Purpose: To output the Area array for my new Project and Task UserForms.
' Updated: 8/29/2023
' Reviewd: 5/20/2023

' Change Log:
'       12/16/2020: Added the conditional compiler constant to determine the file location if on my Personal computer.
'       4/14/2021:  Added Emmy into the Personal Array
'       7/27/2021:  Removed 'D/A Strategy'
'       5/7/2022:   Reduced the Personal Array to only 4, updated 'House / Yard' to be 'Household' and elminated 'Financial'
'                   Updated to use ReDim w/ a dynamic variant array
'       6/28/2023:  Updataed my Areas of Focus to include 'Recurring', 'Yard', and 'Finances'
'       8/29/2023:  Updated to include the new 'Infrastructure' and 'Strategy' options for my new role

' ***********************************************************************************************************************************
    
    #If Personal <> 1 Then
        ReDim arry_Area(1 To 7)
        arry_Area(1) = "Projects"
        arry_Area(2) = "Infrastructure"
        arry_Area(3) = "Strategy"
        arry_Area(4) = "Recurring"
        arry_Area(5) = "Continuous"
        arry_Area(6) = "Personal"
        arry_Area(7) = "D/A Requests"
    #Else
        ReDim arry_Area(1 To 6)
        arry_Area(1) = "Family"
        arry_Area(2) = "Household"
        arry_Area(3) = "Yard"
        arry_Area(4) = "Finances"
        arry_Area(5) = "Personal"
        arry_Area(6) = "Continuous"
    #End If

    GetAreaArray = arry_Area

End Function
Public Function GetContextArray() As Variant
    
' Purpose: To output the Context array for my new Project and Task UserForms.
' Updated: 4/3/2023
' Reviewd: 5/20/2023

' Change Log:
'       5/7/2022:   Initial Creation
'       4/3/2023:   Added @ EMAIL to my work Contexts

' ***********************************************************************************************************************************
        
    #If Personal <> 1 Then
        ReDim arry_Context(1 To 2)
        arry_Context(1) = "@ TASKS"
        arry_Context(2) = "@ EMAIL"
    #Else
        ReDim arry_Context(1 To 3)
        arry_Context(1) = "@ TASKS"
        arry_Context(2) = "@ YARD"
        arry_Context(3) = "@ HOUSE"
    #End If

    GetContextArray = arry_Context

End Function
Public Function GetStatusArray() As Variant
    
' Purpose: To output the Status array for new Projects.
' Updated: N/A
' Reviewd: 5/20/2023

' Change Log:
'       5/20/2023:  First created the Change Log for this Function

' ***********************************************************************************************************************************
    
    arry_Status(1) = "Active"
    arry_Status(2) = "Pending"
    arry_Status(3) = "Complete"
    arry_Status(4) = "N/A"
    arry_Status(5) = "Continuous"

    GetStatusArray = arry_Status

End Function
Public Function GetDARequestTypeArray() As Variant
    
' Purpose: To output the D/A Request Type array for my for new D/A Requests.
' Updated: N/A
' Reviewd: 5/20/2023

' Change Log:
'       5/20/2023:  First created the Change Log for this Function

' ***********************************************************************************************************************************
    
    arry_DA_Request_Type(1) = "Ad-Hoc - Quick"
    arry_DA_Request_Type(2) = "Ad-Hoc - Obtain Data"
    arry_DA_Request_Type(3) = "Ad-Hoc - Data Analysis"
    arry_DA_Request_Type(4) = "Analytics Solution"
    arry_DA_Request_Type(5) = "Advisory"
    
    GetDARequestTypeArray = arry_DA_Request_Type

End Function
