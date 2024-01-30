Attribute VB_Name = "Functions"
Option Explicit
Function Name_Reverse()

' Purpose: This function splits and reverses a user name from LAST, FIRST to FIRST LAST
' Trigger: Called
' Updated: 4/1/2020

' Change Log:
'       3/23/2020: Fixed an issue with people with middle names (ex."Elias, Richard J.")
'       4/1/2020: Fixed an issue with people with unique name formatting (First Last) that was breaking due to the missing comma.

' ****************************************************************************

On Error GoTo ErrorHandler

Dim str_User_Name As String
    str_User_Name = Application.UserName

If InStr(str_User_Name, ",") = False Then 'If they have a unique name then abort
    Name_Reverse = Replace(str_User_Name, ".", "")
    Exit Function
End If

Dim str_First_Name As String
    str_First_Name = Right(str_User_Name, Len(str_User_Name) - InStrRev(str_User_Name, ",") - 1)

Dim str_Last_Name As String
    str_Last_Name = Left(str_User_Name, InStrRev(str_User_Name, ",") - 1)

Dim str_Full_Name As String
    str_Full_Name = str_First_Name & " " & str_Last_Name

Name_Reverse = Replace(str_Full_Name, ".", "") 'Output the new user name, removes any periods after a middle initial

Exit Function

ErrorHandler:

MsgBox "There was an error with your username, please let James Rinaldi know and he'll fix it."

End Function
Function Global_Error_Handling(SubName, ErrSource, ErrNum, ErrDesc)

    Dim strTempVer As String
       strTempVer = Mid(String:=ThisWorkbook.Name, Start:=InStr(ThisWorkbook.Name, " (v") + 3, Length:=4)

If Err.Number <> 0 Then MsgBox _
    Title:="I am Error", _
    Buttons:=vbCritical, _
    Prompt:="Something went awry with the CV Tracker, try to hit Cancel and redo the last step. " _
    & "If that doesn't resolve it then reach out to James Rinaldi in Audit for a fix. " _
    & "This tool has a growth mindset, with each issue addressed we itterate to a better version." & Chr(10) & Chr(10) _
    & "Please take a screenshot of this message, and send it to James." & Chr(10) _
    & "Include a brief description of what you were doing when it occurred." & Chr(10) & Chr(10) _
    & "Error Source: " & ErrSource & " " & strTempVer & Chr(10) _
    & "Subroutine: " & SubName & Chr(10) _
    & "Error Desc.: #" & ErrNum & " - " & ErrDesc & Chr(10)

'Or include all of the details in an auto email to me and just prompt them for what happened.

End Function
Sub TEST_Error_Hanlder()

On Error GoTo ErrorHandler

Err.Raise (1)

ErrorHandler:

Global_Error_Handling SubName:="TEST", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Function fx_Create_Headers(strHeaderTitle As String, arryHeader As Variant)

' Purpose: To determine the column number for a specific title in the header.
' Trigger: Called
' Updated: 12/11/2020

' Change Log:
'       5/1/2020: Intial Creation
'       12/11/2020: Updated to use an array instead of the range, reducing the time to run by 75%.

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    Dim i As Integer

' -----------
' Loop through the array
' -----------

    For i = LBound(arryHeader) To UBound(arryHeader)
        If arryHeader(i, 1) = strHeaderTitle Then
            fx_Create_Headers = i
            Exit Function
        End If
    Next i

End Function
Function fx_Create_Unique_List(rngListValues As Range)

' Purpose: To create a unique list of values based on the passed range.
' Trigger: Called
' Updated: 5/5/2020

' Change Log:
'          5/5/2020: Intial Creation

' ****************************************************************************

' -----------
' Declare your variables
' -----------
           
    Dim strsUniqueValues As New Collection
    
    Dim strValue As Variant
    
    Dim arryTempData()
           
' -----------
' Copy in the data selected for rngListValues into the array, then into the collection
' -----------

    arryTempData = Application.Transpose(rngListValues)

On Error Resume Next 'If a duplicate is found skip it, instead of erroring
    
    For Each strValue In arryTempData
        strsUniqueValues.Add strValue, strValue
    Next

On Error GoTo 0

' -----------
' Pass the collection of unique values
' -----------

Set fx_Create_Unique_List = strsUniqueValues

End Function
Function fx_Privileged_User()

' Purpose: To output if the user is on the exception list or not.
' Trigger: Called
' Updated: 12/28/2020

' Change Log:
'       9/23/2020: Intial Creation
'       12/28/2020: Added the conditional compiler constant to determine if DebugMode was on, if so make Priviledged User false.
    
' ****************************************************************************

' -----------
' Declare your variables
' -----------

    Dim strUserID As String
        strUserID = Application.UserName

    Dim bolPrivilegedUser As Boolean

' -----------
' Determine if they are on the list
' -----------

    If _
        strUserID = "Rinaldi, James" Or _
        strUserID = "Duarte Espinoza, Axcel S." Or _
        strUserID = "Hogan, Elizabeth" Or _
        strUserID = "Barcikowski, Melissa H." Or _
        strUserID = "Soto, Jason A." Or _
        strUserID = "Gulick, Jocelyn" Or _
        strUserID = "Thompson, Gregory" Or _
        strUserID = "Zolty, Jonathan" Or _
        strUserID = "Demas, Stephanie" Or _
        strUserID = "Vargas, Laurie J." Or _
        strUserID = "Renzulli, Scott W." Or _
        strUserID = "Dowling, Mark J." Or _
        strUserID = "Mollo, Marie G." Or _
        strUserID = "Panno, Julio" _
    Then
        bolPrivilegedUser = True
    Else
        bolPrivilegedUser = False
    End If

    ' Turn off the Privileged User code if Debug Mode is on
    #If DebugMode = 1 Then
        bolPrivilegedUser = False
    #End If

    'Output the value to the function
    fx_Privileged_User = bolPrivilegedUser

End Function
Function fx_Array_Contains_Value(arryToSearch As Variant, valToFind As Variant) As Boolean

' Purpose: To determine if the given value is present in the array being searched.
' Trigger: Called
' Updated: 11/25/2020

' Change Log:
'          11/25/2020: Intial Creation
    
' ****************************************************************************

' -----------
' Declare your variables
' -----------

    Dim bolValueFound As Boolean

    Dim i As Long

' -----------
' Check for the value
' -----------
    
    For i = LBound(arryToSearch) To UBound(arryToSearch)
        If arryToSearch(i) = valToFind Then
            bolValueFound = True
            Exit For
        End If
    Next i
    
    fx_Array_Contains_Value = bolValueFound

End Function
Function fx_Open_Workbook(strPromptTitle As String) As Workbook
             
' Purpose: This function will prompt the user for the workbook to open and returns that workbook.
' Trigger: Called Function
' Updated: 2/12/2021
' Use Example: Set wbTEST = fx_Open_Workbook(strPromptTitle:="Select the current Sageworks data dump")

' Change Log:
'       2/12/2021: Initial Creation
'       2/12/2021: Added the code to abort if the user selects cancel.
'       2/12/2021: Added the code to determine if the Workbook is already open.

' ****************************************************************************
             
' -----------
' Declare your variables
' -----------
             
    Dim str_wbPath As String
        str_wbPath = Application.GetOpenFilename( _
        Title:=strPromptTitle, FileFilter:="Excel Workbooks (*.xls*;*.csv),*.xls*;*.csv")
             
        If str_wbPath = "False" Then
            MsgBox "No Workbook was selected, the code cannont continue."
            PrivateMacros.DisableForEfficiencyOff
            End
        End If
        
' -----------
' Determine if the Workbook is already open
' -----------
        
    Dim bolAlreadyOpen As Boolean
        
     Dim str_wbName As String
         str_wbName = Right(str_wbPath, Len(str_wbPath) - InStrRev(str_wbPath, "\"))
        
    On Error Resume Next
        Dim wb As Workbook
        Set wb = Workbooks(str_wbName)
        bolAlreadyOpen = Not wb Is Nothing
    On Error GoTo 0
        
' -----------
' Obtain the Workbook
' -----------
        
    If bolAlreadyOpen = True Then
        Set fx_Open_Workbook = Workbooks(str_wbName)
    Else
        Set fx_Open_Workbook = Workbooks.Open(str_wbPath, UpdateLinks:=False, ReadOnly:=True)
    End If

End Function

