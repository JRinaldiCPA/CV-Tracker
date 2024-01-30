VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_CV_Tracker_Admin 
   Caption         =   "CMML CV Tracker - ADMIN"
   ClientHeight    =   3300
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   12912
   OleObjectBlob   =   "uf_CV_Tracker_Admin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_CV_Tracker_Admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Dim Worksheets
    Dim wsData As Worksheet
    Dim wsLists As Worksheet
    Dim wsArrays As Worksheet
    Dim wsChangeLog As Worksheet
    Dim wsUpdates As Worksheet
    Dim wsChecklist As Worksheet
    Dim wsValidation As Worksheet
    Dim wsFormulas As Worksheet
    Dim wsLOBReview As Worksheet

'Dim Strings
    Dim strCustomerName As String
    Dim strNewFileFullPath As String 'Used by o_51_Create_a_XLSX_Copy to create the XLSX to attach to the email
    Dim strLastCol_wsData As String
    Dim strUserID As String

'Dim Integers
    Dim intLastRow_wsData As Long
    Dim intLastCol_wsData As Integer
    Dim col_Offset As Integer

'Dim Flag "Ranges"
    Dim col_NewFirst As Integer
    Dim col_NewLast As Integer
    Dim col_OldFirst As Integer
    Dim col_OldLast As Integer
    Dim col_ChangeFlag As Integer
    Dim col_FirstJason As Integer
    Dim col_LastJason As Integer
    Dim col_DateHighRisk As Integer
    
'Dim CDR "Ranges"
    Dim col_Customer As Integer
    Dim col_LOB As Integer
    Dim col_Helper As Integer
    Dim col_OrigMarket As Integer
    Dim col_Market As Integer
    Dim col_Exempt As Integer

    Dim col_Vulnerable As Integer
    
'Dim PM "Ranges"
    Dim col_2Q_RefreshReq As Integer
    Dim col_2Q_RefreshComplete As Integer
    Dim col_PPP As Integer
    Dim col_MSLP As Integer
    Dim col_4013CARES As Integer
    
    Dim col_SupplyConcern As Integer
    Dim col_EconConcern As Integer
    Dim col_OverallRisk  As Integer
    
'Dim Arrays / Other
    Dim bolPrivilegedUser As Boolean
    
    Dim arryHeader()

    Dim ary_Customers
    Dim ary_SelectedCustomers
    Dim ary_LOB_Customers
    Dim ary_Market_Customers
Private Sub cmd_Cancel_Click()
    
    Unload Me

End Sub

Private Sub UserForm_Initialize()
' ****************************************************************************
'
' Author:   James Rinaldi
' Created:  1/29/2021
' Updated:  1/29/2021
'
' ----------------------------------------------------------------------------
'
' Purpose:  To track changes to the Commercial Loan Portfolio as a result of the 2020 Covid-19 pandemic.
'           Portfolio Managers assess their customers on a weekly basis and submit changes using this UserForm.
'           These changes are aggregated, and all changes are applied to the data.
'
' Trigger:  Workbook Open / Quick Access Toolbar
'
' Change Log:
'          1/29/2021: Intial Creation based on Regular Dashboard
'
' ****************************************************************************

Call Me.o_02_Assign_Private_Variables

' -----------
' Initialize the initial values
' -----------
    Me.StartUpPosition = 0 'Allow you to set the position
        Me.Top = Application.Top + (Application.UsableHeight / 1.5) - (Me.Height / 2) 'Open near the bottom of the screen
        Me.Left = Application.Left + (Application.UsableWidth / 2) - (Me.Width / 2)
    
        Call Me.o_62_UnProtect_Ws

        Call Me.o_64_UnHide_Worksheets

End Sub
Sub o_02_Assign_Private_Variables()

' Purpose: To declare all of the Public variables that were dimensioned "above the line".
' Trigger: Called
' Updated: 9/24/2020

' Change Log:
'       4/23/2020: Intial Creation
'       9/24/2020: Updated the Exception User to be a function instead

' ****************************************************************************

' -----------
' Declare your variables
' -----------
    
    'Dim sheets
    
            Set wsData = ThisWorkbook.Sheets("Data")
    
            Set wsLists = ThisWorkbook.Sheets("Lists")
            
            Set wsArrays = ThisWorkbook.Sheets("Array Values")
    
            Set wsChangeLog = ThisWorkbook.Sheets("Change Log")
            
            Set wsUpdates = ThisWorkbook.Sheets("Updates")
    
            Set wsChecklist = ThisWorkbook.Sheets("Checklist")
            
            Set wsValidation = ThisWorkbook.Sheets("VALIDATION")
            
            Set wsFormulas = ThisWorkbook.Sheets("FORMULAS")
            
            Set wsLOBReview = ThisWorkbook.Sheets("LOB Detail Review")
    
    'Dim Integers / Strings
    
        'Public intLastCol_wsData As Integer
            intLastCol_wsData = wsData.Cells(1, Columns.Count).End(xlToLeft).Column
            
        'Public intLastRow_wsData As Long
            intLastRow_wsData = wsData.Cells(Rows.Count, "A").End(xlUp).Row
            'intLastRow_wsData = wsData.Range("A:A").Find("").Row - 1
                        
        'Public strLastCol_wsData as String
            strLastCol_wsData = Split(Cells(1, intLastCol_wsData).Address, "$")(1)
            
        Dim intLastRow_wsArrays As Long
            intLastRow_wsArrays = wsArrays.Cells(Rows.Count, "A").End(xlUp).Row
            If intLastRow_wsArrays = 1 Then intLastRow_wsArrays = 2
      
    'Dim Strings
            
        'Dim strUserID As String
            strUserID = Application.UserName
            
    'Dim Ranges
        
        'Dim arryHeader()
            arryHeader = Application.Transpose(wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, intLastCol_wsData)))
    
    'Dim Flag "Ranges"
            
        col_NewFirst = fx_Create_Headers("Property Type", arryHeader)
        col_NewLast = fx_Create_Headers("Relief Comments", arryHeader)
        
        col_Offset = col_NewLast - col_NewFirst + 1

        col_OldFirst = col_NewFirst + col_Offset
        col_OldLast = col_NewLast + col_Offset
    
        col_ChangeFlag = fx_Create_Headers("Change Flag", arryHeader)
        
        col_FirstJason = fx_Create_Headers("1st Payment Mod", arryHeader) - 1
        
        col_DateHighRisk = fx_Create_Headers("Date High Overall Risk", arryHeader)
            col_LastJason = col_DateHighRisk
            
    'Dim CDR "Ranges"
    
        col_Customer = fx_Create_Headers("Customer Name", arryHeader)
        col_LOB = fx_Create_Headers("Line of Business", arryHeader)
        col_Helper = fx_Create_Headers("HELPER", arryHeader)
        col_OrigMarket = fx_Create_Headers("Original Market", arryHeader)
        col_Market = fx_Create_Headers("Market", arryHeader)
    
        col_Vulnerable = fx_Create_Headers("Vulnerable & Not Reviewed Elsewhere", arryHeader)
        col_Exempt = fx_Create_Headers("Exemption", arryHeader)
            
        col_SupplyConcern = fx_Create_Headers("Supply Chain Concern", arryHeader)
        col_EconConcern = fx_Create_Headers("US Economic Concern", arryHeader)
        col_OverallRisk = fx_Create_Headers("OVERALL CONCERN", arryHeader)
            
    'Dim PM "Ranges"
        col_2Q_RefreshReq = fx_Create_Headers("2Q Rating Refresh Req.", arryHeader)
        col_2Q_RefreshComplete = fx_Create_Headers("2Q Rating Refresh Completed", arryHeader)
        
        col_4013CARES = fx_Create_Headers("4013 CARES Act Eligible", arryHeader)
            
        col_PPP = fx_Create_Headers("PPP", arryHeader)
        col_MSLP = fx_Create_Headers("MSLP Inquiry?", arryHeader)
                
    'Dim Arrays
            
        Dim i As Integer
        
        'Public ary_SelectedCustomers
            If wsArrays.Range("A2").Value2 = "" Then
                ReDim ary_SelectedCustomers(1) 'Only set if there aren't old customers in there
            Else
                ReDim ary_SelectedCustomers(1 To 999)
                
                For i = 2 To intLastRow_wsArrays
                    ary_SelectedCustomers(i - 1) = wsArrays.Range("A" & i)
                Next i
                
            End If
    
        'Public ary_Customers
            ReDim ary_Customers(1)

    'Dim Booleans
    
        'Dim bolPrivilegedUser As Boolean
            bolPrivilegedUser = fx_Privileged_User

End Sub
Private Sub cmd_Prep_File_Click()

' Purpose: To prepare the file for the PMs, CV Impact Meeting, or output to the CV Mod Aggregation workbook.
' Trigger: cmd_Prep_File
' Updated: 2/3/2021

' Change Log:
'       10/13/2020: Added the code for the CV Mod Aggregation
'       2/3/2021: Added the code to apply the formulas as part of the prep for PM distribution

' ****************************************************************************

Call PrivateMacros.DisableForEfficiency

' -----------
' Declare your variables
' -----------
    
    Dim intRun_PrepforDist As Integer
        intRun_PrepforDist = MsgBox("Do you want to prepare the file for distribution to the PMs?", vbYesNo + vbQuestion, "Prep for Distribution")
        If intRun_PrepforDist = vbYes Then GoTo RunCode

    Dim intRun_PrepforCVImpact As Integer
        intRun_PrepforCVImpact = MsgBox("Do you want to prepare the file for the CV Impact meeting?", vbYesNo + vbQuestion, "Prep for CV Impact")
        If intRun_PrepforCVImpact = vbYes Then GoTo RunCode

    Dim intRun_ExportForCVModAgg As Integer
        intRun_ExportForCVModAgg = MsgBox("Do you want to export the file for the CV Mod Aggregation process?", vbYesNo + vbQuestion, "Export for CV Mod. Aggregation")
        If intRun_ExportForCVModAgg = vbYes Then GoTo RunCode

    'If nothing else has been triggered, abort
    GoTo Abort
    
' -----------
' Run the update process
' -----------

RunCode:

    Call Me.o_41_Clear_Filters
    
    Call Me.o_42_Clear_Saved_Array
       
    If intRun_PrepforDist = vbYes Then
        Call Me.o_43_Prep_CV_Tracker_For_Distribution
        Call o_2_Import_PM_Updates.o_02_Assign_Private_Variables
        Call o_2_Import_PM_Updates.o_23_Apply_Formulas
        Call o_2_Import_PM_Updates.o_24_Refresh_Pivots
    ElseIf intRun_PrepforCVImpact = vbYes Then
        Call Me.o_44_Prep_For_CV_Impact_Call
    ElseIf intRun_ExportForCVModAgg = vbYes Then
        Call Me.o_43_Prep_CV_Tracker_For_Distribution
        Call Me.o_45_Export_For_CV_Mod_Aggregation
        Call PrivateMacros.DisableForEfficiencyOff
        ActiveWorkbook.Close saveChanges:=False 'Close the workbook
    Else
        Call PrivateMacros.DisableForEfficiencyOff
    End If
    
    Unload Me

Abort:

Call PrivateMacros.DisableForEfficiencyOff

End Sub
Private Sub cmd_Update_w_PM_Updates_Click()

    Call o_2_Import_PM_Updates.o_01_MAIN_PROCEDURE_Update_Data_ws
    
    Unload Me

End Sub
Private Sub cmd_Update_w_CV_Mod_Agg_Data_Click()
    
    Call o_1_Import_CV_Mod_Agg_Data.o_01_MAIN_PROCEDURE
    
    Unload Me

End Sub
Sub o_41_Clear_Filters()

' Purpose: To reset all of the current filtering.
' Trigger: Called: uf_CV_Tracker_Regular.cmd_Clear_Filter
' Updated: 8/19/2020

' Change Log:
'   3/23/2020: Intial Creation
'   8/19/2020: Added the logic to exclude the exempt customers
'   9/2/2020: Added the logic to filter out the blank rows

' ****************************************************************************

On Error GoTo ErrorHandler

Call Me.o_62_UnProtect_Ws

' -----------
' If the AutoFilter is on turn it off and then reapply
' -----------

    If wsData.AutoFilterMode = True Then
        wsData.AutoFilter.ShowAllData
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_Exempt, Criteria1:="=", Operator:=xlFilterValues 'Hide exempt values
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_Customer, Criteria1:="*", Operator:=xlFilterValues
    End If
    
Call Me.o_61_Protect_Ws

Exit Sub

ErrorHandler:

    Global_Error_Handling SubName:="o_41_Clear_Filters", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    
    Call Me.o_61_Protect_Ws

End Sub
Sub o_42_Clear_Saved_Array()

On Error GoTo ErrorHandler

' Purpose: To remove all of the values from the Selected Customers Array and the Array ws.
' Trigger: Called: uf_CV_Tracker_Regular.cmd_Clear_Filter
' Updated: 3/23/2020

' Change Log:
'   3/23/2020: Intial Creation

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    Dim intLastRow As Long
        intLastRow = wsArrays.Cells(Rows.Count, "A").End(xlUp).Row

' -----------
' Remove the old values and empty the array
' -----------

    If intLastRow <> 1 Then
        wsArrays.Range("A2:A" & intLastRow).Clear
    End If

    If IsEmpty(ary_SelectedCustomers) = False Then Erase ary_SelectedCustomers

Exit Sub

ErrorHandler:

    Global_Error_Handling SubName:="o_42_Clear_Saved_Array", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_43_Prep_CV_Tracker_For_Distribution()

' Purpose: To prepare the CV Tracker for sending out to everyone.
' Trigger: Called: uf_CV_Tracker_Regular.cmd_Prep_File
' Updated: 11/10/2020

' Change Log:
'       4/1/2020: Added in a sort based on the customer #
'       4/6/2020: Added the code to hide the extra columns for Jason's CV Impact meeting
'       4/6/2020: Added a sort for the LOB / CUstomer in LIsts to show alphabetical
'       8/19/2020: Updated to clear any filters on the Updates or Change Log ws
'       9/28/2020: Removed redundant code related to clearing filters and arrays
'       11/10/2020: Updated to convert the sorting to regular, not using AutoFilter
'       11/10/2020: Updated to handle ALL the different versions of the R&R cost centers
'       3/3/2021: Added code to force the Prep for Distribution to turn on the AutoFilter
'       5/6/2021: Commented out the code around Supply Concern, so that the PMs can review for the 5/14 data

' ****************************************************************************

If bolPrivilegedUser = False Then Exit Sub 'Abort if your not one of the chosen few

' -----------
' Dimension your variables
' -----------
    
    'Dim Integers
    
    Dim intLastRow_wsChangeLog As Long
         intLastRow_wsChangeLog = wsChangeLog.Range("A:A").Find("").Row
            If intLastRow_wsChangeLog = 1 Then intLastRow_wsChangeLog = 2
        
    Dim intLastRow_wsUpdates As Long
        intLastRow_wsUpdates = wsUpdates.Range("A:A").Find("").Row
            If intLastRow_wsUpdates = 1 Then intLastRow_wsUpdates = 2

    Dim intLastRow_wsLists As Long
        intLastRow_wsLists = wsLists.Cells(Rows.Count, "E").End(xlUp).Row

    Dim i As Long

    'Dim CDR "Ranges"

    Dim col_SubSector As Integer
        col_SubSector = fx_Create_Headers("Sub Sector", arryHeader)

    Dim col_ExemptDeets As Integer
        col_ExemptDeets = fx_Create_Headers("Exemption Details", arryHeader)

' -----------
' Copy over the old data
' -----------
    
Application.EnableEvents = False '9/2/20: Added to prevent the change macro from starting during an update
    
    With wsData
        .Range(.Cells(2, col_OldFirst), .Cells(intLastRow_wsData, col_OldLast)).Value2 = _
        .Range(.Cells(2, col_NewFirst), .Cells(intLastRow_wsData, col_NewLast)).Value2
    End With
        
Application.EnableEvents = True
        
' -----------
' Clear the logs
' -----------
    
    'If the AutoFilter isn't on already then turn it on
    If wsChangeLog.AutoFilterMode = False Then
        wsChangeLog.Range("A:G").AutoFilter
    Else
        wsChangeLog.AutoFilter.ShowAllData
    End If
    
    wsChangeLog.Range(wsChangeLog.Cells(2, 1), wsChangeLog.Cells(intLastRow_wsChangeLog, 7)).ClearContents 'Include Source
    
    'If the AutoFilter isn't on already then turn it on
    If wsUpdates.AutoFilterMode = False Then
        wsUpdates.Range("A:F").AutoFilter
    Else
        wsUpdates.AutoFilter.ShowAllData
    End If
    
    wsUpdates.Range(wsUpdates.Cells(2, 1), wsUpdates.Cells(intLastRow_wsUpdates, 6)).ClearContents

' -----------
' Unhide the columns and rows required to be updated
' -----------
        
    If wsData.AutoFilterMode = True Then
        wsData.AutoFilter.ShowAllData
    Else
        wsData.Cells.AutoFilter
    End If
        
    With wsData
        .Range(.Cells(1, 1), .Cells(1, col_NewLast)).EntireColumn.Hidden = False 'Unhide the current data
        .Range(.Cells(1, 1), .Cells(intLastRow_wsData, 1)).EntireRow.Hidden = False 'Unhide ALL Rows, incase the data was filtered
    End With
    
    intLastRow_wsData = wsData.Cells(Rows.Count, "A").End(xlUp).Row 'Reset the LastRow

' -----------
' Clear the filters
' -----------
    
    Call Me.o_41_Clear_Filters
    
' -----------
' Make sure the old columns are hidden and the Data ws is sorted
' -----------

    With wsData
        .Range(.Cells(1, col_OrigMarket), .Cells(1, col_Helper)).EntireColumn.Hidden = True       'Hide Helper & Original Market
        .Range(.Cells(1, col_SubSector), .Cells(1, col_Vulnerable)).EntireColumn.Hidden = True     'Hide Sub-Sector - Vulnerable
        .Range(.Cells(1, col_Exempt), .Cells(1, col_ExemptDeets)).EntireColumn.Hidden = True        'Hide Exemption - Exemption Details
        .Range(.Cells(1, col_OldFirst), .Cells(1, col_OldLast)).EntireColumn.Hidden = True        'Hide the 'OLD' data
        .Range(.Cells(1, col_FirstJason), .Cells(1, col_LastJason)).EntireColumn.Hidden = True    'Hide the columns for the CV Impact meeting
        .Range(.Cells(1, col_LastJason + 2), .Cells(1, Columns.Count)).EntireColumn.Hidden = True    'Hide the blank columns to the right
        
        .Cells(1, col_PPP).EntireColumn.Hidden = True
        .Cells(1, col_MSLP).EntireColumn.Hidden = True
        .Cells(1, col_2Q_RefreshReq).EntireColumn.Hidden = True
        .Cells(1, col_2Q_RefreshComplete).EntireColumn.Hidden = True
        .Cells(1, col_4013CARES).EntireColumn.Hidden = True
        
        '.Cells(1, col_SupplyConcern).EntireColumn.Hidden = True ' 5/6/21: Removed so that people can review at the upcoming meeting
        .Cells(1, col_EconConcern).EntireColumn.Hidden = True
        
        .Range(.Cells(1, 1), .Cells(intLastRow_wsData, intLastCol_wsData)).Sort Key1:=.Range("A:A"), Order1:=xlAscending, Header:=xlYes
        .Range(.Cells(2, 1), .Cells(intLastRow_wsData, 1)).RowHeight = 45
        .Range(.Cells(intLastRow_wsData + 1, 1), .Cells(.Rows.Count, 1)).EntireRow.Hidden = True 'Hide all the rows at the bottom
    End With

' -----------
' Reset the row height
' -----------

    wsData.Range("2:" & intLastRow_wsData).Rows.RowHeight = 45

' -----------
' Update the Lists with the new customers
' -----------
   
    intLastRow_wsData = wsData.Cells(Rows.Count, "A").End(xlUp).Row 'Reset the LastRow
   
    'Remove the old data
    wsLists.Range("C1:E" & intLastRow_wsLists).ClearContents
   
    With wsData
        .Range(.Cells(1, col_LOB), .Cells(intLastRow_wsData, col_LOB)).Copy Destination:=wsLists.Range("C1") 'LOB
        .Range(.Cells(1, col_Market), .Cells(intLastRow_wsData, col_Market)).Copy Destination:=wsLists.Range("D1") 'Market
        .Range(.Cells(1, col_Customer), .Cells(intLastRow_wsData, col_Customer)).Copy Destination:=wsLists.Range("E1") 'Customer Name
    End With

    'Convert the LOB for the R&R LOB
    For i = 2 To intLastRow_wsData
        If InStr(1, wsLists.Range("D" & i), "Remediation") > 0 Then
            wsLists.Range("C" & i) = "Restructure & Recovery"
        End If
    Next i

    'Reset the Last Row
    
    intLastRow_wsLists = wsLists.Cells(Rows.Count, "E").End(xlUp).Row

    'Sort based on LOB, Customer Name...not Market that causes issues

    wsLists.Range("C1:E" & intLastRow_wsLists).Sort _
        Key1:=wsLists.Range("C1"), Order1:=xlAscending, _
        Key2:=wsLists.Range("E1"), Order2:=xlAscending, _
        Header:=xlYes
    
' -----------
' Delete the Temp worksheets, if they exist
' -----------

    Application.DisplayAlerts = False
        If Evaluate("ISREF('Temp Arrays'!A1)") = True Then ThisWorkbook.Sheets("Temp Arrays").Delete
        If Evaluate("ISREF('Only in CV Tracker'!A1)") = True Then ThisWorkbook.Sheets("Only in CV Tracker").Delete
        If Evaluate("ISREF('Not in CV Tracker'!A1)") = True Then ThisWorkbook.Sheets("Not in CV Tracker").Delete
    Application.DisplayAlerts = True

' -----------
' Hide the other ws's
' -----------

    Call Me.o_63_Hide_Worksheets

End Sub
Sub o_44_Prep_For_CV_Impact_Call()

' Purpose: To prepare the CV Tracker for the CV Impact Call on Thursdays.
' Trigger: Called: uf_CV_Tracker_Regular.cmd_Prep_File
' Updated: 8/20/2020

' Change Log:
'          5/27/2020: Initial Creation
'          8/20/2020: Hiding additional fields, including Property Type, End Market, and the 2Q Rating Refresh fields

' ****************************************************************************

' -----------
' Dimension your variables
' -----------
    
    'Dim CDR "Ranges"
    
    Dim col_RelMgr As Integer
        col_RelMgr = fx_Create_Headers("Relationship Manager", arryHeader)
        
    Dim col_LFT As Integer
        col_LFT = fx_Create_Headers("LFT", arryHeader)
    
    'Dim PM Updated "Ranges"
    
    Dim col_PropType As Integer
        col_PropType = fx_Create_Headers("Property Type", arryHeader)

    Dim col_EndMarket As Integer
        col_EndMarket = fx_Create_Headers("End Market", arryHeader)
    
    Dim col_Comments As Integer
        col_Comments = fx_Create_Headers("Comments", arryHeader)
    
' -----------
' Unhide everything first
' -----------
    
    With wsData
        .Range(.Cells(1, 1), .Cells(1, col_DateHighRisk)).EntireColumn.Hidden = False 'Unhide the current data
        .Range(.Cells(1, 1), .Cells(intLastRow_wsData, 1)).EntireRow.Hidden = False 'Unhide ALL Rows, incase the data was filtered
    End With
    
' -----------
' Make sure the old columns are hidden and the Data ws is sorted
' -----------

    With wsData
        .Range(.Cells(1, col_OrigMarket), .Cells(1, col_Helper)).EntireColumn.Hidden = True   'Hide Helper & Original Market
        .Range(.Cells(1, col_RelMgr), .Cells(1, col_Vulnerable)).EntireColumn.Hidden = True  'Hide Portfolio Manager - Vulnerable
        .Range(.Cells(1, col_Exempt), .Cells(1, col_LFT)).EntireColumn.Hidden = True          'Hide Exemption - LFT
        .Cells(1, col_LFT).EntireColumn.Hidden = False                                           'Unhide LFT
        .Cells(1, col_PropType).EntireColumn.Hidden = True
        .Cells(1, col_EndMarket).EntireColumn.Hidden = True
        .Cells(1, col_Comments).EntireColumn.Hidden = True
        .Cells(1, col_2Q_RefreshReq).EntireColumn.Hidden = True
        .Cells(1, col_2Q_RefreshComplete).EntireColumn.Hidden = True
        .Cells(1, col_4013CARES).EntireColumn.Hidden = True
        .Cells(1, col_PPP).EntireColumn.Hidden = True
        .Cells(1, col_MSLP).EntireColumn.Hidden = True
        .Range(.Cells(1, col_OldFirst), .Cells(1, col_OldLast)).EntireColumn.Hidden = True 'Hide the 'OLD' data
    End With

' -----------
' Reset the row height
' -----------

    wsData.Range("2:" & intLastRow_wsData).Rows.RowHeight = 30

End Sub
Sub o_45_Export_For_CV_Mod_Aggregation()

' Purpose: To export the CV Tracker for use in the CV Mod Aggregation process.
' Trigger: Called: uf_CV_Tracker_Regular.cmd_Prep_File
' Updated: 3/6/2021

' Change Log:
'       10/14/2020: Initial Creation
'       12/3/2020: Automated the conversion to an xlsx
'       3/6/2021: Update the export process for the CV Mod Agg to save and close the workbook

' ****************************************************************************

Application.DisplayAlerts = False
Application.EnableEvents = False

' -----------
' Dimension your variables
' -----------

    Dim ws As Worksheet

    'Dim Strings

    Dim strNewFileName As String
        strNewFileName = Replace(ThisWorkbook.Name, ".xlsm", " (CV Mod Agg).xlsx")
            strNewFileName = "(2) " & strNewFileName
    
    Dim strNewFileFullPath As String
        strNewFileFullPath = ThisWorkbook.Path & "\" & strNewFileName
    
' -----------
' Save back wsData as values only
' -----------
    If wsData.AutoFilterMode = True Then wsData.AutoFilter.ShowAllData

    wsData.Range(wsData.Cells(1, 1), wsData.Cells(2, intLastCol_wsData)).EntireColumn.Hidden = False
    wsData.Range(wsData.Cells(2, 1), wsData.Cells(intLastRow_wsData, intLastCol_wsData)).Value2 = wsData.Range(wsData.Cells(2, 1), wsData.Cells(intLastRow_wsData, intLastCol_wsData)).Value2

' -----------
' Delete the other worksheets
' -----------
    
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
        If ws.Name <> wsData.Name Then ws.Delete
    Next ws
    
' -----------
' Delete the prior columns
' -----------
    
    wsData.Range(wsData.Cells(1, col_OldFirst), wsData.Cells(1, intLastCol_wsData)).EntireColumn.Delete
    
' -----------
' Save the workbook as the version for the CV Mod agg.
' -----------
    
    ThisWorkbook.SaveAs Filename:=strNewFileFullPath, FileFormat:=xlOpenXMLWorkbook
    
Application.DisplayAlerts = True
Application.EnableEvents = True

End Sub
Sub o_61_Protect_Ws()

On Error GoTo ErrorHandler

' -----------
' Skip the data protection if the user is on the exception list
' -----------

    If bolPrivilegedUser = True Then Exit Sub

' -----------
' Turn on data protection
' -----------

    With wsData
        .Protect AllowFiltering:=True, AllowSorting:=True
        .EnableAutoFilter = True
        .EnableSelection = xlUnlockedCells
    End With

Exit Sub

ErrorHandler:

    Debug.Print Global_Error_Handling("o_61_Protect_Ws", Err.Source, Err.Number, Err.Description)
    Global_Error_Handling SubName:="o_21_Add_Customer_To_Selected_Customers_Array", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_62_UnProtect_Ws()

    wsData.Unprotect

End Sub
Sub o_63_Hide_Worksheets()

' Purpose: To hide the worksheets that the PMs don't need to see.
' Trigger: Called: uf_CV_Tracker_Regular
' Updated: 9/25/2020

' Change Log:
'       9/25/2020: Initial Creation

' ****************************************************************************

Call PrivateMacros.DisableForEfficiency

    wsArrays.Visible = xlSheetVeryHidden
    wsLists.Visible = xlSheetVeryHidden
    wsChangeLog.Visible = xlSheetVeryHidden
    wsUpdates.Visible = xlSheetVeryHidden
    wsChecklist.Visible = xlSheetHidden
    wsValidation.Visible = xlSheetHidden
    wsFormulas.Visible = xlSheetHidden
    'wsLOBReview.Visible = xlSheetHidden

Call PrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_64_UnHide_Worksheets()

' Purpose: To unhide the worksheets that the PMs don't need to see.
' Trigger: Called: uf_CV_Tracker_Regular
' Updated: 9/25/2020

' Change Log:
'       9/25/2020: Initial Creation

' ****************************************************************************

Call PrivateMacros.DisableForEfficiency

    'Make all the worksheets visible for Exception Users
    wsLists.Visible = xlSheetVisible
    wsChangeLog.Visible = xlSheetVisible
    wsUpdates.Visible = xlSheetVisible
    wsChecklist.Visible = xlSheetVisible
    wsValidation.Visible = xlSheetVisible
    wsFormulas.Visible = xlSheetVisible
    wsLOBReview.Visible = xlSheetVisible

Call PrivateMacros.DisableForEfficiencyOff

End Sub
