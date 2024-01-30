Attribute VB_Name = "o_1_Import_CV_Mod_Agg_Data"
'Dim Workbooks / Sheets
    Dim wbCVModAgg As Workbook
    Dim wsAllCust As Worksheet
    Dim wsTracker As Worksheet
    Dim wsValidation As Worksheet
    
'Dim Integers
    Dim intLastRow_wsAllCust As Long
    Dim intLastCol_wsAllCust As Integer

    Dim intLastRow_wsTracker As Long
    Dim intLastCol_wsTracker As Integer
    
    Dim intCurRowValidation As Integer

'Dim Ranges
    Dim arryHeader_Tracker()
    Dim arryHeader_AllCust()

'Dim Arrays
    Dim ary_AllCustomers
    Dim ary_Tracker

'Dim "Ranges"
     Dim col_LastUsedCol_wsTracker As Integer
     Dim col_LastUsedCol_wsAllCust As Integer
     
'Dim CV Tracker Array "Ranges"
    Dim col_CustName_Tracker As Integer
    Dim col_LOB_Tracker As Integer
    Dim col_Helper_Tracker  As Integer
    Dim col_OrigMarket_Tracker As Integer
    Dim col_Market_Tracker As Integer
    Dim col_TIN_Tracker As Integer
    Dim col_RM_Tracker As Integer
    Dim col_PM_Tracker As Integer
    Dim col_Industry_Tracker As Integer
    Dim col_Exemption_Tracker As Integer
    
    Dim col_SIC_Tracker As Integer
    Dim col_BRG_Tracker As Integer
    Dim col_FRG_Tracker As Integer
    Dim col_CCRP_Tracker As Integer
    Dim col_LFT_Tracker As Integer
    
    Dim col_Outstanding_Tracker As Integer
    Dim col_Exposure_Tracker As Integer
    
    Dim col_AddedDate_Tracker As Integer
    Dim col_OpenDate_Tracker As Integer

'Dim CV Mod Aggregation Array "Ranges"
    Dim col_CustName_AllCust As Integer
    Dim col_LOB_AllCust As Integer
    Dim col_Helper_AllCust  As Integer
    Dim col_OrigMarket_AllCust As Integer
    Dim col_TIN_AllCust As Integer
    Dim col_RM_AllCust As Integer
    Dim col_PM_AllCust As Integer
    Dim col_Industry_AllCust As Integer
    
    Dim col_SIC_AllCust As Integer
    Dim col_BRG_AllCust As Integer
    Dim col_FRG_AllCust As Integer
    Dim col_CCRP_AllCust As Integer
    Dim col_LFT_AllCust As Integer
    
    Dim col_Outstanding_AllCust As Integer
    Dim col_Exposure_AllCust As Integer
    
    Dim col_OpenDate_AllCust As Integer

Option Explicit
Sub o_01_MAIN_PROCEDURE()

' Purpose: To import the data from the 6330 report to call out new customers and update the risk ratings and outstanding balances.
' Trigger: N/A
' Updated: 5/13/2021

' Change Log:
'       4/23/2020: Intial Creation
'       11/25/2020: Added the highlights for anomolies code
'       5/12/2021: Removed the o_33_Highlight_Anomolies_Old, it wasn't flagging anything meaningful to be addressed.
'       5/13/2021: Added the code for o_51_Validate_Control_Totals and automated the review of the validation totals.

' ****************************************************************************

Call PrivateMacros.DisableForEfficiency

    Call o_1_Import_CV_Mod_Agg_Data.o_02_Assign_Private_Variables
    Call o_1_Import_CV_Mod_Agg_Data.o_03_Assign_Array_Variables
    
    Call o_1_Import_CV_Mod_Agg_Data.o_11_Import_New_Customers
    Call o_1_Import_CV_Mod_Agg_Data.o_12_Flag_Changed_or_Old_Customers

    Call o_1_Import_CV_Mod_Agg_Data.o_2_Import_Static_Fields

    Call o_1_Import_CV_Mod_Agg_Data.o_31_Copy_Array_Values_to_wsData
    Call o_1_Import_CV_Mod_Agg_Data.o_32_Highlight_Anomolies_New
    
    Call o_1_Import_CV_Mod_Agg_Data.o_41_Hide_Columns
    
    Call o_1_Import_CV_Mod_Agg_Data.o_51_Validate_Control_Totals

Call PrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_02_Assign_Private_Variables()

' Purpose: To declare all of the Public variables that were dimensioned "above the line".
' Trigger: Called
' Updated: 3/2/2021

' Change Log:
'       4/23/2020: Intial Creation
'       6/5/2020: Updated intLastRow_wsTracker to look for newly added accounts
'       8/31/2020: Added code to unhide everything in the Data tab of the CMML CV Tracker to make updates easier
'       2/12/2021: Switched the workbook open to a function
'       3/2/2021: Added the code to unhide all of the rows in the CV Tracker

' ****************************************************************************

'Sets the current directory and path

On Error Resume Next
    ChDrive ThisWorkbook.Path
        ChDir ThisWorkbook.Path
On Error GoTo 0

' -----------
' Declare your variables
' -----------
    
    'Dim Workbooks / Worksheets
    
        Set wbCVModAgg = Functions.fx_Open_Workbook(strPromptTitle:="Select the current CV Mod Aggregation workbook")
    
        Set wsAllCust = wbCVModAgg.Sheets("ALL Customers")
            wsAllCust.Cells.EntireColumn.Hidden = False 'Unhide all columns
            wsAllCust.Cells.EntireRow.Hidden = False 'Unhide all rows
            
        Set wsTracker = ThisWorkbook.Sheets("Data")
            wsTracker.AutoFilter.ShowAllData 'Unhide all rows
            wsTracker.Cells.EntireColumn.Hidden = False 'Unhide all columns
            wsTracker.Cells.EntireRow.Hidden = False 'Unhide all rows
            
        Set wsValidation = ThisWorkbook.Sheets("VALIDATION")
            
    'Dim Integers
        
        intLastRow_wsAllCust = wsAllCust.Cells(Rows.Count, "A").End(xlUp).Row
        intLastCol_wsAllCust = wsAllCust.Cells(1, Columns.Count).End(xlToLeft).Column
        intLastRow_wsTracker = Application.WorksheetFunction.Max(wsTracker.Cells(Rows.Count, "B").End(xlUp).Row, wsTracker.Cells(Rows.Count, "D").End(xlUp).Row)
        intLastCol_wsTracker = wsTracker.Cells(1, Columns.Count).End(xlToLeft).Column
        
    'Dim Ranges
    
        arryHeader_Tracker = Application.Transpose(wsTracker.Range(wsTracker.Cells(1, 1), wsTracker.Cells(1, intLastCol_wsTracker)))
        arryHeader_AllCust = Application.Transpose(wsAllCust.Range(wsAllCust.Cells(1, 1), wsAllCust.Cells(1, intLastCol_wsAllCust)))
        
    'Dim "Ranges"
        
        col_LastUsedCol_wsTracker = fx_Create_Headers("Gross Exposure", arryHeader_Tracker)
        col_LastUsedCol_wsAllCust = fx_Create_Headers("Gross Exposure", arryHeader_AllCust)

    'Dim Arrays
    
        ary_AllCustomers = wsAllCust.Range(wsAllCust.Cells(1, 1), wsAllCust.Cells(intLastRow_wsAllCust, col_LastUsedCol_wsAllCust))
        ary_Tracker = wsTracker.Range(wsTracker.Cells(1, 1), wsTracker.Cells(intLastRow_wsTracker, col_LastUsedCol_wsTracker))
    
End Sub
Sub o_03_Assign_Array_Variables()

' Purpose: To declare all of the Public variables that were dimensioned for the array comparison.
' Trigger: Called o_1_Import_CV_Mod_Agg_Data
' Updated: 7/25/2020

' Change Log:
'          4/23/2020: Intial Creation
'          7/25/2020: Updated to include the lookup for TIN.

' -----------
' Declare your variables
' -----------

    Dim i As Integer

' -----------
' Declare the CV Tracker Array "Ranges"
' -----------

    col_CustName_Tracker = fx_Create_Headers("Customer Name", arryHeader_Tracker)
    col_LOB_Tracker = fx_Create_Headers("Line of Business", arryHeader_Tracker)
    col_Helper_Tracker = fx_Create_Headers("HELPER", arryHeader_Tracker)
    col_OrigMarket_Tracker = fx_Create_Headers("Original Market", arryHeader_Tracker)
    col_Market_Tracker = fx_Create_Headers("Market", arryHeader_Tracker)
    col_TIN_Tracker = fx_Create_Headers("TIN", arryHeader_Tracker)
    col_RM_Tracker = fx_Create_Headers("Relationship Manager", arryHeader_Tracker)
    col_PM_Tracker = fx_Create_Headers("Portfolio Manager", arryHeader_Tracker)
    col_Industry_Tracker = fx_Create_Headers("Scorecard Industry", arryHeader_Tracker)
    'col_Exemption_Tracker = fx_Create_Headers("Exemption", arryHeader_Tracker) -> Temp disable, causing issues on 5/17
    'col_Exemption_Tracker = 20
    
    col_SIC_Tracker = fx_Create_Headers("SIC", arryHeader_Tracker)
    col_BRG_Tracker = fx_Create_Headers("BRG", arryHeader_Tracker)
    col_FRG_Tracker = fx_Create_Headers("FRG", arryHeader_Tracker)
    col_CCRP_Tracker = fx_Create_Headers("CCRP", arryHeader_Tracker)
    col_LFT_Tracker = fx_Create_Headers("LFT", arryHeader_Tracker)

    col_Outstanding_Tracker = fx_Create_Headers("Direct Outstanding", arryHeader_Tracker)
    col_Exposure_Tracker = fx_Create_Headers("Gross Exposure", arryHeader_Tracker)
    
    col_AddedDate_Tracker = fx_Create_Headers("Added to Tracker", arryHeader_Tracker)
    col_OpenDate_Tracker = fx_Create_Headers("Date Opened", arryHeader_Tracker)

' -----------
' Declare the CV Mod Aggregation Array "Ranges"
' -----------

    col_CustName_AllCust = fx_Create_Headers("Customer Name", arryHeader_AllCust)
    col_LOB_AllCust = fx_Create_Headers("Line of Business", arryHeader_AllCust)
    col_Helper_AllCust = fx_Create_Headers("HELPER", arryHeader_AllCust)
    col_OrigMarket_AllCust = fx_Create_Headers("Market", arryHeader_AllCust)
    col_TIN_AllCust = fx_Create_Headers("Tax ID Number", arryHeader_AllCust)
    col_RM_AllCust = fx_Create_Headers("Relationship Manager", arryHeader_AllCust)
    col_PM_AllCust = fx_Create_Headers("Portfolio Manager", arryHeader_AllCust)
    col_Industry_AllCust = fx_Create_Headers("Industry", arryHeader_AllCust)

    col_SIC_AllCust = fx_Create_Headers("SIC Code", arryHeader_AllCust)
    col_BRG_AllCust = fx_Create_Headers("BRG", arryHeader_AllCust)
    col_FRG_AllCust = fx_Create_Headers("FRG", arryHeader_AllCust)
    col_CCRP_AllCust = fx_Create_Headers("CCRP", arryHeader_AllCust)
    col_LFT_AllCust = fx_Create_Headers("LFT", arryHeader_AllCust)

    col_Outstanding_AllCust = fx_Create_Headers("Outstanding", arryHeader_AllCust)
    col_Exposure_AllCust = fx_Create_Headers("Gross Exposure", arryHeader_AllCust)

    col_OpenDate_AllCust = fx_Create_Headers("Open Date", arryHeader_AllCust)

    'Close the CV Mod Aggregation Workbook
        wbCVModAgg.Close saveChanges:=False

End Sub
Sub o_11_Import_New_Customers()

' Purpose: To import the new customers form the CV Mod Aggregation array into the CV Tracker array.
' Trigger: Called o_1_Import_CV_Mod_Agg_Data.o_01_MAIN_PROCEDURE
' Updated: 5/12/2021

' Change Log:
'       5/4/2020: Intial Creation
'       6/5/2020: Added code so if there are no records delete
'       7/24/2020: Switch from looking at just the Cust Name to the Helper record when doing the compare.
'       3/4/2021: Added code to remove the Webster Financial Corp customer
'       5/12/2021: Updated the code to properly hide the fields, things got shifted w/ the addition of Portfolio Manager

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    Dim intRowCnt As Long
    
    Dim arry_Tracker_Custs
        arry_Tracker_Custs = Application.Transpose(Application.Index(ary_Tracker, , col_Helper_Tracker))
        
    'Declare Customer Variance variables
    
    Application.DisplayAlerts = False
        If Evaluate("ISREF('Not in CV Tracker'!A1)") = True Then ThisWorkbook.Sheets("Not in CV Tracker").Delete
        ThisWorkbook.Worksheets.Add().Name = "Not in CV Tracker"
    Application.DisplayAlerts = True

    Dim wsOnlyCVMods As Worksheet
        Set wsOnlyCVMods = ThisWorkbook.Sheets("Not in CV Tracker")

    Dim aryRowData
        aryRowData = Application.WorksheetFunction.Index(ary_AllCustomers, 1, 0)
        wsOnlyCVMods.Range(wsOnlyCVMods.Cells(1, 1), wsOnlyCVMods.Cells(1, col_LastUsedCol_wsAllCust)) = aryRowData
        
    Dim intLastRow_wsOnlyCVMods As Integer
        intLastRow_wsOnlyCVMods = wsOnlyCVMods.Cells(Rows.Count, "A").End(xlUp).Row + 1
        
    Dim i As Integer

' -----------
' Copy the CMML customers that are unique to the CV Mod Agg. into the CV Tracker
' -----------
    
    For intRowCnt = 2 To intLastRow_wsAllCust

        If IsNumeric(Application.Match(ary_AllCustomers(intRowCnt, col_Helper_AllCust), arry_Tracker_Custs, 0)) Then
        Else 'The ones unique to the CV Mod Aggregation
            aryRowData = Application.WorksheetFunction.Index(ary_AllCustomers, intRowCnt, 0)
            wsOnlyCVMods.Range(wsOnlyCVMods.Cells(intLastRow_wsOnlyCVMods, 1), wsOnlyCVMods.Cells(intLastRow_wsOnlyCVMods, col_LastUsedCol_wsAllCust)) = aryRowData
                If wsOnlyCVMods.Range("F" & intLastRow_wsOnlyCVMods) = "Middle Market Banking - WEBSTER FINANCIAL CORP" Then
                    wsOnlyCVMods.Range("F" & intLastRow_wsOnlyCVMods).EntireRow.Delete
                    intLastRow_wsOnlyCVMods = intLastRow_wsOnlyCVMods - 1
                End If
            intLastRow_wsOnlyCVMods = intLastRow_wsOnlyCVMods + 1
        End If
   
    Next intRowCnt

' -----------
' Format the data
' -----------

    With wsOnlyCVMods
        .Range(.Cells(1, 1), .Cells(1, col_LastUsedCol_wsAllCust)).EntireColumn.AutoFit
    End With

    With wsOnlyCVMods
        .Range("H1:I1").EntireColumn.Hidden = True 'TIN Helper -> Industry
        .Range("L1:R1").EntireColumn.Hidden = True 'SIC Code -> Non-Accrual
        .Range("T1").EntireColumn.Hidden = True 'PPP Flag
    End With

' -----------
' Delete the ws if there are no records
' -----------

    If intLastRow_wsOnlyCVMods = 2 Then
    
        Application.DisplayAlerts = False
            wsOnlyCVMods.Delete
        Application.DisplayAlerts = True
    
    End If

End Sub
Sub o_12_Flag_Changed_or_Old_Customers()

' Purpose: To copy old / new customers into the applicable ws.
' Trigger: Called o_1_Import_CV_Mod_Agg_Data.o_01_MAIN_PROCEDURE
' Updated: 5/12/2021

' Change Log:
'       5/4/2020: Intial Creation
'       6/5/2020: Added code so if there are no records delete
'       7/24/2020: Switch from looking at just the Cust Name to the Helper record when doing the compare.
'       5/12/2021: Added code to apply the $ formatting to the balances

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    Dim intRowCnt As Long
  
    Dim arry_Helper_CVMods
        arry_Helper_CVMods = Application.Transpose(Application.Index(ary_AllCustomers, , col_Helper_AllCust))
        
    'Declare Customer Variance variables
    
    Application.DisplayAlerts = False
        If Evaluate("ISREF('Only in CV Tracker'!A1)") = True Then ThisWorkbook.Sheets("Only in CV Tracker").Delete
        ThisWorkbook.Worksheets.Add().Name = "Only in CV Tracker"
    Application.DisplayAlerts = True

    Dim wsOnlyTracker As Worksheet
        Set wsOnlyTracker = ThisWorkbook.Sheets("Only in CV Tracker")

    Dim aryRowData
        aryRowData = Application.WorksheetFunction.Index(ary_Tracker, 1, 0)
        wsOnlyTracker.Range(wsOnlyTracker.Cells(1, 1), wsOnlyTracker.Cells(1, col_LastUsedCol_wsTracker)) = aryRowData

    Dim intLastRow_wsOnlyTracker As Integer
        intLastRow_wsOnlyTracker = wsOnlyTracker.Cells(Rows.Count, "A").End(xlUp).Row + 1

' -----------
' Copy the CMML customers that are not in the CV Mod Agg.
' -----------
    
    For intRowCnt = 2 To intLastRow_wsTracker

        If IsNumeric(Application.Match(ary_Tracker(intRowCnt, col_Helper_Tracker), arry_Helper_CVMods, 0)) Then
        Else
            aryRowData = Application.WorksheetFunction.Index(ary_Tracker, intRowCnt, 0)
            wsOnlyTracker.Range(wsOnlyTracker.Cells(intLastRow_wsOnlyTracker, 1), wsOnlyTracker.Cells(intLastRow_wsOnlyTracker, col_LastUsedCol_wsTracker)) = aryRowData
            intLastRow_wsOnlyTracker = intLastRow_wsOnlyTracker + 1
        End If
    
    Next intRowCnt

' -----------
' Format the data
' -----------

    With wsOnlyTracker
        .Range(.Cells(1, 1), .Cells(1, col_LastUsedCol_wsAllCust)).EntireColumn.AutoFit
    End With

    With wsOnlyTracker
        .Range(.Cells(2, 27), .Cells(intLastRow_wsOnlyTracker, 28)).NumberFormat = "$#,##0.00"
    End With

    With wsOnlyTracker
        .Range("F1:S1").EntireColumn.Hidden = True
    End With

' -----------
' Delete the ws if there are no records
' -----------

    If intLastRow_wsOnlyTracker = 2 Then
    
        Application.DisplayAlerts = False
            wsOnlyTracker.Delete
        Application.DisplayAlerts = True
    
    End If

End Sub
Sub o_2_Import_Static_Fields()

' Purpose: To import the data from the CV Mod Aggregation array.
' Trigger: Called o_1_Import_CV_Mod_Agg_Data.o_01_MAIN_PROCEDURE
' Updated: 7/3/2021

' Change Log:
'       5/4/2020: Intial Creation
'       6/5/2020: Removed the CMMLOnly fields now that the CV Mod Agg. is exported w/ CMML Only
'       1/17/2021: Updated to convert the accounts for Kristen and Melissa
'       3/4/2021: Moved the code related to Original Market to apply to ALL customers
'       3/26/2021: (1) Since row index is added 1 at the end after reaching bottom of array, must create a "trap"
'                     that will catch any executions for a row index that is more than the fields that are contained in the array.
'       5/12/2021: Updated to import the Portfolio Manager field
'       5/12/2021: Simplified the code to reduce the if statements and refresh all fields for all borrowers
'       5/13/2021: Updated to wipe the old validation data
'       5/16/2021: Added code to skip the LOB update for customers with an exemption
'       5/17/2021: Removed code for Kristen and Melissa to be overwriten as RM
'       7/3/2021: Uncommented the code related to updating the LOB

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    Dim intRowAllCust As Long
        intRowAllCust = 2
    
    Dim intRowTracker As Long
    
    Dim intAllCustOutstandingSum As Double
    
    Dim intTrackerOutstandingSum As Double
    
    Dim rngOrigMarket As Range
        Set rngOrigMarket = ThisWorkbook.Sheets("Lists").Range("O2:O99")
    
    Dim i As Long

' -----------
' Copy the static data into the CV Tracker array
' -----------
    
    For intRowTracker = 2 To intLastRow_wsTracker
    
    intRowAllCust = 1

LoopStart:
        
        If intRowTracker > intLastRow_wsTracker Then Exit For
        
        If ary_Tracker(intRowTracker, col_Helper_Tracker) = ary_AllCustomers(intRowAllCust, col_Helper_AllCust) Then  ' If record matches
            
            ary_Tracker(intRowTracker, 1) = intRowTracker - 1                                                               'Ref #
            
            ary_Tracker(intRowTracker, col_CustName_Tracker) = ary_AllCustomers(intRowAllCust, col_CustName_AllCust)        'Cust Name
'            If ary_Tracker(intRowTracker, col_Exemption_Tracker) = "" Then ' Only update if there is no exemption
'                ary_Tracker(intRowTracker, col_LOB_Tracker) = ary_AllCustomers(intRowAllCust, col_LOB_AllCust)              'LOB
'            End If
            ary_Tracker(intRowTracker, col_TIN_Tracker) = ary_AllCustomers(intRowAllCust, col_TIN_AllCust)                  'TIN
            ary_Tracker(intRowTracker, col_OrigMarket_Tracker) = ary_AllCustomers(intRowAllCust, col_OrigMarket_AllCust)    'Original Market
            ary_Tracker(intRowTracker, col_Market_Tracker) = _
                rngOrigMarket.Find(ary_Tracker(intRowTracker, col_OrigMarket_Tracker)).Offset(0, 1).Value2                  'Market

            ary_Tracker(intRowTracker, col_RM_Tracker) = ary_AllCustomers(intRowAllCust, col_RM_AllCust)                    'Relationship Manager
            ary_Tracker(intRowTracker, col_PM_Tracker) = ary_AllCustomers(intRowAllCust, col_PM_AllCust)                    'Portfolio Manager

            ary_Tracker(intRowTracker, col_Industry_Tracker) = ary_AllCustomers(intRowAllCust, col_Industry_AllCust)        'Industry
            ary_Tracker(intRowTracker, col_SIC_Tracker) = ary_AllCustomers(intRowAllCust, col_SIC_AllCust)                  'SIC
            
            ary_Tracker(intRowTracker, col_BRG_Tracker) = ary_AllCustomers(intRowAllCust, col_BRG_AllCust)                  'BRG
            ary_Tracker(intRowTracker, col_FRG_Tracker) = ary_AllCustomers(intRowAllCust, col_FRG_AllCust)                  'FRG
            ary_Tracker(intRowTracker, col_CCRP_Tracker) = ary_AllCustomers(intRowAllCust, col_CCRP_AllCust)                'CCRP
            ary_Tracker(intRowTracker, col_LFT_Tracker) = ary_AllCustomers(intRowAllCust, col_LFT_AllCust)                  'LFT
            
            ary_Tracker(intRowTracker, col_Outstanding_Tracker) = ary_AllCustomers(intRowAllCust, col_Outstanding_AllCust)  'Outstanding
            ary_Tracker(intRowTracker, col_Exposure_Tracker) = ary_AllCustomers(intRowAllCust, col_Exposure_AllCust)        'Exposure
            
        Else
            
            intRowAllCust = intRowAllCust + 1
                
                '3/26/2021 - AD: Consider moving this out of the Else statement and move it as a step/verification before the If statement starts.
                
                'If you got to the end of AllCust then reset and start the loop again
                If intRowAllCust > intLastRow_wsAllCust Then
                    intRowTracker = intRowTracker + 1
                    intRowAllCust = 1
                    GoTo LoopStart
                End If
            
            GoTo LoopStart
            
        End If
   
    Next intRowTracker

' -----------
' Wipe out the old validation data
' -----------
    
    wsValidation.Range("A2:G4").ClearContents

' -----------
' Output the control totals
' -----------

    For i = 2 To UBound(ary_AllCustomers)
        intAllCustOutstandingSum = intAllCustOutstandingSum + ary_AllCustomers(i, col_Outstanding_AllCust)
    Next i

    For i = 2 To UBound(ary_Tracker)
        intTrackerOutstandingSum = intTrackerOutstandingSum + ary_Tracker(i, col_Outstanding_Tracker)
    Next i

    With wsValidation
        intCurRowValidation = .Range("A:A").Find("").Row
        'intCurRowValidation = .Cells(Rows.Count, "A").End(xlUp).Row
    
        .Range("A" & intCurRowValidation + 0) = Now
        .Range("B" & intCurRowValidation + 0) = "o_1_Import_CV_Mod_Agg_Data"
        .Range("C" & intCurRowValidation + 0) = "All Customers Array"
        .Range("D" & intCurRowValidation + 0) = Format(intAllCustOutstandingSum, "$#,##0")
        .Range("F" & intCurRowValidation + 0) = Format(UBound(ary_AllCustomers) - 1, "0,0")
    
        .Range("A" & intCurRowValidation + 1) = Now
        .Range("B" & intCurRowValidation + 1) = "o_1_Import_CV_Mod_Agg_Data"
        .Range("C" & intCurRowValidation + 1) = "CV Tracker Array"
        .Range("D" & intCurRowValidation + 1) = Format(intTrackerOutstandingSum, "$#,##0")
        .Range("F" & intCurRowValidation + 1) = Format(UBound(ary_Tracker) - 1, "0,0")

    End With

End Sub
Sub o_31_Copy_Array_Values_to_wsData()
    
' Purpose: To copy the updated values from the Array back into the Data ws.
' Trigger: Called o_1_Import_CV_Mod_Agg_Data.o_01_MAIN_PROCEDURE
' Updated: 5/13/2021

' Change Log:
'       4/30/2020: Intial Creation
'       5/13/2021: Updated to use Find to add the validation data

' ****************************************************************************
    
    wsTracker.Range(wsTracker.Cells(1, 1), wsTracker.Cells(intLastRow_wsTracker, col_LastUsedCol_wsTracker)) = ary_Tracker
    'wsTracker.Range(wsTracker.Cells(1, 1), wsTracker.Cells(intLastRow_wsTracker, intLastCol_wsTracker)) = ary_Tracker

' -----------
' Output the control totals
' -----------

    With wsValidation
        intCurRowValidation = .Range("A:A").Find("").Row
        'intCurRowValidation = .Cells(Rows.Count, "A").End(xlUp).Row
    
        .Range("A" & intCurRowValidation) = Now
        .Range("B" & intCurRowValidation) = "o_1_Import_CV_Mod_Agg_Data"
        .Range("C" & intCurRowValidation) = "CV Tracker - wsData"
        .Range("D" & intCurRowValidation) = Format(Application.WorksheetFunction.Sum(wsTracker.Range(wsTracker.Cells(2, col_Outstanding_Tracker), wsTracker.Cells(intLastRow_wsTracker, col_Outstanding_Tracker))), "$#,##0")
        .Range("F" & intCurRowValidation) = Format(wsTracker.Cells(Rows.Count, "B").End(xlUp).Row - 1, "0,0")

    End With

End Sub
Sub o_32_Highlight_Anomolies_New()

' Purpose: To highlight the new / removed customers that need to be addressed, including false positives.
' Trigger: Called o_1_Import_CV_Mod_Agg_Data.o_01_MAIN_PROCEDURE
' Updated: 11/27/2020

' Change Log:
'       11/27/2020: Intial Creation
'       5/12/2021: Updated the code for dtPriorMonth to use the v_PriorRunDate now that this process is run less then monthly
    
' ****************************************************************************

    If Evaluate("ISREF('Not in CV Tracker'!A1)") = False Then Exit Sub
    
' -----------
' Declare your variables
' -----------
    
    ' Dim Worksheets
    
    Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("Not in CV Tracker")
    
    ' Dim "Ranges"
    Dim arryHeader()
        arryHeader = Application.Transpose(ws.Range("A1:Z1"))
    
    Dim col_OpenDate As Integer
        col_OpenDate = fx_Create_Headers("Open Date", arryHeader)
        
    Dim col_TIN As Integer
        col_TIN = fx_Create_Headers("Tax ID Number", arryHeader)
    
    ' Dim Integers
    Dim intLastRow As Integer
        intLastRow = ws.Cells(Rows.Count, "B").End(xlUp).Row
    
    ' Dim Loop Variables
    Dim i As Integer

    Dim aryTIN() As Variant
        aryTIN = Application.Transpose(Application.Index(ary_Tracker, 0, col_TIN_Tracker))
        
    Dim dtPriorMonth As Date
        dtPriorMonth = DateValue(Range("v_PriorRunDate"))

    ' Dim Colors
    Dim intOrange As Long
        intOrange = RGB(254, 233, 217)

' -----------
' Highlight changes in the 'Not in CV Tracker' ws
' -----------

    For i = 2 To intLastRow
        If ws.Cells(i, col_OpenDate).Value <= dtPriorMonth Then ws.Cells(i, col_OpenDate).Interior.Color = intOrange
        If fx_Array_Contains_Value(aryTIN, ws.Cells(i, col_TIN)) = True Then ws.Cells(i, col_TIN).Interior.Color = intOrange
    Next i

End Sub
Sub o_41_Hide_Columns()

' Purpose: To unhide the worksheets that the PMs don't need to see.
' Trigger: Called o_1_Import_CV_Mod_Agg_Data.o_01_MAIN_PROCEDURE
' Updated: 3/2/2021

' Change Log:
'       3/2/2021: Initial Creation

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    ' Dim "Ranges"

    Dim col_SubSector_Tracker As Integer
        col_SubSector_Tracker = fx_Create_Headers("Sub Sector", arryHeader_Tracker)

    Dim col_Vulnerable_Tracker As Integer
        col_Vulnerable_Tracker = fx_Create_Headers("Vulnerable & Not Reviewed Elsewhere", arryHeader_Tracker)

' -----------
' Hide the unused Columns
' -----------

wsTracker.Range(wsTracker.Cells(1, col_SubSector_Tracker), wsTracker.Cells(1, col_Vulnerable_Tracker)).EntireColumn.Hidden = True
'wsTracker.Columns(col_SubSector_Tracker & ":" & col_Vulnerable_Tracker).EntireColumn.Hidden = True

Call PrivateMacros.DisableForEfficiency

End Sub
Sub o_51_Validate_Control_Totals()

' Purpose: To validate that the control totals for the data imported match.
' Trigger: Called o_1_Import_CV_Mod_Agg_Data.o_01_MAIN_PROCEDURE
' Updated: 5/13/2021

' Change Log:
'       5/13/2021: Intial Creation, from Sageworks project

' ****************************************************************************

On Error GoTo ErrorHandler

' -----------
' Declare your variables
' -----------
    
    ' Dim Integers
    
    Dim int1stTotal As Double
        int1stTotal = wsValidation.Range("D2").Value
    
    Dim int2ndTotal As Double
        int2ndTotal = wsValidation.Range("D3").Value
    
    Dim int3rdTotal As Double
        int3rdTotal = wsValidation.Range("D4").Value
    
    Dim int1stCount As Long
        int1stCount = wsValidation.Range("F2").Value
    
    Dim int2ndCount As Long
        int2ndCount = wsValidation.Range("F3").Value
        
    Dim int3rdCount As Long
        int3rdCount = wsValidation.Range("F4").Value
        
    ' Dim Booleans
    
    Dim bolTotalsMatch As Boolean ' If the three totals match, and the three counts match, then True
        If int1stTotal = int2ndTotal And int2ndTotal = int3rdTotal And _
           int1stCount = int2ndCount And int2ndCount = int3rdCount Then
            bolTotalsMatch = True
        Else
            bolTotalsMatch = False
        End If
    
' -----------
' Output the messagebox with the results
' -----------
   
    If bolTotalsMatch = True Then
    MsgBox Title:="Control Totals Match", _
        Buttons:=vbOKOnly, _
        Prompt:="The validation totals match between the source and output." & Chr(10) & Chr(10) _
        & "Validation Total: " & Format(int1stTotal, "$#,##0") & Chr(10) _
        & "Validation Count: " & Format(int1stCount, "0,0")
       
    ElseIf bolTotalsMatch = False Then
    MsgBox Title:="Control Totals DO NOT Match", _
        Buttons:=vbCritical, _
        Prompt:="The validation totals from the CMML CV Tracker don't match the data in the CV Mod Agg. (what was imported). " _
        & "Please review the totals in the Validation worksheet to determine what went awry. " _
        & "Pay special attention to new customers that need to have their 'Helper' field added to the Data tab." & Chr(10) & Chr(10) _
        & "1st Validation Total: " & Format(int1stTotal, "$#,##0") & Chr(10) _
        & "1st Validation Total Variance: " & Format(int1stTotal - int2ndTotal, "$#,##0") & Chr(10) & Chr(10) _
        & "1st Validation Count: " & Format(int1stCount, "0,0") & Chr(10) _
        & "1st Validation Count Variance: " & Format(int1stCount - int2ndCount, "0,0")
    
    End If

Exit Sub
    
ErrorHandler:

PrivateMacros.DisableForEfficiencyOff
    
End Sub

