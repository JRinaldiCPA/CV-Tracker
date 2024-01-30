Attribute VB_Name = "o_2_Import_PM_Updates"
Option Explicit

'Dim Workbooks / Sheets
    Dim wsData As Worksheet
    Dim wsValidation As Worksheet
    Dim wsUpdates As Worksheet
    Dim wsFormulas As Worksheet
    
'Dim Integers
    Dim intLastRow As Long
    Dim intLastCol As Integer
    
    Dim intCurRowValidation As Integer

'Dim Colors
    Dim intOrange As Long
    Dim intWhite As Long

'Dim Ranges / "Ranges"
    Dim arryHeader()
    Dim arryHeaderOldOnly()
    
    Dim col_CustName As Integer
    Dim col_Outstanding As Integer
    Dim col_OverallRisk As Integer
    Dim col_Relief As Integer
    Dim col_Relief2nd As Integer
    Dim col_ModType As Integer
    Dim col_ModType2nd As Integer
    
    Dim col_ActiveMod As Integer
    Dim col_PaymentMod As Integer
    Dim col_ChangeFlag As Integer

    Dim col_NewFirst As Integer
    Dim col_NewLast As Integer
    Dim col_Offset As Integer

'Dim Arrays
    Dim ary_Data
    Dim ary_Updates
'
Sub o_01_MAIN_PROCEDURE_Update_Data_ws()

' Purpose: To manipulate the Data ws to include all the additional exception reporting and whatnot we need.
' Trigger: N/A
' Updated: 5/11/2020

' Change Log:
'          5/11/2020: Intial Creation

' ****************************************************************************

Call PrivateMacros.DisableForEfficiency
Application.EnableEvents = False 'Stop the update macro from running

    Call o_02_Assign_Private_Variables
    
    Call o_1_Update_Data_from_Change_Logs

    Call o_21_Update_Date_Relief_Requested

    Call o_22_Update_Date_High_Overall_Risk

    Call o_23_Apply_Formulas
    
    Call o_24_Refresh_Pivots
    
    Call o_31_Over_5MM_No_Rating

    Call o_32_Relief_But_No_Mod

    Call o_33_Q2_Rating_Refresh_Not_Completed

    Call o_34_Missing_End_Market

    Call o_35_Missing_Property_Type

    Call o_36_2nd_Round_Relief_Only

    Call o_37_Missing_Mod_Maturity_Date

Application.EnableEvents = True
Call PrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_02_Assign_Private_Variables()

' Purpose: To declare all of the Public variables that were dimensioned "above the line".
' Trigger: Called by o_01_MAIN_PROCEDURE_Update_Data_ws
' Updated: 5/11/2020

' Change Log:
'          5/11/2020: Intial Creation
'          9/14/2020: Updated to include the Formulas

' ****************************************************************************

' -----------
' Declare your variables
' -----------
    
    'Dim Workbooks / Worksheets
    
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsValidation = ThisWorkbook.Sheets("VALIDATION")
    Set wsUpdates = ThisWorkbook.Sheets("Updates")
    Set wsFormulas = ThisWorkbook.Sheets("FORMULAS")
    
    'Dim Integers
    
    intLastRow = wsData.Cells(Rows.Count, "A").End(xlUp).Row
    intLastCol = wsData.Cells(1, Columns.Count).End(xlToLeft).Column
        
    'Dim Colors
    
    intOrange = RGB(254, 233, 217)
    intWhite = RGB(255, 255, 255)
    
    'Dim Ranges
    
    arryHeader = Application.Transpose(wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, intLastCol)))
                   
    'Dim "Ranges"

    col_CustName = fx_Create_Headers("Customer Name", arryHeader)
    col_Outstanding = fx_Create_Headers("Direct Outstanding", arryHeader)
    col_OverallRisk = fx_Create_Headers("OVERALL CONCERN", arryHeader)
    col_Relief = fx_Create_Headers("1st Round Relief", arryHeader)
    col_Relief2nd = fx_Create_Headers("2nd Round Relief", arryHeader)
    col_ModType = fx_Create_Headers("1st Round Modification Type", arryHeader)
    col_ModType2nd = fx_Create_Headers("2nd Round Modification Type", arryHeader)
    
    col_ActiveMod = fx_Create_Headers("Active Mod", arryHeader)
    col_PaymentMod = fx_Create_Headers("Active Payment Mod", arryHeader)
    col_ChangeFlag = fx_Create_Headers("Change Flag", arryHeader)
    
    col_NewFirst = fx_Create_Headers("Property Type", arryHeader)
    col_NewLast = fx_Create_Headers("Relief Comments", arryHeader)
    col_Offset = col_NewLast - col_NewFirst + 1
    
    'Dim "Old Ranges"
            
    arryHeaderOldOnly = Application.Transpose(wsData.Range(wsData.Cells(1, col_NewLast + 1), wsData.Cells(1, intLastCol)))
        
End Sub
Sub o_1_Update_Data_from_Change_Logs()

' Purpose: To update the data in the Data ws based on the update files provided by the Portfolio Managers.
' Trigger: N/A
' Updated: 8/20/2020

' Change Log:
'          5/11/2020: Moved into o_2_Import_PM_Updates and updated
'          8/20/2020: Moved into o_2_Import_PM_Updates and updated

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    'Dim Integers
    
        Dim intLastRow_wsUpdates As Long
            intLastRow_wsUpdates = wsUpdates.Cells(Rows.Count, "A").End(xlUp).Row
                If intLastRow_wsUpdates = 1 Then Exit Sub 'If there is no data abort

        Dim intRowData As Long
        
        Dim col_Data As Long
        
        Dim intRowUpdates As Long
        
        Dim col_Updates As Long
        
        Dim col_Match As Integer 'Find the matching column
        
        Dim intChangeCount As Integer
            intChangeCount = 0
        
' -----------
' Clear the filters on the Data ws
' -----------
        
    wsData.AutoFilter.ShowAllData
                
' -----------
' Sort the Updates ws
' -----------

    If wsUpdates.AutoFilterMode = False Then
        wsUpdates.Range("A:F").AutoFilter
    End If

    With wsUpdates.AutoFilter.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("C2:C" & intLastRow_wsUpdates), Order:=xlAscending
        .SortFields.Add Key:=Range("D2:D" & intLastRow_wsUpdates), Order:=xlAscending
        .SortFields.Add Key:=Range("A2:A" & intLastRow_wsUpdates), Order:=xlAscending
        .Header = xlYes
        .Apply
    End With

' -----------
' Set your Arrays
' -----------
     
    'Dim ary_Data
        ary_Data = wsData.Range(wsData.Cells(1, 1), wsData.Cells(intLastRow, col_NewLast))
        
    'Dim ary_Updates
        ary_Updates = wsUpdates.Range("A1:F" & intLastRow_wsUpdates)

' -----------
' Identify any changes to the customer data
' -----------
    intRowUpdates = 2  'How can I get this to exit when it hits the first match?
    
    For intRowData = 2 To intLastRow
        For intRowUpdates = 2 To intLastRow_wsUpdates
            If ary_Data(intRowData, col_CustName) = ary_Updates(intRowUpdates, 3) Then 'If Cust. Name matches
                                
                For col_Match = 2 To col_NewLast
                    If Trim(Replace(ary_Updates(intRowUpdates, 4), Chr(10), "")) = Trim(Replace(ary_Data(1, col_Match), Chr(10), "")) Then Exit For 'Remove the line breaks (chr(10) if there are any
                Next
                                                       
                If col_Match <= col_NewLast Then
                    ary_Data(intRowData, col_Match) = ary_Updates(intRowUpdates, 6)
                    intChangeCount = intChangeCount + 1
                
                End If
            End If
        Next intRowUpdates
    Next intRowData

' -----------
' Output the control totals
' -----------
    
    Dim intCustCount As Integer
        intCustCount = fx_Create_Unique_List(wsUpdates.Range("C1:C" & intLastRow_wsUpdates)).Count
    
    With wsValidation
        intCurRowValidation = .Cells(Rows.Count, "A").End(xlUp).Row
    
        .Range("A" & intCurRowValidation + 1) = Now
        .Range("B" & intCurRowValidation + 1) = "o_2_Import_PM_Updates"
        .Range("C" & intCurRowValidation + 1) = "CV Tracker - CV File Agg."
        .Range("F" & intCurRowValidation + 1) = Format(intCustCount, "0,0") 'Customer Count
        .Range("G" & intCurRowValidation + 1) = Format(intChangeCount, "0,0") 'Change Count

    End With

' -----------
' Copy the changes to the Data ws
' -----------

    'Pull the data into the wsData
           
    Application.EnableEvents = False 'Stop the update macro from running
        wsData.Range(wsData.Cells(1, 1), wsData.Cells(intLastRow, col_NewLast)).Value2 = ary_Data
    Application.EnableEvents = True
    
End Sub
Sub o_21_Update_Date_Relief_Requested()

' Purpose: To update the Date Relief Requested field based on the updated data from this update.
' Trigger: Called by o_01_MAIN_PROCEDURE_Update_Data_ws
' Updated: 5/11/2020

' Change Log:
'          5/11/2020: Intial Creation

' ****************************************************************************

' -----------
' Declare your variables
' -----------
    
    'Dim Strings
        Dim strCurDate As String
            strCurDate = Date
        
    'Dim Integers
        Dim i As Integer
        
    'Dim "Ranges"
            
        'Dim col_Relief As Integer
            'col_Relief = fx_Create_Headers("1st Round Relief", arryHeader)
            
        'Dim col_Relief2nd As Integer
            'col_Relief2nd = fx_Create_Headers("2nd Round Relief", arryHeader)
            
        Dim col_DateReliefReq As Integer
            col_DateReliefReq = fx_Create_Headers("Date Relief Requested", arryHeader)
        
        Dim col_DateReliefReq2nd As Integer
            col_DateReliefReq2nd = fx_Create_Headers("Date 2nd Relief Requested", arryHeader)
            
        Dim col_ReliefOld As Integer
            col_ReliefOld = fx_Create_Headers("1st Round Relief", arryHeaderOldOnly) + col_NewLast
        
        Dim col_Relief2ndOld As Integer
            col_Relief2ndOld = fx_Create_Headers("2nd Round Relief", arryHeaderOldOnly) + col_NewLast
            
' -----------
' Run the loop to input the date for any new Relief requests
' -----------

    With wsData
        For i = 2 To intLastRow
            
            If Left(.Cells(i, col_Relief), 1) = "Y" Then 'First Relief
                If Left(.Cells(i, col_ReliefOld), 1) <> "Y" Then
                    .Cells(i, col_DateReliefReq).Value = strCurDate
                End If
            End If
            
            If Left(.Cells(i, col_Relief2nd), 1) = "Y" Then 'Second Relief
                If Left(.Cells(i, col_Relief2ndOld), 1) <> "Y" Then
                    .Cells(i, col_DateReliefReq2nd).Value = strCurDate
                End If
            End If
        
        Next i
    End With

End Sub
Sub o_22_Update_Date_High_Overall_Risk()

' Purpose: To update the Date High Overall Risk field based on the updated data from this update.
' Trigger: Called by o_01_MAIN_PROCEDURE_Update_Data_ws
' Updated: 5/11/2020

' Change Log:
'          5/11/2020: Intial Creation

' ****************************************************************************

' -----------
' Declare your variables
' -----------
    
    'Dim Strings
        Dim strCurDate As String
            strCurDate = Date
        
    'Dim Integers
        Dim i As Integer
        
    'Dim "Ranges"
            
        'Dim col_OverallRisk As Integer
            'col_OverallRisk = fx_Create_Headers("OVERALL CONCERN", arryHeader)
            
        Dim col_OverallRiskOld As Integer
            col_OverallRiskOld = fx_Create_Headers("OVERALL CONCERN", arryHeaderOldOnly) + col_NewLast
            
        Dim col_DateHighRisk As Integer
            col_DateHighRisk = fx_Create_Headers("Date High Overall Risk", arryHeader)
        
' -----------
' Run the loop to input the date for any new High Risk customers
' -----------

    With wsData
        For i = 2 To intLastRow
            
            If .Cells(i, col_OverallRisk) = "High" Then
                If .Cells(i, col_OverallRiskOld) <> "High" Then
                    .Cells(i, col_DateHighRisk).Value = strCurDate
                End If
            End If
            
        Next i
    End With

End Sub
Sub o_23_Apply_Formulas()

' Purpose: To copy the formulas from the FORMULAS ws to the Data tab.
' Trigger: Called by o_01_MAIN_PROCEDURE_Update_Data_ws
' Updated: 9/14/2020

' Change Log:
'          8/19/2020: Intial Creation
'          9/3/2020: Combined various formula subs into one
'          9/3/2020: Added the Payment Mod formula
'          9/14/2020: Replaced the macro with the code from the CV Mod Aggregation workbook
'          11/6/2020: Updated to include the Change Flag field

' ****************************************************************************

' -----------
' Declare your variables
' -----------
   
    'Dim Ranges
    
        'Dim arryHeader()
            'arryHeader = wsData.Range("1:" & intLastCol)
            
    'Dim "Ranges"
        
    Dim intFirstCol_wsData As Integer
        intFirstCol_wsData = fx_Create_Headers("Active Mod", arryHeader)
        
    Dim intLastCol_wsData As Integer
        intLastCol_wsData = fx_Create_Headers("Active Payment Mod", arryHeader)
        
    Dim intFirstRow_wsFormulas As Integer
        intFirstRow_wsFormulas = wsFormulas.Range("C:C").Find("Active Mod").Row

    Dim intLastRow_wsFormulas As Integer
        intLastRow_wsFormulas = wsFormulas.Range("C:C").Find("Active Payment Mod").Row
        
    Dim intRow_ChangeFlag_wsformulas As Integer
        intRow_ChangeFlag_wsformulas = wsFormulas.Range("C:C").Find("Change Flag").Row
        
    'Dim Loop Variables
    
    Dim x As Integer
    
    Dim y As Integer
    
    Dim strFormula As String
    
' -----------
' Copy the Active Mod formulas into the All Customers ws
' -----------

    For x = intFirstCol_wsData To intLastCol_wsData
    
        For y = intFirstRow_wsFormulas To intLastRow_wsFormulas
        
            If wsData.Cells(1, x) = wsFormulas.Cells(y, 3) Then
                strFormula = wsFormulas.Cells(y, 4).Value2
                wsData.Range(wsData.Cells(2, x), wsData.Cells(intLastRow, x)).Formula = strFormula
    
                Exit For
            End If
        Next y
    
    Next x

' -----------
' Copy the Change Flag formula into the All Customers ws
' -----------

    strFormula = wsFormulas.Cells(intRow_ChangeFlag_wsformulas, 4).Value2
    wsData.Range(wsData.Cells(2, col_ChangeFlag), wsData.Cells(intLastRow, col_ChangeFlag)).Formula = strFormula

End Sub
Sub o_24_Refresh_Pivots()

' Purpose: To refresh the pivot tables in the LOB Detail Review.
' Trigger: Called by o_01_MAIN_PROCEDURE_Update_Data_ws
' Updated: 12/3/2020

' Change Log:
'       10/13/2020: Original Creation
'       12/3/2020: Refreshed to mirror the approach for the ACH Past Due

' ****************************************************************************

' -----------
' Update the Named Range and refresh the pivot tables
' -----------
    
    'Update the Named Range
    ThisWorkbook.Names("DATA").RefersTo = "='Data'!$A$1:$CA$" & intLastRow
    
    'Refresh All Pivot Tables
    ThisWorkbook.RefreshAll

End Sub
Sub o_31_Over_5MM_No_Rating()

' Purpose: To flag any customers with $5MM+ Outstanding but no Overall Concern rating.
' Trigger: Called by o_01_MAIN_PROCEDURE_Update_Data_ws
' Updated: 5/13/2020

' Change Log:
'          5/13/2020: Intial Creation

' ****************************************************************************

' -----------
' Declare your variables
' -----------
    
    'Dim Integers
        Dim i As Integer
        
    'Dim "Ranges"
            
        'Dim col_Outstanding as Integer
            col_Outstanding = fx_Create_Headers("Direct Outstanding", arryHeader)
            
        'Dim col_OverallRisk As Integer
            'col_OverallRisk = fx_Create_Headers("OVERALL CONCERN", arryHeader)
            
' -----------
' Run the loop to flag the over $5MM customers w/out risk ratings
' -----------

    With wsData
        For i = 2 To intLastRow
            If .Cells(i, col_Outstanding).Value >= 5000000 Then
                If .Cells(i, col_OverallRisk) = "" Then
                    .Cells(i, col_OverallRisk).Interior.Color = intOrange
                Else
                    .Cells(i, col_OverallRisk).Interior.Color = xlNone
                End If
            Else
                .Cells(i, col_OverallRisk).Interior.Color = xlNone
            End If
            
        Next i
    End With

End Sub
Sub o_32_Relief_But_No_Mod()

' Purpose: To flag any customers with relief requested, but no mod listed.
' Trigger: Called by o_01_MAIN_PROCEDURE_Update_Data_ws
' Updated: 5/13/2020

' Change Log:
'          5/13/2020: Intial Creation

' ****************************************************************************

' -----------
' Declare your variables
' -----------
    
    'Dim Integers
        Dim i As Integer
        
    'Dim "Ranges"
            
        'Dim col_Relief As Integer
            'col_Relief = fx_Create_Headers("1st Round Relief", arryHeader)
            
        'Dim col_Relief2nd As Integer
            'col_Relief2nd = fx_Create_Headers("2nd Round Relief", arryHeader)
    
        'Dim col_ModType As Integer
            'col_ModType = fx_Create_Headers("1st Round Modification Type", arryHeader)
            
        'Dim col_ModType2nd As Integer
            'col_ModType2nd = fx_Create_Headers("2nd Round Modification Type", arryHeader)
            
' -----------
' Run the loop to flag the over $5MM customers w/out risk ratings
' -----------
    
    With wsData
        For i = 2 To intLastRow
            
            If Left(.Cells(i, col_Relief), 1) = "Y" Then 'First Relief
                If .Cells(i, col_ModType) = "" Or .Cells(i, col_ModType) = "None" Then
                    .Cells(i, col_ModType).Interior.Color = intOrange
                Else
                    .Cells(i, col_ModType).Interior.Color = xlNone
                End If
            Else
                .Cells(i, col_ModType).Interior.Color = xlNone
            End If
            
            If Left(.Cells(i, col_Relief2nd), 1) = "Y" Then 'Second Relief
                If .Cells(i, col_ModType2nd) = "" Or .Cells(i, col_ModType2nd) = "None" Then
                    .Cells(i, col_ModType2nd).Interior.Color = intOrange
                Else
                    .Cells(i, col_ModType2nd).Interior.Color = xlNone
                End If
            Else
                .Cells(i, col_ModType2nd).Interior.Color = xlNone
            End If
            
        Next i
    End With

End Sub
Sub o_33_Q2_Rating_Refresh_Not_Completed()

' Purpose: To flag any customers that are due for a risk rating refresh but that have not completed the field yet.
' Trigger: Called by o_01_MAIN_PROCEDURE_Update_Data_ws
' Updated: 5/14/2020

' Change Log:
'          5/14/2020: Intial Creation

' ****************************************************************************

' -----------
' Declare your variables
' -----------
    
    'Dim Integers
        Dim i As Integer
        
    'Dim "Ranges"
            
        Dim col_2Q_Refresh_Req As Integer
            col_2Q_Refresh_Req = fx_Create_Headers("2Q Rating Refresh Req.", arryHeader)
            
        Dim col_2Q_Refresh_Complete As Integer
            col_2Q_Refresh_Complete = fx_Create_Headers("2Q Rating Refresh Completed", arryHeader)
            
' -----------
' Run the loop to flag customers that require a refresh, but the PM never updated the complete field
' -----------

    With wsData
        For i = 2 To intLastRow
            If .Cells(i, col_2Q_Refresh_Req) = "Y" And .Cells(i, col_2Q_Refresh_Complete) = "" Then
                .Cells(i, col_2Q_Refresh_Complete).Interior.Color = intOrange
            Else
                .Cells(i, col_2Q_Refresh_Complete).Interior.Color = xlNone
            End If
            
        Next i
    End With

End Sub
Sub o_34_Missing_End_Market()

' Purpose: To flag any customers that should have an End Market but don't currently.
' Trigger: Called by o_01_MAIN_PROCEDURE_Update_Data_ws
' Updated: 5/27/2020

' Change Log:
'          5/27/2020: Intial Creation

' ****************************************************************************

' -----------
' Declare your variables
' -----------
    
    'Dim Integers
        Dim i As Integer
        
    'Dim "Ranges"
            
        Dim col_End_Market As Integer
            col_End_Market = fx_Create_Headers("End Market", arryHeader)
            
        Dim col_Prop_Type As Integer
            col_Prop_Type = fx_Create_Headers("Property Type", arryHeader)
            
        Dim col_LOB As Integer
            col_LOB = fx_Create_Headers("Line of Business", arryHeader)
            
        'Dim col_Outstanding as Integer
            'col_Outstanding = fx_Create_Headers("Direct Outstanding", arryHeader)
            
' -----------
' Run the loop to flag customers that should have an End Market (anyone excluding CRE) but don't
' -----------

    With wsData
        For i = 2 To intLastRow
            If .Cells(i, col_End_Market) = "" _
            And .Cells(i, col_LOB) <> "Commercial Real Estate" _
            And .Cells(i, col_Prop_Type) = "" _
            And .Cells(i, col_Outstanding) > 0 Then
                .Cells(i, col_End_Market).Interior.Color = intOrange
            Else
                .Cells(i, col_End_Market).Interior.Color = xlNone
            End If
            
        Next i
    End With

End Sub
Sub o_35_Missing_Property_Type()

' Purpose: To flag any customers that should have a Property Type but don't currently.
' Trigger: Called by o_01_MAIN_PROCEDURE_Update_Data_ws
' Updated: 5/27/2020

' Change Log:
'          5/27/2020: Intial Creation

' ****************************************************************************

' -----------
' Declare your variables
' -----------
    
    'Dim Integers
        Dim i As Integer
        
    'Dim "Ranges"
            
        Dim col_Prop_Type As Integer
            col_Prop_Type = fx_Create_Headers("Property Type", arryHeader)
            
        Dim col_LOB As Integer
            col_LOB = fx_Create_Headers("Line of Business", arryHeader)
            
        'Dim col_Outstanding as Integer
            'col_Outstanding = fx_Create_Headers("Direct Outstanding", arryHeader)
            
' -----------
' Run the loop to flag customers that should have a Property Type (CRE) but don't
' -----------

    With wsData
        For i = 2 To intLastRow
            If .Cells(i, col_Prop_Type) = "" And .Cells(i, col_LOB) = "Commercial Real Estate" And .Cells(i, col_Outstanding) > 0 Then
                .Cells(i, col_Prop_Type).Interior.Color = intOrange
            Else
                .Cells(i, col_Prop_Type).Interior.Color = xlNone
            End If
            
        Next i
    End With

End Sub
Sub o_36_2nd_Round_Relief_Only()

' Purpose: To flag any customers that have a 2nd Round Relief but no 1st Round Relief.
' Trigger: Called by o_01_MAIN_PROCEDURE_Update_Data_ws
' Updated: 6/8/2020

' Change Log:
'          6/8/2020: Intial Creation

' ****************************************************************************

' -----------
' Declare your variables
' -----------
    
    'Dim Integers
        Dim i As Integer
        
    'Dim "Ranges"
            
        'Dim col_Relief As Integer
            'col_Relief = fx_Create_Headers("1st Round Relief", arryHeader)
            
        'Dim col_Relief2nd As Integer
            'col_Relief2nd = fx_Create_Headers("2nd Round Relief", arryHeader)
            
' -----------
' Run the loop to flag customers that ONLY have 2nd Round Relief but NOT 1st Round Relief
' -----------

    With wsData
        For i = 2 To intLastRow
            If Left(.Cells(i, col_Relief2nd), 1) = "Y" And Left(.Cells(i, col_Relief), 1) <> "Y" Then
                .Cells(i, col_Relief2nd).Interior.Color = intOrange
            Else
                .Cells(i, col_Relief2nd).Interior.Color = xlNone
            End If
            
        Next i
    End With
         
End Sub
Sub o_37_Missing_Mod_Maturity_Date()

' Purpose: To flag any customers that don't have a Mod Maturity Date.
' Trigger: Called by o_01_MAIN_PROCEDURE_Update_Data_ws
' Updated: 6/8/2020

' Change Log:
'          6/8/2020: Intial Creation
'          6/8/2020: Updated to only include relief granted

' ****************************************************************************

' -----------
' Declare your variables
' -----------
    
    'Dim Integers
        Dim i As Integer
        
    'Dim "Ranges"
            
        'Dim col_Relief As Integer
            'col_Relief = fx_Create_Headers("1st Round Relief", arryHeader)
            
        'Dim col_Relief2nd As Integer
            'col_Relief2nd = fx_Create_Headers("2nd Round Relief", arryHeader)

        Dim col_ModMaturityDate1st As Integer
            col_ModMaturityDate1st = fx_Create_Headers("1st Round Mod Maturity Date", arryHeader)

        Dim col_ModMaturityDate2nd As Integer
            col_ModMaturityDate2nd = fx_Create_Headers("2nd Round Mod Maturity Date", arryHeader)

' -----------
' Run the loop to flag customers that have had a mod but don't have a Mod Maturity Date
' -----------

    With wsData
        For i = 2 To intLastRow
            If .Cells(i, col_Relief) = "Y - Granted" And .Cells(i, col_ModMaturityDate1st) = "" Then '1st Relief
                .Cells(i, col_ModMaturityDate1st).Interior.Color = intOrange
            Else
                .Cells(i, col_ModMaturityDate1st).Interior.Color = xlNone
            End If
            
            If .Cells(i, col_Relief2nd) = "Y - Granted" And .Cells(i, col_ModMaturityDate2nd) = "" Then '2nd Relief
                .Cells(i, col_ModMaturityDate2nd).Interior.Color = intOrange
            Else
                .Cells(i, col_ModMaturityDate2nd).Interior.Color = xlNone
            End If
            
        Next i
    End With

End Sub


