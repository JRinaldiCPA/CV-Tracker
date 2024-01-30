VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_CV_Tracker_Regular 
   Caption         =   "CMML CV Tracker - Customer Selector UserForm"
   ClientHeight    =   8160
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   14412
   OleObjectBlob   =   "uf_CV_Tracker_Regular.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_CV_Tracker_Regular"
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
    
'Dim Arrays / Other
    Dim bolPrivilegedUser As Boolean
    
    Dim arryHeader()

    Dim ary_Customers
    Dim ary_SelectedCustomers
    Dim ary_LOB_Customers
    Dim ary_Market_Customers
Private Sub UserForm_Initialize()
' ****************************************************************************
'
' Author:   James Rinaldi
' Created:  3/23/2020
' Updated:  8/19/2020
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
'          3/23/2020: Intial Creation
'          8/19/2020: Added the logic to exclude the exempt customers
'          9/2/2020: Added an additional filter to remove the blank fields at the bottom
'
' ****************************************************************************

Call Me.o_02_Assign_Private_Variables

' -----------
' Initialize the initial values
' -----------
    Me.StartUpPosition = 0 'Allow you to set the position
        Me.Top = Application.Top + (Application.UsableHeight / 1.5) - (Me.Height / 2) 'Open near the bottom of the screen
        Me.Left = Application.Left + (Application.UsableWidth / 2) - (Me.Width / 2)
    
    'Add the values for the LOB ListBox
        Me.lst_LOB.List = Get_LOB_Array

    'If the AutoFilter isn't on already then turn it on
        If wsData.AutoFilterMode = False Then
            wsData.Range("A:" & strLastCol_wsData).AutoFilter
        End If

    'Hide the values for the exempt customers (paid off / sold / written off
        Call o_41_Clear_Filters

' -----------
' Hide the worksheets and objects that shouldn't be seen unless it's an exception user
' -----------

    If bolPrivilegedUser = False Then
        Call Me.o_61_Protect_Ws
        Call Me.o_63_Hide_Worksheets
    ElseIf bolPrivilegedUser = True Then
        Call Me.o_62_UnProtect_Ws
        Call Me.o_63_Hide_Worksheets
    End If
    
    Call Me.o_11_Create_Customer_List_Dynamic
    
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
Private Sub lst_LOB_Click()
    
    Call Me.o_12_Create_Customer_List_By_LOB
    
    Call Me.o_14_Create_Market_List
    
    Me.cmb_DynamicSearch.Value = Null
        Me.cmb_DynamicSearch.SetFocus
    
End Sub
Private Sub lst_LOB_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Call Me.o_32_Filter_Customers_by_LOB
    
    Me.cmb_DynamicSearch.Value = Null
        Me.cmb_DynamicSearch.SetFocus

End Sub
Private Sub lst_Market_Click()
        
    Call Me.o_13_Create_Customer_List_By_Market

    Me.cmb_DynamicSearch.Value = Null
        Me.cmb_DynamicSearch.SetFocus

End Sub
Private Sub lst_Market_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Call Me.o_33_Filter_Customers_by_Market
    
    Me.cmb_DynamicSearch.Value = Null
        Me.cmb_DynamicSearch.SetFocus

End Sub
Private Sub cmb_DynamicSearch_Change()

    Call Me.o_11_Create_Customer_List_Dynamic

End Sub
Private Sub lst_Customers_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
      
    'If there is only one valid record select it instead of the blank, otherwise go w/ what I picked
      
    If lst_Customers.ListCount = 2 Then lst_Customers.Selected(0) = True
        If lst_Customers.Value = "" Then Exit Sub
        
    Call Me.o_21_Add_Customer_To_Selected_Customers_Array
    
    Call Me.o_31_Filter_Customers
End Sub
Private Sub lst_Customers_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    'If I hit enter and have selected a customer filter the list to JUST that customer
    
    If lst_Customers.ListCount = 2 Then lst_Customers.Selected(0) = True
    
    If KeyCode = vbKeyReturn And Me.lst_Customers.Value <> "" Then
        Call Me.o_23_Filter_Single_Customer
            Me.cmb_DynamicSearch.SetFocus
    End If

End Sub
Private Sub lst_Customers_Enter()

    'If there is only one valid record select it, otherwise abort

    If lst_Customers.ListCount = 2 Then lst_Customers.Selected(0) = True
        If lst_Customers.Value = "" Then Exit Sub
        If lst_Customers.ListCount > 2 Then Exit Sub
        
    Call Me.o_21_Add_Customer_To_Selected_Customers_Array
    
    Call Me.o_31_Filter_Customers

End Sub
Private Sub cmd_Filter_Customers_Click()
  
    Call Me.o_31_Filter_Customers

    Call Me.o_34_Save_SelectedCustomer_Array_Values

End Sub
Private Sub cmd_Filter_Customers_by_LOB_Click()
    
    If Me.lst_Market <> "" Then
        Call Me.o_33_Filter_Customers_by_Market
    Else
        Call Me.o_32_Filter_Customers_by_LOB
    End If

End Sub
Private Sub cmd_Clear_Filter_Click()

    Call Me.o_41_Clear_Filters
    
    Call Me.o_42_Clear_Saved_Array

    Me.lst_LOB.Value = ""
    Me.lst_Market.Clear
    Me.lst_Customers.Clear
    Me.cmb_DynamicSearch.Value = Null

End Sub
Private Sub cmd_Email_Credit_Risk_Click()

    Call Me.o_71_Create_Faux_Change_Log
    
    Call Me.o_35_Filter_Only_Changes
    
    Call Me.o_51_Create_a_XLSX_Copy

    Call Me.o_52_Email_Credit_Risk
       
    Call PrivateMacros.DisableForEfficiencyOff
    
    Unload Me

End Sub
Private Sub cmd_Cancel_Click()
    
    Unload Me

End Sub
Private Sub cmd_Filter_Anomalies_Click()
    
    Call Me.o_37_Filter_Edits_For_PMs
    
End Sub
Sub o_11_Create_Customer_List_Dynamic()

' Purpose: To create the list of customers to be used in the Customer ListBox.
' Trigger: Start typing in the DynamicSearch combo box
' Updated: 10/13/2020

' Change Log:
'       3/23/2020: Intial Creation
'       10/13/2020: Updated to compare the Market as well

' ****************************************************************************

On Error GoTo ErrorHandler

' -----------
' Declare your variables
' -----------
    
    Dim CustomerName As String

    Dim strLOB As String
    
    Dim strMarket As String

    Dim x As Integer
    
    Dim y As Integer
        y = 1
    
    'Dim ary_Customers As Variant
        ReDim ary_Customers(1 To 9999)
    
' -----------
' Run the loop
' -----------
       
    Me.lst_Customers.Clear
       
    With wsLists
            x = 2
        Do While .Range("E" & x).Value2 <> ""
            
            strLOB = .Range("C" & x).Value2
            strMarket = .Range("D" & x).Value2
            CustomerName = .Range("E" & x).Value2

                If InStr(1, CustomerName, Me.cmb_DynamicSearch.Value, vbTextCompare) Then
                    
                    If IsNull(Me.lst_LOB) Or Me.lst_LOB = strLOB Then
                        If IsNull(Me.lst_Market) Or Me.lst_Market = "" Or Me.lst_Market = strMarket Then
                            ary_Customers(y) = CustomerName
                            y = y + 1
                        End If
                    End If
        
                End If
            
            x = x + 1
        Loop
    End With

    ReDim Preserve ary_Customers(1 To y)

    Me.lst_Customers.List = ary_Customers

Exit Sub

ErrorHandler:

Global_Error_Handling SubName:="o_11_Create_Customer_List_Dynamic", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_12_Create_Customer_List_By_LOB()

' Purpose: To create the list of customers to be used in the Customer ListBox, based on the selected LOB.
' Trigger: Select a customer from the LOB ListBox.
' Updated: 3/23/2020

' Change Log:
'   3/23/2020: Intial Creation

' ****************************************************************************

On Error GoTo ErrorHandler

' -----------
' Declare your variables
' -----------
    
    Dim str_LOB_Selected As String
        str_LOB_Selected = Me.lst_LOB.Value
    
    'Dim ary_LOB_Customers As Variant
        ReDim ary_LOB_Customers(1 To 99999)
    
    Dim x As Integer
    
    Dim y As Integer
        y = 1
    
' -----------
' Run your code your variables
' -----------
    
    Me.lst_Customers.Clear
    
    With wsLists
            x = 2
        Do While .Range("C" & x).Value2 <> ""
            
            LOB = .Range("C" & x).Value2
            CustomerName = .Range("E" & x).Value2
                                    
                If Me.lst_LOB = LOB Then
                    ary_LOB_Customers(y) = CustomerName
                    y = y + 1
                End If

            x = x + 1
        Loop
    End With

    ReDim Preserve ary_LOB_Customers(1 To y)

    Me.lst_Customers.List = ary_LOB_Customers
    
Exit Sub

ErrorHandler:

Global_Error_Handling SubName:="o_12_Create_Customer_List_By_LOB", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    
End Sub
Sub o_13_Create_Customer_List_By_Market()

' Purpose: To create the list of customers to be used in the Customer ListBox, based on the selected Market.
' Trigger: Select a customer from the Market ListBox.
' Updated: 3/30/2020

' Change Log:
'   3/30/2020: Intial Creation

' ****************************************************************************

' -----------
' Declare your variables
' -----------
    
    Dim str_LOB_Selected As String
        str_LOB_Selected = Me.lst_LOB.Value
    
    Dim str_Market_Selected As String
        str_Market_Selected = Me.lst_Market.Value
        
    'Dim ary_Market_Customers As Variant
        ReDim ary_Market_Customers(1 To 99999)
    
    Dim x As Integer
    
    Dim y As Integer
        y = 1
    
    Dim strLOB As String
    
    Dim strMarket As String
    
' -----------
' Run your code your variables
' -----------
    
    Me.lst_Customers.Clear
    
    With wsLists
            x = 2
        Do While .Range("E" & x).Value2 <> ""
            
            strLOB = .Range("C" & x).Value2
            strMarket = .Range("D" & x).Value2
            CustomerName = .Range("E" & x).Value2
                                    
                If Me.lst_LOB = strLOB And Me.lst_Market = strMarket Then
                    ary_Market_Customers(y) = CustomerName
                    y = y + 1
                End If

            x = x + 1
        Loop
    End With

    ReDim Preserve ary_Market_Customers(1 To y)

    Me.lst_Customers.List = ary_Market_Customers
    
Exit Sub

End Sub
Sub o_14_Create_Market_List()
   
' Purpose: To create the list of Markets to be used in the Market ListBox, based on the selected LOB.
' Trigger: Select a customer from the LOB ListBox.
' Updated: 3/30/2020

' Change Log:
'   3/30/2020: Intial Creation

' ****************************************************************************
        
' -----------
' Declare your variables
' -----------
    
    Dim x As Integer
        x = 2
    
' -----------
' Run the loop
' -----------

        lst_Market.Clear
    
        With wsLists
            Do While .Range("L" & x).Value2 <> ""
                If Me.lst_LOB = .Range("K" & x).Value2 Then lst_Market.AddItem (.Range("L" & x))
                x = x + 1
            Loop
        End With
              
End Sub
Sub o_21_Add_Customer_To_Selected_Customers_Array()

' Purpose: To add to a growing array of customers that will then be used to filter the data.
' Trigger: Double Click Customer
' Updated: 3/23/2020

' Change Log:
'   3/23/2020: Intial Creation

' ****************************************************************************

On Error GoTo ErrorHandler

ReDim Preserve ary_SelectedCustomers(1 To 9999)

' -----------
' Declare your variables
' -----------
    
    'Dim strCustomerName As String
        strCustomerName = lst_Customers.Value

    Dim i As Integer
        i = 1

    Dim intArrayLast As Integer

        Do Until ary_SelectedCustomers(i) = Empty
            intArrayLast = i
            i = i + 1
        Loop

        If ary_SelectedCustomers(1) = Empty Then intArrayLast = 0

' -----------
' Input the value into the array in the first empty slot
' -----------

    ary_SelectedCustomers(intArrayLast + 1) = strCustomerName

Exit Sub

ErrorHandler:

    Global_Error_Handling SubName:="o_21_Add_Customer_To_Selected_Customers_Array", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_23_Filter_Single_Customer()

' Purpose: To filter the list of customers in the Data ws based on only the currently swelected customer in the Customers List.
' Trigger: Called: uf_CV_Tracker_Regular
' Updated: 8/25/2020

' Change Log:
'          4/22/2020: Intial Creation
'          8/25/2020: Added error handling if a blank was selected instead of a customer

' ****************************************************************************

If Me.lst_Customers.Value = "" Then Exit Sub

Call Me.o_62_UnProtect_Ws

    With wsData

      .AutoFilterMode = False
          .Range("A:" & strLastCol_wsData).AutoFilter Field:=col_Customer, Criteria1:=Me.lst_Customers.Value, Operator:=xlFilterValues
    
    End With

Call Me.o_61_Protect_Ws

End Sub
Sub o_31_Filter_Customers()

' Purpose: To filter the list of customers in the Data ws based on the Customer Selected Array.
' Trigger: Called: uf_CV_Tracker_Regular.cmd_Filter_Customers
' Updated: 3/23/2020

' Change Log:
'   3/23/2020: Intial Creation

' ****************************************************************************

Call PrivateMacros.DisableForEfficiency

On Error GoTo ErrorHandler

Call Me.o_62_UnProtect_Ws

' -----------
' Filter the Summary worksheet based on the customers selected previously
' -----------
    With wsData

      .AutoFilterMode = False
          .Range("A:" & strLastCol_wsData).AutoFilter Field:=col_Customer, Criteria1:=ary_SelectedCustomers, Operator:=xlFilterValues
    
    End With

    Application.GoTo Reference:=wsData.Range("A1"), Scroll:=True
    
Call Me.o_61_Protect_Ws

Call PrivateMacros.DisableForEfficiencyOff

Exit Sub

ErrorHandler:

    Global_Error_Handling SubName:="o_31_Filter_Customers", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    
    Call Me.o_61_Protect_Ws
    
    Call PrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_32_Filter_Customers_by_LOB()

' Purpose: To filter the list of customers in the Data ws based on the selected LOB.
' Trigger: Called: uf_CV_Tracker_Regular.cmd_Filter_Customers_by_LOB
' Updated: 3/23/2020

' Change Log:
'   3/23/2020: Intial Creation

' ****************************************************************************

Call PrivateMacros.DisableForEfficiency

On Error GoTo ErrorHandler

Call Me.o_62_UnProtect_Ws
        
Dim aryTemp As Variant
    If Me.lst_LOB.Value = "Restructure & Recovery" Then aryTemp = Application.Transpose(Me.lst_Customers.List)
        
' -----------
' Filter the Summary worksheet based on the customers selected previously
' -----------
    With wsData

      .AutoFilterMode = False
          .Range("A:" & strLastCol_wsData).AutoFilter Field:=col_LOB, Criteria1:=Me.lst_LOB.Value, Operator:=xlFilterValues
    End With

    If Me.lst_LOB.Value = "Restructure & Recovery" Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_LOB
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_Customer, Criteria1:=aryTemp, Operator:=xlFilterValues
    End If

    Application.GoTo Reference:=wsData.Range("A1"), Scroll:=True
    
Call Me.o_61_Protect_Ws

Call PrivateMacros.DisableForEfficiencyOff

Exit Sub

ErrorHandler:

    Global_Error_Handling SubName:="o_32_Filter_Customers_by_LOB", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    
    Call Me.o_61_Protect_Ws
    
    Call PrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_33_Filter_Customers_by_Market()

' Purpose: To filter the list of customers in the Data ws based on the selected Market.
' Trigger: Called: uf_CV_Tracker_Regular.cmd_Filter_Customers_by_LOB
' Updated: 3/23/2020

' Change Log:
'   3/23/2020: Intial Creation

' ****************************************************************************

Call PrivateMacros.DisableForEfficiency

On Error GoTo ErrorHandler

Call Me.o_62_UnProtect_Ws

' -----------
' Filter the Summary worksheet based on the customers selected previously
' -----------
    wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_Market, Criteria1:=Me.lst_Market.Value, Operator:=xlFilterValues

    Application.GoTo Reference:=wsData.Range("A1"), Scroll:=True
    
Call Me.o_61_Protect_Ws

Call PrivateMacros.DisableForEfficiencyOff

Exit Sub

ErrorHandler:

    Global_Error_Handling SubName:="o_33_Filter_Customers_by_Market", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    
    Call Me.o_61_Protect_Ws
    
    Call PrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_34_Save_SelectedCustomer_Array_Values()

' Purpose: To save the selected customers so that they can be accessed later.
' Trigger: Called: uf_CV_Tracker_Regular.cmd_Filter_Customers_Click
' Updated: 3/23/2020

' Change Log:
'   3/23/2020: Intial Creation

' ****************************************************************************

' -----------
' Copy the values from the array into the worksheet
' -----------

On Error Resume Next

    wsArrays.Range("A2:A" & UBound(ary_SelectedCustomers) + 1) = WorksheetFunction.Transpose(ary_SelectedCustomers)

End Sub
Sub o_35_Filter_Only_Changes()

' Purpose: To filter the Data ws down to only records that had a Change before emailing Credit Risk.
' Trigger: Called: cmd_Email_Credit_Risk_Click
' Updated: 3/23/2020

' Change Log:
'   3/23/2020: Intial Creation

' ****************************************************************************

Call PrivateMacros.DisableForEfficiency

On Error GoTo ErrorHandler

Call Me.o_62_UnProtect_Ws
       
' -----------
' Filter the Summary worksheet based on the customers selected previously
' -----------
    With wsData

      .AutoFilterMode = False
          .Range("A:" & strLastCol_wsData).AutoFilter Field:=col_ChangeFlag, Criteria1:="CHANGE", Operator:=xlFilterValues
    
    End With

    Application.GoTo Reference:=wsData.Range("A1"), Scroll:=True
    
Call Me.o_61_Protect_Ws

Call PrivateMacros.DisableForEfficiencyOff

Exit Sub

ErrorHandler:

    Global_Error_Handling SubName:="o_35_Filter_Only_Changes", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    
    Call Me.o_61_Protect_Ws
    
    Call PrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_37_Filter_Edits_For_PMs()

' Purpose: To filter the Data down to just the Edits that still need to be addressed by the PMs.
' Trigger: Called: cmd_Filter_Anomalies
' Updated: 11/23/2020

' Change Log:
'          11/23/2020: Intial Creation
    
' ****************************************************************************

On Error GoTo ErrorHandler

PrivateMacros.DisableForEfficiency

Call Me.o_62_UnProtect_Ws

' -----------
' Declare your variables
' -----------

    'Dim "Ranges"
    
    Dim col_EditFlag As Integer
        col_EditFlag = fx_Create_Headers("Edit Flag", arryHeader)
        
    Dim rngData As Range
        Set rngData = wsData.Range(wsData.Cells(2, 1), wsData.Cells(intLastRow, intLastCol))

    Dim cell As Variant

    ' Dim Colors

    Dim intOrange As Long
        intOrange = RGB(253, 233, 217)
        
' -----------
' Refresh the Filter Flag data
' -----------
    
    'Clear out the old data
    wsData.Range(wsData.Cells(2, col_EditFlag), wsData.Cells(intLastRow, col_EditFlag)).ClearContents

    For Each cell In rngData.SpecialCells(xlCellTypeVisible) 'Visible only to account for filtered data
        If cell.Interior.Color = intOrange Then
            wsData.Cells(cell.Row, col_EditFlag).Value2 = "Yes"
        End If
    Next cell
    
' -----------
' Filter the data
' -----------
          
    If bol_Edit_Filter = False Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_EditFlag, Criteria1:="Yes", Operator:=xlFilterValues
        Me.cmd_Filter_Anomalies.BackColor = RGB(240, 248, 224)
        Me.cmd_Filter_Anomalies.Caption = "Anomalies Only"
    ElseIf bol_Edit_Filter = True Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_EditFlag
        Me.cmd_Filter_Anomalies.BackColor = RGB(240, 240, 240)
        Me.cmd_Filter_Anomalies.Caption = "Filter Anomalies"
    End If
    
    bol_Edit_Filter = Not bol_Edit_Filter 'Switch the boolean

    Call Me.o_61_Protect_Ws

PrivateMacros.DisableForEfficiencyOff

Exit Sub

' -----------
' Error Handler
' -----------

ErrorHandler:

    Global_Error_Handling SubName:="o_37_Filter_Edits_For_PMs", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description
    Call Me.o_61_Protect_Ws
    PrivateMacros.DisableForEfficiencyOff

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
Sub o_51_Create_a_XLSX_Copy()

On Error GoTo ErrorHandler

' Purpose: To create a copy of the workbook in XLSX to aid in providing it to Lizzy.
' Trigger: Called: uf_CV_Tracker_Regular
' Updated: 3/31/2020

' Change Log:
'   3/31/2020: Updated the strOldFileName to include TEMP and look simpler

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    Dim strFullName As String
        strFullName = Functions.Name_Reverse()

    Dim objFSO As Object
        Set objFSO = VBA.CreateObject("Scripting.FileSystemObject")

    'Public strNewFileFullPath As String
        strNewFileFullPath = ThisWorkbook.Path & "\" & Format(Now, "mm-dd") & " " & Format(Now, "HH-MM-SS") & " COVID-19 Update by " & strFullName

    Dim strOldFileName As String
        strOldFileName = "CV TRACKER - TEMP" & "(" & Format(Now, "mm-dd") & " " & Format(Now, "HH-MM-SS") & ")" & ".xlsm"
    
    Dim strOldFileFullPath As String
        strOldFileFullPath = ThisWorkbook.Path & "\" & strOldFileName

' -----------
' Create a copy of the workbook using the current date, time, and the individuals name
' -----------

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
        
        ThisWorkbook.Save
        ThisWorkbook.SaveCopyAs strOldFileFullPath
        
        Workbooks.Open strOldFileFullPath
        
        Workbooks(strOldFileName).SaveAs _
            Filename:=strNewFileFullPath, FileFormat:=xlOpenXMLWorkbook
        
        ActiveWorkbook.Close
                
    On Error Resume Next
        If Right(strOldFileFullPath, 5) = ".xlsm" Then Kill strOldFileFullPath
    On Error GoTo 0
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

Debug.Print ThisWorkbook.FullName 'Take the full path incase something goes awry

Exit Sub

ErrorHandler:

Global_Error_Handling SubName:="o_51_Create_a_XLSX_Copy", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

End Sub
Sub o_52_Email_Credit_Risk()

On Error GoTo ErrorHandler

' Purpose: To attach the template to an email to send to Credit Reporting & Analysis.
' Trigger: Called: uf_CV_Tracker_Regular
' Updated: 4/30/2021

' Change Log:

'       3/23/2020: Initial Creation
'       9/25/2020: Updated the strTempVersion to be a length of 5, and added another space for the diagnostic info
'       4/30/2021: Added Axcel's version of the spell checker

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    Dim OutApp As Object
    Dim OutMail As Object
    Dim strbody As String

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    Dim strTempVersion As String
        strTempVersion = Mid(String:=ThisWorkbook.Name, Start:=InStr(ThisWorkbook.Name, " (v") + 3, Length:=6)

' -----------
' Send the Save Me email
' -----------

    On Error Resume Next
    
        strbody = "James," & vbNewLine & vbNewLine & _
            "This is an automated notification of a new COVID-19 related risk rating change to include in the tracker." & vbNewLine & vbNewLine & _
            "Thanks," & vbNewLine & _
            "The CV Tracker" & vbNewLine & vbNewLine & _
            "Diagnostic Info:" & vbNewLine & vbNewLine & _
            "CV Tracker Version: " & strTempVersion & vbNewLine & _
            "User: " & strUserID & vbNewLine & _
            "File Path: " & ThisWorkbook.FullName
        
        With OutMail
            .CC = "JRinaldi@WebsterBank.com"
            '.CC = "JRinaldi@WebsterBank.com; aduarteespinoza@WebsterBank.com"
            .Subject = "COVID-19 Tracker Update - Auto Email"
            .Body = strbody
            .Display
            .Attachments.Add strNewFileFullPath & ".xlsx"
            '    Application.SendKeys "%s"
            
            '   3/31/2021: Added extra code to avoid spell corrector
            'Send Key: Send
            Application.SendKeys "%{s}", True
            'Send Key: Cancel/Escape Spell Checking
            Application.SendKeys "{ESC}", True
            'Send Key: Yes
            Application.SendKeys "%{y}", True
            
        End With
        
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing

    MsgBox _
    Title:="It worked!", _
    Buttons:=vbInformation, _
    Prompt:="Your email has been sent, and the CV Tracker was saved, you may now exit the CV Tracker. " & Chr(10) & Chr(10) _
    & "If you want to make additional changes please open the Excel workbook called 'CV Tracker (v XX.X)' not the TEMP file with your name in the title."

Exit Sub

ErrorHandler:

    Global_Error_Handling SubName:="o_52_Email_Credit_Risk", ErrSource:=Err.Source, ErrNum:=Err.Number, ErrDesc:=Err.Description

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
Sub o_71_Create_Faux_Change_Log()

' Purpose: To create a change log based on actual changes made, but not recorded in the "real" Change Log.
' Trigger: Manual
' Updated: 3/27/2020

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    ' Dim Integers
    
    Dim CurRowLog As Long
        CurRowLog = [MATCH(TRUE,INDEX(ISBLANK('Change Log'!A:A),0),0)]

    Dim rowID As Integer
    
    Dim colID As Integer
    
    ' Dim Strings / Ranges
    
    Dim strOldValue As String
    
    Dim strNewValue As String
 
    Dim strFullName As String
        strFullName = Functions.Name_Reverse()
 
    Dim strDateTime As String
        strDateTime = Format(Now, "m/d/yyyy hh:mm:ss")
 
' -----------
' Capture the data
' -----------
           
    For rowID = 2 To intLastRow_wsData
        With wsData
            If .Rows(rowID).EntireRow.Hidden = False Then 'If the row is visible
                If .Cells(rowID, col_ChangeFlag).Value2 = "CHANGE" Then 'If the record was flagged as having changes
                    
                    For colID = col_NewFirst To col_NewLast
                        If .Cells(rowID, colID).Value2 <> .Cells(rowID, colID).Offset(0, col_Offset).Value2 Then
                        
                            ' Copy the data into the change log
                            
                            With wsChangeLog
                                .Range("A" & CurRowLog).Value2 = strDateTime                                                 ' Change Made Data
                                .Range("B" & CurRowLog).Value2 = strFullName                                                 ' By Who
                                .Range("C" & CurRowLog).Value2 = wsData.Range("B" & rowID)                                   ' Customer
                                .Range("D" & CurRowLog).Value2 = wsData.Cells(1, colID)                                      ' Field Changed
                                .Range("E" & CurRowLog).Value2 = wsData.Cells(rowID, colID).Offset(0, col_Offset).Value2  ' Current Value
                                .Range("F" & CurRowLog).Value2 = wsData.Cells(rowID, colID).Value2                           ' New Value
                                .Range("G" & CurRowLog).Value2 = "Faux Log"                                                  ' Source
                                CurRowLog = CurRowLog + 1
                           End With
                        
                        End If
                        
                    Next colID
                
                End If
            End If
        End With
    
    Next rowID

End Sub




