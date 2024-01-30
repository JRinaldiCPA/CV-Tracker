VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_Impact_Meeting 
   Caption         =   "COVID-19 Impact Meeting"
   ClientHeight    =   8868.001
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   21480
   OleObjectBlob   =   "uf_Impact_Meeting.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_Impact_Meeting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Dim Worksheets
    Dim wsData As Worksheet
    Dim wsLists As Worksheet
    Dim wsArrays As Worksheet

'Dim Strings
    Dim strCustomerName As String
    Dim strLastCol_wsData As String

'Dim Integers
    Dim intLastCol_wsData As Integer
    
    Dim intLastRow_wsData As Long
    Dim intLastRow_wsChangeLog As Long
    Dim intLastRow_wsUpdates As Long
    Dim intLastRow_wsArrays As Long
    
'Dim Arrays
    Dim ary_Customers
    Dim ary_SelectedCustomers
    Dim ary_LOB_Customers
    Dim ary_Market_Customers
    
'Dim Ranges
    Dim arryHeader()

'Dim "Ranges"
    Dim col_Customer As Integer
    Dim col_LOB As Integer
    Dim col_Market As Integer
    Dim col_RelMgr As Integer
    Dim col_Industry As Integer
    Dim col_PropType As Integer
    Dim col_EndMarket As Integer
    
    Dim col_BRG As Integer
    Dim col_FRG As Integer
    Dim col_CCRP As Integer
    Dim col_LFT As Integer
    
    Dim col_Outstanding As Integer
    Dim col_Exposure As Integer
    
    Dim col_SupplyChain As Integer
    Dim col_OverallRisk  As Integer
    
    Dim col_1st_Relief As Integer
    Dim col_2nd_Relief As Integer
    Dim col_3rd_Relief As Integer
    
    Dim col_1st_ModMaturity As Integer
    Dim col_2nd_ModMaturity As Integer
    Dim col_3rd_ModMaturity As Integer
    
    Dim col_ActiveMod As Integer
    Dim col_ActiveMod_Status As Integer
    Dim col_PaymentMod As Integer
    
    Dim col_MSLPInquiry As Integer
    Dim col_PPPRequest As Integer
     
    Dim col_Comments As Integer
    Dim col_ReliefComments As Integer
    
    Dim col_1st_PayMod As Integer
    Dim col_2nd_PayMod As Integer
    Dim col_3rd_PayMod As Integer
    
    Dim col_ChangeFlag As Integer
    Dim col_NewHighRisk As Integer
    Dim col_NewReliefRequest1st As Integer
    Dim col_NewReliefRequest2nd As Integer

    Dim col_NewFirst As Integer
    Dim col_NewLast As Integer
    Dim col_OldFirst As Integer
    Dim col_OldLast As Integer
    Dim col_Offset As Integer

'Dim Booleans
    Dim bol_ExceptionUser As Boolean
    
    Dim bol_HighRiskFilter As Boolean
    Dim bol_NewHighRisk As Boolean
    
    Dim bol_1st_ReliefGranted As Boolean
    Dim bol_NewReliefRequest As Boolean
    
    Dim bol_MSLPInquiry As Boolean
    Dim bol_PPP As Boolean

    Dim bol_1st_Relief As Boolean
    Dim bol_2nd_Relief As Boolean

'Dim "Booleans"
    Dim bol_ReliefStatus As String
    Dim bol_ActiveMod As String
    Dim bol_PaymentMod As String

Private Sub frm_Cust_Ratings_and_Commentary_Click()

End Sub

Private Sub txt_Agenda_Change()

End Sub

Private Sub UserForm_Initialize()

' Purpose:  To initialize the userform, including adding in the data from the arrays.
' Trigger:  Workbook Open
' Updated:  3/23/2020
' Author:   James Rinaldi

' Change Log:
'   3/23/2020: Intial Creation

' ****************************************************************************

Call Me.o_02_Assign_Private_Variables

'Call Me.o_03_Reset_UserForm

' -----------
' Initialize the initial values
' -----------
    
    Me.StartUpPosition = 0 'Allow you to set the position
        Me.Top = Application.Top + (Application.UsableHeight / 1.35) - (Me.Height / 2) 'Open near the bottom of the screen
        Me.Left = Application.Left + (Application.UsableWidth / 2) - (Me.Width / 2)
    
    ' Add values that weren't cleared from the Arrays ws
        If wsArrays.Range("A2").Value2 <> "" Then
            ary_SelectedCustomers = WorksheetFunction.Transpose(wsArrays.Range("A2:A" & intLastRow_wsArrays).Value2)
        End If
       
    ' Add the values for the LOB ListBox
        Me.lst_LOB.List = Get_LOB_Array

    ' If the AutoFilter isn't on already then turn it on
        If wsData.AutoFilterMode = False Then
            wsData.Range("A:" & strLastCol_wsData).AutoFilter
        End If

    ' Create the initial LOB Detail metrics
        Call Me.o_24_Add_LOB_Details

End Sub
Sub o_02_Assign_Private_Variables()

' Purpose: To declare all of the Public variables that were dimensioned "above the line".
' Trigger: Called by UserForm_Initialize Event
' Updated: 4/23/2020

' Change Log:
'          4/23/2020: Intial Creation

' ****************************************************************************

' -----------
' Declare your variables
' -----------
    
    'Dim sheets
    
        'Dim wsData As Worksheet
            Set wsData = ThisWorkbook.Sheets("Data")
    
        'Dim wsLists As Worksheet
            Set wsLists = ThisWorkbook.Sheets("Lists")
            
        'Dim wsArrays As Worksheet
            Set wsArrays = ThisWorkbook.Sheets("Array Values")

    'Dim Integers / Strings
    
        'Public intLastCol_wsData As Integer
            intLastCol_wsData = wsData.Cells(1, Columns.Count).End(xlToLeft).Column
            
        'Public intLastRow_wsData As Long
            intLastRow_wsData = [MATCH(TRUE,INDEX(ISBLANK('Data'!A:A),0),0)] - 1

        'Public strLastCol_wsData as String
            strLastCol_wsData = Split(Cells(1, intLastCol_wsData).Address, "$")(1)
            
        'Dim intLastRow_wsArrays As Long
            intLastRow_wsArrays = wsArrays.Cells(Rows.Count, "A").End(xlUp).Row
            If intLastRow_wsArrays = 1 Then intLastRow_wsArrays = 2

    'Dim Ranges
        
        'Dim arryHeader()
            arryHeader = Application.Transpose(wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, intLastCol_wsData)))

    'Dim "Ranges"
            
            col_Customer = fx_Create_Headers("Customer Name", arryHeader)
            col_LOB = fx_Create_Headers("Line of Business", arryHeader)
            col_Market = fx_Create_Headers("Market", arryHeader)
            col_RelMgr = fx_Create_Headers("Relationship Manager", arryHeader)
            col_Industry = fx_Create_Headers("Scorecard Industry", arryHeader)
            col_PropType = fx_Create_Headers("Property Type", arryHeader)
            col_EndMarket = fx_Create_Headers("End Market", arryHeader)
            
            col_BRG = fx_Create_Headers("BRG", arryHeader)
            col_FRG = fx_Create_Headers("FRG", arryHeader)
            col_CCRP = fx_Create_Headers("CCRP", arryHeader)
            col_LFT = fx_Create_Headers("LFT", arryHeader)
            
            col_Outstanding = fx_Create_Headers("Direct Outstanding", arryHeader)
            col_Exposure = fx_Create_Headers("Gross Exposure", arryHeader)
            
            col_SupplyChain = fx_Create_Headers("Supply Chain Concern", arryHeader)
            col_OverallRisk = fx_Create_Headers("OVERALL CONCERN", arryHeader)
            
            col_1st_Relief = fx_Create_Headers("1st Round Relief", arryHeader)
            col_2nd_Relief = fx_Create_Headers("2nd Round Relief", arryHeader)
            col_3rd_Relief = fx_Create_Headers("3rd+ Round Relief", arryHeader)
            
            col_1st_ModMaturity = fx_Create_Headers("1st Round Mod Maturity Date", arryHeader)
            col_2nd_ModMaturity = fx_Create_Headers("2nd Round Mod Maturity Date", arryHeader)
            col_3rd_ModMaturity = fx_Create_Headers("3rd+ Round Mod Maturity Date", arryHeader)
            
            col_ActiveMod = fx_Create_Headers("Active Mod", arryHeader)
            col_PaymentMod = fx_Create_Headers("Active Payment Mod", arryHeader)
            col_ActiveMod_Status = fx_Create_Headers("Active Mod Status", arryHeader)
            
            col_1st_PayMod = fx_Create_Headers("1st Payment Mod", arryHeader)
            col_2nd_PayMod = fx_Create_Headers("2nd Payment Mod", arryHeader)
            col_3rd_PayMod = fx_Create_Headers("3rd+ Payment Mod", arryHeader)
            
            col_MSLPInquiry = fx_Create_Headers("MSLP Inquiry?", arryHeader)
            col_PPPRequest = fx_Create_Headers("PPP", arryHeader)
            
            col_Comments = fx_Create_Headers("Comments", arryHeader)
            col_ReliefComments = fx_Create_Headers("Relief Comments", arryHeader)
            
            col_ChangeFlag = fx_Create_Headers("Change Flag", arryHeader)
            col_NewHighRisk = fx_Create_Headers("Date High Overall Risk", arryHeader)
            col_NewReliefRequest1st = fx_Create_Headers("Date Relief Requested", arryHeader)
            col_NewReliefRequest2nd = fx_Create_Headers("Date 2nd Relief Requested", arryHeader)
            
            col_NewFirst = col_PropType
            col_NewLast = col_ReliefComments
            col_Offset = col_NewLast - col_NewFirst + 1
            col_OldFirst = col_NewFirst + col_Offset
            col_OldLast = col_NewLast + col_Offset
    
    'Dim Arrays
            
        'Public ary_SelectedCustomers
            ReDim ary_SelectedCustomers(1)
    
        'Public ary_SelectedCustomers
            ReDim ary_Customers(1)

End Sub
Private Sub lst_LOB_Click()
    
    Call Me.o_12_Create_Customer_List_By_LOB
    
    Call Me.o_14_Create_Market_List
    
    Me.cmb_DynamicSearch.Value = Null
        Me.cmb_DynamicSearch.SetFocus
    
End Sub
Private Sub lst_LOB_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Call Me.o_44_Clear_Customer_Filter
    
    Call Me.o_32_Filter_Customers_by_LOB
    
    Call Me.o_24_Add_LOB_Details
    
    Me.cmb_DynamicSearch.Value = Null
        Me.cmb_DynamicSearch.SetFocus

End Sub
Private Sub lst_Market_Click()
        
    Call Me.o_13_Create_Customer_List_By_Market

    Me.cmb_DynamicSearch.Value = Null
        Me.cmb_DynamicSearch.SetFocus

End Sub
Private Sub lst_Market_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Call Me.o_44_Clear_Customer_Filter
    
    Call Me.o_33_Filter_Customers_by_Market
    
    Call Me.o_24_Add_LOB_Details
    
    Me.cmb_DynamicSearch.Value = Null
        Me.cmb_DynamicSearch.SetFocus

End Sub
Private Sub cmb_DynamicSearch_Change()

    Call Me.o_11_Create_Customer_List_Dynamic

End Sub
Private Sub lst_Customers_Enter()

    If lst_Customers.ListCount = 2 Then
        lst_Customers.Selected(0) = True
            
            Call Me.o_21_Add_Customer_To_Selected_Customers_Array '9/16 Added to only run when one customer selected
            Call Me.o_22_Add_Customer_Details_to_UserForm
            Call Me.o_23_Filter_Single_Customer
    
    End If
    
    If lst_Customers.Value = "" Then Exit Sub

    Me.cmb_DynamicSearch.SetFocus

End Sub

Private Sub lst_Customers_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If lst_Customers.ListCount = 2 Then
        lst_Customers.Selected(0) = True
    End If
    
    txt_SingleCustomer = Null
    
    If lst_Customers.Value = "" Then Exit Sub
    
    Call Me.o_21_Add_Customer_To_Selected_Customers_Array
    
    Call Me.o_22_Add_Customer_Details_to_UserForm
       
    Call Me.o_31_Filter_Customers
    
    Me.cmb_DynamicSearch.SetFocus

End Sub
Private Sub lst_Customers_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    'If I hit enter and have selected a customer filter the list to JUST that customer
    
    If lst_Customers.ListCount = 2 Then
        lst_Customers.Selected(0) = True
    End If
    
    If KeyCode = vbKeyReturn And Me.lst_Customers.Value <> "" Then
        Call Me.o_21_Add_Customer_To_Selected_Customers_Array
        Call Me.o_22_Add_Customer_Details_to_UserForm
        Call Me.o_23_Filter_Single_Customer
        Me.cmb_DynamicSearch.SetFocus
    End If

End Sub
Private Sub cmd_Filter_Customers_Click()
  
    Call Me.o_31_Filter_Customers

    Call Me.o_34_Save_SelectedCustomer_Array_Values

End Sub
Private Sub cmd_Clear_Filter_Click()

    Call Me.o_41_Clear_Filters
    
    Call Me.o_42_Clear_Saved_Array
    
    Call Me.o_43_Clear_Values
    
    Call Me.o_24_Add_LOB_Details

End Sub
Private Sub cmd_Filter_High_Risk_Click()

    Me.o_36_Filter_High_Risk

End Sub
Private Sub cmd_Filter_New_High_Risk_Click()

    Me.o_37_Filter_New_High_Risk

End Sub
Private Sub cmd_Filter_New_Relief_Request_Click()

    Me.o_39_Filter_New_Relief_Request

End Sub
Private Sub cmd_Filter_Relief_Request_Status_Click()

    Me.o_38_Filter_Relief_Request_Status
    
    Call Me.o_24_Add_LOB_Details

End Sub
Private Sub cmd_Filter_Active_Mod_Round_Click()

    Call Me.o_310_Filter_Active_Mod_Round
    
    Call Me.o_24_Add_LOB_Details

End Sub
Private Sub cmd_Filter_ModType_Click()

    Call Me.o_311_Filter_Mod_Type_Payment
    
    Call Me.o_24_Add_LOB_Details

End Sub
Private Sub cmd_Filter_MSLP_Inquiry_Click()

    Call Me.o_312_Filter_MSLP_Inquiry

End Sub
Private Sub cmd_Filter_PPP_Click()

    Call Me.o_313_Filter_PPP_Inquiry

End Sub
Private Sub cmd_Cancel_Click()
    
    Unload Me

End Sub
Sub o_11_Create_Customer_List_Dynamic()

' Purpose: To create the list of customers to be used in the Customer ListBox.
' Trigger: Start typing in the DynamicSearch combo box
' Updated: 3/23/2020

' Change Log:
'   3/23/2020: Intial Creation

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    Dim LOB As String
    
    Dim CustomerName As String

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
            
            LOB = .Range("C" & x).Value2
            CustomerName = .Range("E" & x).Value2

                If InStr(1, CustomerName, Me.cmb_DynamicSearch.Value, vbTextCompare) Then
                    
                    If IsNull(Me.lst_LOB) Or Me.lst_LOB = LOB Then
                        ary_Customers(y) = CustomerName
                        y = y + 1
                    End If
        
                End If
            
            x = x + 1
        Loop
    End With

    ReDim Preserve ary_Customers(1 To y)

    Me.lst_Customers.List = ary_Customers

End Sub
Sub o_12_Create_Customer_List_By_LOB()

' Purpose: To create the list of customers to be used in the Customer ListBox, based on the selected LOB.
' Trigger: Select a customer from the LOB ListBox.
' Updated: 3/23/2020

' Change Log:
'   3/23/2020: Intial Creation

' ****************************************************************************

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

End Sub
Sub o_22_Add_Customer_Details_to_UserForm()

' Purpose: To add the details from the selected customer into the form.
' Trigger: Double Click Customer
' Updated: 3/12/2021

' Change Log:
'       4/2/2020: Intial Creation
'       4/30/2020: Updated to reflect the change in Relief and PPP, and to appluy UCase
'       5/15/2020: Updated to use the new named fields instead of the hardcoding
'       11/3/2020: Updated to include the Round 3 mod fields
'       11/3/2020: Updated to pull from the "Single Customer" field
'       11/13/2020: Switched the order for the "single customer" if statement
'       3/12/2021: Removed the US Economic and Supply Chain concern

' ****************************************************************************

' -----------
' Declare your variables
' -----------
    
    'Dim strCustomerName As String
        If txt_SingleCustomer.Value <> "" Then
            strCustomerName = txt_SingleCustomer.Value
        Else
            strCustomerName = lst_Customers.Value
        End If
        
    Dim intCustRow As Long

    Dim ary_CustList
        ary_CustList = WorksheetFunction.Transpose(wsData.Range("B1:B" & intLastRow_wsData))

    Dim i As Integer
        i = 2

    'Dim Rating Values
    Dim strSupplyChainConcern As String
    Dim strOverallConcern As String
    
    'Dim Relief Values
    Dim strReliefStatus1st As String
    
    Dim bol_1st_ReliefReq As Boolean
    Dim bol_2nd_ReliefReq As Boolean
    Dim bol_3rd_ReliefReq As Boolean
    
    Dim bol_1st_ReliefGranted As Boolean
    Dim bol_2nd_ReliefGranted As Boolean
    Dim bol_3rd_ReliefGranted As Boolean
    
    Dim bol_1st_Payment As Boolean
    Dim bol_2nd_Payment As Boolean
    Dim bol_3rd_Payment As Boolean

    'Dim PPP Values
    Dim strPPPStatus As String
    
    Dim arryPPPReq
        arryPPPReq = Array("Funded", "Approved in eTrans", "OK To Close", "In Process", "Y")
    
    Dim strPPPFunded
        strPPPFunded = "FUNDED"

    Dim bol_PPPReq As Boolean
    Dim bol_PPPFunded As Boolean

' -----------
' Identify the customer number changes to the customer data
' -----------
   
    For i = 2 To intLastRow_wsData
        If ary_CustList(i) = strCustomerName Then
            intCustRow = i
            Exit For
        End If
    Next i
    
' -----------
' Set the values for the ratings
' -----------
           
    With wsData
    
        strSupplyChainConcern = StrConv(.Cells(i, col_SupplyChain), vbProperCase)
        strOverallConcern = StrConv(.Cells(i, col_OverallRisk), vbProperCase)
        strPPPStatus = UCase(.Cells(i, col_PPPRequest).Value)
        
        strReliefStatus1st = UCase(.Cells(i, col_1st_Relief).Value)
        strReliefStatus2nd = UCase(.Cells(i, col_2nd_Relief).Value)
        strReliefStatus3rd = UCase(.Cells(i, col_3rd_Relief).Value)
      
    End With

    'PPP Requested
'    If IsNumeric(Application.Match(strPPPStatus, arryPPPReq, 0)) Then
'        bol_PPPReq = True
'    Else
'        bol_PPPReq = False
'    End If
'
'    'PPP Funded
'    If strPPPStatus = strPPPFunded Then
'        bol_PPPFunded = True
'    Else
'        bol_PPPFunded = False
'    End If

    '1st Relief Requested
    If strReliefStatus1st = "Y - REQUESTED" Or strReliefStatus1st = "Y - GRANTED" Then
        bol_1st_ReliefReq = True
    Else
        bol_1st_ReliefReq = False
    End If
    
    '2nd Relief Requested
    If strReliefStatus2nd = "Y - REQUESTED" Or strReliefStatus2nd = "Y - GRANTED" Then
        bol_2nd_ReliefReq = True
    Else
        bol_2nd_ReliefReq = False
    End If
                    
    '3rd Relief Requested
    If strReliefStatus3rd = "Y - REQUESTED" Or strReliefStatus3rd = "Y - GRANTED" Then
        bol_3rd_ReliefReq = True
    Else
        bol_3rd_ReliefReq = False
    End If
                    
    '1st Relief Granted
    If strReliefStatus1st = "Y - GRANTED" Then
        bol_1st_ReliefGranted = True
    Else
        bol_1st_ReliefGranted = False
    End If

    '2nd Relief Granted
    If strReliefStatus2nd = "Y - GRANTED" Then
        bol_2nd_ReliefGranted = True
    Else
        bol_2nd_ReliefGranted = False
    End If

    '3rd Relief Granted
    If strReliefStatus3rd = "Y - GRANTED" Then
        bol_3rd_ReliefGranted = True
    Else
        bol_3rd_ReliefGranted = False
    End If

' -----------
' Load in the Customer Details from the CV Tracker
' -----------

    Me.frm_Cust_Details.Caption = strCustomerName

    With wsData
        
        Me.txt_LineOfBusiness.Caption = .Cells(i, col_LOB)
        
        Me.txt_Market.Caption = .Cells(i, col_Market)
        
        Me.txt_Port_Mgr.Caption = .Cells(i, col_RelMgr)
        
        Me.txt_Industry.Caption = .Cells(i, col_Industry)
                
        Me.txt_Prop_Type.Caption = .Cells(i, col_PropType)
        
        Me.txt_End_Market.Caption = .Cells(i, col_EndMarket)
        
        Me.txt_Direct_Outstanding.Caption = Format(.Cells(i, col_Outstanding), "$#,##0")
        
        Me.txt_Exposure.Caption = Format(.Cells(i, col_Exposure), "$#,##0")
        
        If .Cells(i, col_ActiveMod_Status) <> "" Then
            Me.lbl_ActiveMod.Value = .Cells(i, col_ActiveMod) & " / " & .Cells(i, col_ActiveMod_Status)
        Else
            Me.lbl_ActiveMod.Value = ""
        End If
        
    End With

' -----------
' Load in the Covid Relief Fields
' -----------
    
    'Relief Details
    
    Me.chk_1st_Relief_Requeseted.Value = bol_1st_ReliefReq
    Me.chk_2nd_Relief_Requeseted.Value = bol_2nd_ReliefReq
    Me.chk_3rd_Relief_Requeseted.Value = bol_3rd_ReliefReq
    
    Me.chk_1st_Relief_Granted.Value = bol_1st_ReliefGranted
    Me.chk_2nd_Relief_Granted.Value = bol_2nd_ReliefGranted
    Me.chk_3rd_Relief_Granted.Value = bol_3rd_ReliefGranted
    
    'Payment Mods
    
    If wsData.Cells(i, col_1st_PayMod) = "Yes" Then
        Me.chk_1st_Payment = True
    Else
        Me.chk_1st_Payment = False
    End If
    
    If wsData.Cells(i, col_2nd_PayMod) = "Yes" Then
        Me.chk_2nd_Payment = True
    Else
        Me.chk_2nd_Payment = False
    End If
    
    If wsData.Cells(i, col_3rd_PayMod) = "Yes" Then
        Me.chk_3rd_Payment = True
    Else
        Me.chk_3rd_Payment = False
    End If
    
    'Mod Maturity Dates
    
    If wsData.Cells(i, col_1st_ModMaturity) <> "2022+" Then '1st Mod
        Me.txt_1st_MaturityDate.Value = Format(wsData.Cells(i, col_1st_ModMaturity), "MMM YYYY")
    Else
        Me.txt_1st_MaturityDate.Value = wsData.Cells(i, col_1st_ModMaturity)
    End If
    
    If wsData.Cells(i, col_2nd_ModMaturity) <> "2022+" Then '2nd Mod
        Me.txt_2nd_MaturityDate.Value = Format(wsData.Cells(i, col_2nd_ModMaturity), "MMM YYYY")
    Else
        Me.txt_2nd_MaturityDate.Value = wsData.Cells(i, col_2nd_ModMaturity)
    End If

    If wsData.Cells(i, col_3rd_ModMaturity) <> "2022+" Then '3rd Mod
        Me.txt_3rd_MaturityDate.Value = Format(wsData.Cells(i, col_3rd_ModMaturity), "MMM YYYY")
    Else
        Me.txt_3rd_MaturityDate.Value = wsData.Cells(i, col_3rd_ModMaturity)
    End If

' -----------
' Load in the PPP and and MSLP Fields
' -----------
    
'    'PPP Requested
'    Me.chk_PPP_Requeseted.Value = bol_PPPReq
'
'    'PPP Funded
'    Me.chk_PPP_Funded.Value = bol_PPPFunded

' -----------
' Load in the Comments Fields
' -----------

    Me.frm_Cust_Ratings_and_Commentary.Caption = strCustomerName
    
    With wsData
        Me.txt_Comments = .Cells(i, col_Comments)
        Me.txt_Relief_Comments = .Cells(i, col_ReliefComments)
    End With

' -----------
' Reset to the default colors
' -----------

    Me.txt_Rating_Overall_Concern.BackColor = RGB(240, 240, 240)
    Me.txt_Rating_Supply_Chain_Concern.BackColor = RGB(240, 240, 240)
    Me.chk_1st_Relief_Requeseted.BackColor = RGB(240, 240, 240)

' -----------
' Load in the Risk Ratings from the CV Tracker
' -----------
    With wsData
                        
        Me.txt_BRG.Caption = .Cells(i, col_BRG)
        
        Me.txt_FRG.Caption = .Cells(i, col_FRG)
        
        Me.txt_CCRP.Caption = .Cells(i, col_CCRP)
            If Me.txt_CCRP.Caption = "6.5" Then Me.txt_CCRP.Caption = "6W"
        
        Me.txt_LFT.Caption = .Cells(i, col_LFT)
        
        Me.txt_Rating_Supply_Chain_Concern = strSupplyChainConcern
        
            With Me.txt_Rating_Supply_Chain_Concern
                 Select Case strSupplyChainConcern
                    Case "Low"
                        .ForeColor = RGB(155, 187, 89)
                    Case "Medium"
                        .ForeColor = RGB(247, 150, 70)
                    Case "High"
                        .ForeColor = RGB(192, 0, 0)
                End Select
            End With
        
        Me.txt_Rating_Overall_Concern = strOverallConcern
    
            With Me.txt_Rating_Overall_Concern
                 Select Case strOverallConcern
                    Case "Low"
                        .ForeColor = RGB(155, 187, 89)
                    Case "Medium"
                        .ForeColor = RGB(247, 150, 70)
                    Case "High"
                        .ForeColor = RGB(192, 0, 0)
                End Select
            End With
        
       ' Change the background color for New High Risk
        If .Cells(i, col_NewHighRisk) <> "" Then
            If CDate(.Cells(i, col_NewHighRisk)) >= Date - 6 Then
                Me.txt_Rating_Overall_Concern.BackColor = RGB(253, 234, 219)
            End If
        End If
    
       ' Change the background color for New Relief Request
        If .Cells(i, col_NewReliefRequest1st) <> "" Then
            If CDate(.Cells(i, col_NewReliefRequest1st)) >= Date - 7 Then
                Me.chk_1st_Relief_Requeseted.BackColor = RGB(253, 234, 219)
            End If
        End If
        
    End With

End Sub
Sub o_23_Filter_Single_Customer()

' Purpose: To filter the list of customers in the Data ws based on only the currently swelected customer in the Customers List.
' Trigger: Called: uf_CV_Tracker_Regular
' Updated: 4/2/2020

' Change Log:
'          4/22/2020: Intial Creation

' ****************************************************************************

    With wsData

      .AutoFilterMode = False
    
      .Cells.AutoFilter
          .Range("A:" & strLastCol_wsData).AutoFilter Field:=col_Customer, Criteria1:=Me.lst_Customers.Value, Operator:=xlFilterValues
    
    End With

End Sub
Sub o_24_Add_LOB_Details()

' Purpose: To add the LOB / Market deatils into the userform.
' Trigger: TBD
' Updated: 1/4/2021

' Change Log:
'       9/17/2020: Intial Creation
'       1/4/2021: Added the section for Payment Mods
'       1/14/2021: Updated to reflect the new wording for the Full vs Partial payment mods

' ****************************************************************************

' -----------
' Declare Variables
' -----------

    Dim dict_PayMods As Scripting.Dictionary
        Set dict_PayMods = New Scripting.Dictionary
        
    Dim dict_PayMods_NoLFT As Scripting.Dictionary
        Set dict_PayMods_NoLFT = New Scripting.Dictionary

    Dim i As Integer
    
    Dim dblPayMods As Double
    
    Dim dblPayMods_NoLFT As Double
    
    Dim val As Variant

' -----------
' Set the values for LOB Summary Metrics
' -----------

With wsData

On Error Resume Next

    ' Outstanding

    Dim int_LOB_Outstanding As Double
        int_LOB_Outstanding = Application.WorksheetFunction.Sum(.Range(.Cells(2, col_Outstanding), .Cells(intLastRow_wsData, col_Outstanding)).SpecialCells(xlCellTypeVisible))
        
        int_LOB_Outstanding = int_LOB_Outstanding / 10 ^ 6
    
    Me.txt_LOB_Outstanding = Format(int_LOB_Outstanding, "$#,##0 MM")

    ' Exposure

    Dim int_LOB_Exposure As Double
        int_LOB_Exposure = Application.WorksheetFunction.Sum(.Range(.Cells(2, col_Exposure), .Cells(intLastRow_wsData, col_Exposure)).SpecialCells(xlCellTypeVisible))
        
        int_LOB_Exposure = int_LOB_Exposure / 10 ^ 6
    
    Me.txt_LOB_Exposure = Format(int_LOB_Exposure, "$#,##0 MM")

    ' Create Payment Mod Dictionaries

    For i = 2 To intLastRow_wsData
        If Rows(i).EntireRow.Hidden = False Then
            If Left(wsData.Cells(i, col_ActiveMod).Value2, 3) = "Yes" Then
                If Left(wsData.Cells(i, col_PaymentMod), 3) = "Yes" Then
                    
                    dict_PayMods.Add Key:=wsData.Cells(i, col_Customer), Item:=wsData.Cells(i, col_Outstanding)
                    
                    If wsData.Cells(i, col_LFT) <> 8 Then
                        dict_PayMods_NoLFT.Add Key:=wsData.Cells(i, col_Customer), Item:=wsData.Cells(i, col_Outstanding)
                    End If
                    
                End If
            End If
        End If
    Next i

    ' Payment Mods Total

    For Each val In dict_PayMods
        dblPayMods = dblPayMods + dict_PayMods(val)
    Next val
    
    dblPayMods = dblPayMods / 10 ^ 6
    
    Me.txt_PayMods = Format(dblPayMods, "$#,##0 MM")

    ' Payment Mods (No LFT 8) Total

    For Each val In dict_PayMods_NoLFT
        dblPayMods_NoLFT = dblPayMods_NoLFT + dict_PayMods_NoLFT(val)
    Next val
    
    dblPayMods_NoLFT = dblPayMods_NoLFT / 10 ^ 6
    
    Me.txt_PayMods_NoLFT = Format(dblPayMods_NoLFT, "$#,##0 MM")

On Error GoTo 0

End With

End Sub
Sub o_31_Filter_Customers()

' Purpose: To filter the list of customers in the Data ws based on the Customer Selected Array.
' Trigger: Called: uf_CV_Tracker_Regular.cmd_Filter_Customers
' Updated: 3/23/2020

' Change Log:
'   3/23/2020: Intial Creation

' ****************************************************************************

Call PrivateMacros.DisableForEfficiency

' -----------
' Filter the Summary worksheet based on the customers selected previously
' -----------
    With wsData

      .AutoFilterMode = False
          .Range("A:" & strLastCol_wsData).AutoFilter Field:=col_Customer, Criteria1:=ary_SelectedCustomers, Operator:=xlFilterValues
    
    End With

    Application.GoTo Reference:=wsData.Range("A1"), Scroll:=True

Call PrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_32_Filter_Customers_by_LOB()

' Purpose: To filter the list of customers in the Data ws based on the selected LOB.
' Trigger: Called: uf_CV_Tracker_Regular.cmd_Filter_Customers_by_LOB
' Updated: 3/23/2020

' Change Log:
'       3/23/2020: Intial Creation
'       9/16/2020: Updated so that only the LOB column will lose the filter

' ****************************************************************************

Call PrivateMacros.DisableForEfficiency
        
Dim aryTemp As Variant
    If Me.lst_LOB.Value = "Restructure & Recovery" Then aryTemp = Application.Transpose(Me.lst_Customers.List)
        
' -----------
' Filter the Summary worksheet based on the customers selected previously
' -----------
    With wsData

          .Range("A:" & strLastCol_wsData).AutoFilter Field:=col_LOB, Criteria1:=Me.lst_LOB.Value, Operator:=xlFilterValues
          .Range("A:" & strLastCol_wsData).AutoFilter Field:=col_Market
    
    End With

    If Me.lst_LOB.Value = "Restructure & Recovery" Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_LOB
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_Customer, Criteria1:=aryTemp, Operator:=xlFilterValues
    End If

    Application.GoTo Reference:=wsData.Range("A1"), Scroll:=True
    
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

' -----------
' Filter the Summary worksheet based on the customers selected previously
' -----------
    
    wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_Market, Criteria1:=Me.lst_Market.Value, Operator:=xlFilterValues

    Application.GoTo Reference:=wsData.Range("A1"), Scroll:=True
    
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
 
 PrivateMacros.DisableForEfficiency

' -----------
' Filter the Summary worksheet based on the customers selected previously
' -----------
    With wsData

      .AutoFilterMode = False
          .Range("A:" & strLastCol_wsData).AutoFilter Field:=col_ChangeFlag, Criteria1:="CHANGE", Operator:=xlFilterValues
    
    End With

    Application.GoTo Reference:=wsData.Range("A1"), Scroll:=True

PrivateMacros.DisableForEfficiencyOff

End Sub
Sub o_36_Filter_High_Risk()

' Purpose: To filter the Data down to just the High Risk customers for a portfolio, and allow the filter to be toggled.
' Trigger: Called: cmd_Filter_High_Risk
' Updated: 4/7/2020

' Change Log:
'   4/7/2020: Intial Creation

' ****************************************************************************

' -----------
' Filter the data
' -----------
          
    If bol_HighRiskFilter = False Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_OverallRisk, Criteria1:="HIGH", Operator:=xlFilterValues
        Me.cmd_Filter_High_Risk.BackColor = RGB(240, 248, 224)
    ElseIf bol_HighRiskFilter = True Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_OverallRisk
        Me.cmd_Filter_High_Risk.BackColor = RGB(227, 227, 227)
    End If
    
    bol_HighRiskFilter = Not bol_HighRiskFilter 'Switch the boolean
     
End Sub
Sub o_37_Filter_New_High_Risk()

' Purpose: To filter the Data down to just the NEW High Risk customers for a portfolio, and allow the filter to be toggled.
' Trigger: Called: cmd_Filter_New_High_Risk
' Updated: 4/7/2020

' Change Log:
'   4/7/2020: Intial Creation

' ****************************************************************************

' -----------
' Filter the data
' -----------
          
    If bol_NewHighRisk = False Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_NewHighRisk, Criteria1:=">=" & Date - 7, Operator:=xlFilterValues
        Me.cmd_Filter_New_High_Risk.BackColor = RGB(240, 248, 224)
    ElseIf bol_NewHighRisk = True Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_NewHighRisk
        Me.cmd_Filter_New_High_Risk.BackColor = RGB(227, 227, 227)
    End If
    
    bol_NewHighRisk = Not bol_NewHighRisk 'Switch the boolean


End Sub
Sub o_38_Filter_Relief_Request_Status()

' Purpose: To filter the data down to just the customers that have Requested or been Granted relief.
' Trigger: Called: cmd_Filter_New_Relief_Request
' Updated: 9/14/2020

' Change Log:
'          4/7/2020: Intial Creation
'          4/30/2020: Updated to reflect the change in Relief Requst
'          9/14/2020: Combined Relief Requested and Granted
'          9/17/2020: Updated to use the 'Active Mod Status' field instead of 1st round

' ****************************************************************************

' -----------
' Filter the data
' -----------
          
     If bol_ReliefStatus = "" Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_ActiveMod_Status, Criteria1:="Y*", Operator:=xlFilterValues
        Me.cmd_Filter_Relief_Request_Status.BackColor = RGB(240, 248, 224)
        Me.cmd_Filter_Relief_Request_Status.Caption = "All Relief"
        bol_ReliefStatus = "All Relief"
          
    ElseIf bol_ReliefStatus = "All Relief" Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_ActiveMod_Status, Criteria1:="Y - Requested", Operator:=xlFilterValues
        Me.cmd_Filter_Relief_Request_Status.BackColor = RGB(240, 248, 224)
        Me.cmd_Filter_Relief_Request_Status.Caption = "Requested"
        bol_ReliefStatus = "Requested"
          
    ElseIf bol_ReliefStatus = "Requested" Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_ActiveMod_Status, Criteria1:="Y - Granted", Operator:=xlFilterValues
        Me.cmd_Filter_Relief_Request_Status.BackColor = RGB(240, 248, 224)
        Me.cmd_Filter_Relief_Request_Status.Caption = "Granted"
        bol_ReliefStatus = "Granted"
          
    Else
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_ActiveMod_Status
        Me.cmd_Filter_Relief_Request_Status.BackColor = RGB(227, 227, 227)
        Me.cmd_Filter_Relief_Request_Status.Caption = "Relief Status"
        bol_ReliefStatus = ""
        
    End If



End Sub
Sub o_39_Filter_New_Relief_Request()

' Purpose: To filter the Data down to just the NEW Relief Requested customers for a portfolio, and allow the filter to be toggled.
' Trigger: Called: cmd_Filter_New_Relief_Request
' Updated: 9/14/2020

' Change Log:
'          4/7/2020: Intial Creation
'          9/14/2020: Combined Relief requested and Granted

' ****************************************************************************

' -----------
' Filter the data
' -----------
          
    If bol_NewReliefRequest = False Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_NewReliefRequest1st, Criteria1:=">=" & Date - 7, Operator:=xlFilterValues
        Me.cmd_Filter_New_Relief_Request.BackColor = RGB(240, 248, 224)
    ElseIf bol_NewReliefRequest = True Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_NewReliefRequest1st
        Me.cmd_Filter_New_Relief_Request.BackColor = RGB(227, 227, 227)
    End If
    
    bol_NewReliefRequest = Not bol_NewReliefRequest 'Switch the boolean

End Sub
Sub o_310_Filter_Active_Mod_Round()

' Purpose: To filter the data to each round of Active Mods.
' Trigger: Called: cmd_Filter_ActiveMod
' Updated: 9/13/2020

' Change Log:
'          9/2/2020: Intial Creation
'          9/13/2020: Updated to cycle through 1st, 2nd, 3rd Round Mods
'          9/13/2020: Combined the 1st relief Only and the Active Mods macros

' ****************************************************************************


    Dim arry_ActiveMod
        arry_ActiveMod = Array("Yes - 1st Mod Active", "Yes - 2nd Mod Active", "Yes - 3rd Mod Active")

' -----------
' Filter the data
' -----------
          
    If bol_ActiveMod = "" Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_ActiveMod, Criteria1:=arry_ActiveMod, Operator:=xlFilterValues
        Me.cmd_Filter_Active_Mod_Round.BackColor = RGB(240, 248, 224)
        Me.cmd_Filter_Active_Mod_Round.Caption = "All Active"
        bol_ActiveMod = "All Active Mods"
    
    ElseIf bol_ActiveMod = "All Active Mods" Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_ActiveMod, Criteria1:="Yes - 1st Mod Active", Operator:=xlFilterValues
        Me.cmd_Filter_Active_Mod_Round.BackColor = RGB(240, 248, 224)
        Me.cmd_Filter_Active_Mod_Round.Caption = "1st Round"
        bol_ActiveMod = "1st Mod"
    
    ElseIf bol_ActiveMod = "1st Mod" Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_ActiveMod, Criteria1:="Yes - 2nd Mod Active", Operator:=xlFilterValues
        Me.cmd_Filter_Active_Mod_Round.BackColor = RGB(240, 248, 224)
        Me.cmd_Filter_Active_Mod_Round.Caption = "2nd Round"
        bol_ActiveMod = "2nd Mod"
    
    ElseIf bol_ActiveMod = "2nd Mod" Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_ActiveMod, Criteria1:="Yes - 3rd Mod Active", Operator:=xlFilterValues
        Me.cmd_Filter_Active_Mod_Round.BackColor = RGB(240, 248, 224)
        Me.cmd_Filter_Active_Mod_Round.Caption = "3rd Round"
        bol_ActiveMod = "3rd Mod"
    
    ElseIf bol_ActiveMod = "3rd Mod" Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_ActiveMod
        Me.cmd_Filter_Active_Mod_Round.BackColor = RGB(227, 227, 227)
        Me.cmd_Filter_Active_Mod_Round.Caption = "Mod Round"
        bol_ActiveMod = ""
    End If

End Sub
Sub o_311_Filter_Mod_Type_Payment()

' Purpose: To filter the data down to just the Payment Mods in the portfolio.
' Trigger: Called: cmd_Filter_ModType
' Updated: 9/2/2020

' Change Log:
'          9/2/2020: Intial Creation
'          9/14/2020: Updated to cycle through payment vs non-payment

' ****************************************************************************

' -----------
' Filter the data
' -----------
          
    If bol_PaymentMod = "" Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_PaymentMod, Criteria1:="Yes*", Operator:=xlFilterValues
        Me.cmd_Filter_ModType.BackColor = RGB(240, 248, 224)
        Me.cmd_Filter_ModType.Caption = "Payment"
        bol_PaymentMod = "Payment"
    
    ElseIf bol_PaymentMod = "Payment" Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_PaymentMod, Criteria1:="No", Operator:=xlFilterValues
        Me.cmd_Filter_ModType.BackColor = RGB(240, 248, 224)
        Me.cmd_Filter_ModType.Caption = "Non-Pay"
        bol_PaymentMod = "Non-Payment"
    
    ElseIf bol_PaymentMod = "Non-Payment" Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_PaymentMod
        Me.cmd_Filter_ModType.BackColor = RGB(227, 227, 227)
        Me.cmd_Filter_ModType.Caption = "Pay vs Non"
        bol_PaymentMod = ""
    End If

End Sub
Sub o_312_Filter_MSLP_Inquiry()

' Purpose: To filter the data down to just the customers that have inquired about the MSLP program.
' Trigger: Called: cmd_Filter_MSLP_Inquiry
' Updated: 4/30/2020

' Change Log:
'          4/30/2020: Intial Creation

' ****************************************************************************

' -----------
' Filter the data
' -----------
          
    If bol_MSLPInquiry = False Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_MSLPInquiry, Criteria1:="Y*", Operator:=xlFilterValues
        Me.cmd_Filter_MSLP_Inquiry.BackColor = RGB(240, 248, 224)
    ElseIf bol_MSLPInquiry = True Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_MSLPInquiry
        Me.cmd_Filter_MSLP_Inquiry.BackColor = RGB(227, 227, 227)
    End If
    
    bol_MSLPInquiry = Not bol_MSLPInquiry 'Switch the boolean

End Sub
Sub o_313_Filter_PPP_Inquiry()

' Purpose: To filter the data down to just the customers that requested PPP funds.
' Trigger: Called: cmd_Filter_PPP
' Updated: 4/30/2020

' Change Log:
'          4/30/2020: Intial Creation

' ****************************************************************************

    Dim arryPPPReq
        arryPPPReq = Array("Funded", "Approved in eTrans", "OK To Close", "In Process", "Y")

' -----------
' Filter the data
' -----------
          
    If bol_PPP = False Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_PPPRequest, Criteria1:=arryPPPReq, Operator:=xlFilterValues
        Me.cmd_Filter_PPP.BackColor = RGB(240, 248, 224)
    ElseIf bol_PPP = True Then
        wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_PPPRequest
        Me.cmd_Filter_PPP.BackColor = RGB(227, 227, 227)
    End If
    
    bol_PPP = Not bol_PPP 'Switch the boolean

End Sub
Sub o_41_Clear_Filters()

' Purpose: To reset all of the current filtering.
' Trigger: Called: uf_CV_Tracker_Regular.cmd_Clear_Filter
' Updated: 3/23/2020

' Change Log:
'   3/23/2020: Intial Creation

' ****************************************************************************

' -----------
' If the AutoFilter is on turn it off and then reapply
' -----------

    If wsData.AutoFilterMode = True Then
        wsData.AutoFilterMode = False
        wsData.Range("A:" & strLastCol_wsData).AutoFilter
    End If
    
End Sub
Sub o_42_Clear_Saved_Array()

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

End Sub
Sub o_43_Clear_Values()

' Purpose: To remove all of the values from the userform.
' Trigger: Called: uf_Impact_Meeting.cmd_Clear_Filter
' Updated: 5/19/2021

' Change Log:
'       4/9/2020: Intial Creation
'       6/9/2020: Added the new fields in to get wiped
'       10/16/2020: Added in additional code to reset the values for the filter buttons
'       3/12/2021: Removed the US Economic and Supply Chain concern
'       5/19/2021: Added back in txt_Rating_Supply_Chain_Concern

' ****************************************************************************

' -----------
' Clear the LOB / Customer values
' -----------

    Me.lst_LOB.Value = ""
    Me.lst_Market.Clear
    Me.lst_Customers.Clear
    Me.cmb_DynamicSearch.Value = Null
    Me.txt_LOB_Outstanding.Caption = ""
    Me.txt_LOB_Exposure.Caption = ""
    Me.txt_PayMods.Caption = ""
    Me.txt_PayMods_NoLFT.Caption = ""

' -----------
' Reset the Customer Details
' -----------

    Me.frm_Cust_Details.Caption = "Customer Details"

    Me.txt_LineOfBusiness.Caption = ""
    Me.txt_Market.Caption = ""
    Me.txt_Port_Mgr.Caption = ""
    Me.txt_Industry.Caption = ""
    Me.txt_Prop_Type.Caption = ""
    Me.txt_End_Market = ""
    
    Me.txt_Direct_Outstanding.Caption = ""
    Me.txt_Exposure.Caption = ""
        
' -----------
' Reset the Comments Fields
' -----------

    Me.frm_Cust_Ratings_and_Commentary.Caption = "Customer Risk Ratings & Relief Commentary"
    
    Me.chk_1st_Relief_Requeseted.Value = False
    Me.chk_2nd_Relief_Requeseted.Value = False
    Me.chk_3rd_Relief_Requeseted.Value = False
    
    Me.chk_1st_Relief_Granted.Value = False
    Me.chk_2nd_Relief_Granted.Value = False
    Me.chk_3rd_Relief_Granted.Value = False
    
    Me.txt_1st_MaturityDate = ""
    Me.txt_2nd_MaturityDate = ""
    Me.txt_3rd_MaturityDate = ""
    
    Me.chk_1st_Payment.Value = False
    Me.chk_2nd_Payment.Value = False
    Me.chk_3rd_Payment.Value = False
    
'    Me.chk_PPP_Requeseted.Value = False
'    Me.chk_PPP_Funded.Value = False
                
    Me.txt_Comments = ""
    Me.txt_Relief_Comments = ""
    Me.lbl_ActiveMod.Value = ""
    
' -----------
' Reset the Risk Ratings
' -----------
                        
    Me.txt_BRG.Caption = ""
    Me.txt_FRG.Caption = ""
    Me.txt_CCRP.Caption = ""
    Me.txt_LFT.Caption = ""
                
    Me.txt_Rating_Overall_Concern = ""
    Me.txt_Rating_Supply_Chain_Concern = ""
    
    Me.txt_Rating_Overall_Concern.BackColor = RGB(240, 240, 240)
    Me.txt_Rating_Supply_Chain_Concern.BackColor = RGB(240, 240, 240)
    Me.chk_1st_Relief_Requeseted.BackColor = RGB(240, 240, 240)
     
' -----------
' Reset the Filter cmd control values
' -----------
     
    Me.cmd_Filter_Relief_Request_Status.Caption = "Relief Status"
        bol_ReliefStatus = ""
     
    Me.cmd_Filter_Active_Mod_Round.Caption = "Mod Round"
        bol_ActiveMod = ""
        
    Me.cmd_Filter_ModType.Caption = "Pay vs Non"
        bol_PaymentMod = ""
     
' -----------
' Reset the Filter cmd control colors
' -----------
     
    Me.cmd_Filter_High_Risk.BackColor = RGB(227, 227, 227)
    Me.cmd_Filter_New_High_Risk.BackColor = RGB(227, 227, 227)
    
    Me.cmd_Filter_Relief_Request_Status.BackColor = RGB(227, 227, 227)
    Me.cmd_Filter_New_Relief_Request.BackColor = RGB(227, 227, 227)
    Me.cmd_Filter_Active_Mod_Round.BackColor = RGB(227, 227, 227)
    
    Me.cmd_Filter_MSLP_Inquiry.BackColor = RGB(227, 227, 227)
    Me.cmd_Filter_PPP.BackColor = RGB(227, 227, 227)
    
    Me.cmd_Filter_ModType.BackColor = RGB(227, 227, 227)
    
End Sub
Sub o_44_Clear_Customer_Filter()

' Purpose: To remove the filtering for the Customer Name field.
' Trigger: Called: uf_CV_Tracker_Regular - Various
' Updated: 5/17/2021

' Change Log:
'       4/23/2020: Intial Creation
'       9/17/2020: Added additional code from 43_Clear_Values
'       3/12/2021: Removed the US Economic and Supply Chain concern
'       5/17/2021: Added back in Supply Chain Concern

' ****************************************************************************

' -----------
' Clear the filter for customer name
' -----------

    wsData.Range("A:" & strLastCol_wsData).AutoFilter Field:=col_Customer
    
' -----------
' Reset the LOB data
' -----------
    
    Me.txt_LOB_Outstanding.Caption = ""
    Me.txt_LOB_Exposure.Caption = ""
    Me.txt_PayMods.Caption = ""
    Me.txt_PayMods_NoLFT.Caption = ""
    
' -----------
' Reset the Customer Details
' -----------

    Me.frm_Cust_Details.Caption = "Customer Details"

    Me.txt_LineOfBusiness.Caption = ""
    Me.txt_Market.Caption = ""
    Me.txt_Port_Mgr.Caption = ""
    Me.txt_Industry.Caption = ""
    Me.txt_Prop_Type.Caption = ""
    Me.txt_End_Market = ""
    
    Me.txt_Direct_Outstanding.Caption = ""
    Me.txt_Exposure.Caption = ""
        
' -----------
' Reset the Comments Fields
' -----------

    Me.frm_Cust_Ratings_and_Commentary.Caption = "Customer Risk Ratings & Relief Commentary"
    
    Me.chk_1st_Relief_Requeseted.Value = False
    Me.chk_2nd_Relief_Requeseted.Value = False
    Me.chk_3rd_Relief_Requeseted.Value = False
    
    Me.chk_1st_Relief_Granted.Value = False
    Me.chk_2nd_Relief_Granted.Value = False
    Me.chk_3rd_Relief_Granted.Value = False
    
    Me.txt_1st_MaturityDate = ""
    Me.txt_2nd_MaturityDate = ""
    Me.txt_3rd_MaturityDate = ""
    
    Me.chk_1st_Payment.Value = False
    Me.chk_2nd_Payment.Value = False
    Me.chk_3rd_Payment.Value = False

'    Me.chk_PPP_Requeseted.Value = False
'    Me.chk_PPP_Funded.Value = False
                
    Me.txt_Comments = ""
    Me.txt_Relief_Comments = ""
    Me.lbl_ActiveMod.Value = ""
    
' -----------
' Reset the Risk Ratings
' -----------
                        
    Me.txt_BRG.Caption = ""
    Me.txt_FRG.Caption = ""
    Me.txt_CCRP.Caption = ""
    Me.txt_LFT.Caption = ""
                
    Me.txt_Rating_Overall_Concern = ""
    Me.txt_Rating_Supply_Chain_Concern = ""
    
    Me.txt_Rating_Overall_Concern.BackColor = RGB(240, 240, 240)
    Me.chk_1st_Relief_Requeseted.BackColor = RGB(240, 240, 240)
    
End Sub
