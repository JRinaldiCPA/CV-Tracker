Attribute VB_Name = "Macros"
Option Explicit

Public strTemplateVersion As String
Public strUserFormSelection As String
Sub Open_CV_Tracker_UF()

' Purpose: To open the applicable UserForm in the CMML CV Tracker.
' Trigger: Called by the Quick Access Toolbar
' Updated: 2/4/2021

' Change Log:
'       1/29/2021: Added the split between the "Regular" and "Admin" forms
'       2/4/2021: Added a new split for the CV Impact UserForm

' ****************************************************************************

' -----------
' Declare your variables
' -----------

    'Dim Booleans
    
        Dim bolPrivilegedUser As Boolean
            bolPrivilegedUser = fx_Privileged_User

    strUserFormSelection = ""

' -----------
' Open the applicable form
' -----------

    ThisWorkbook.Sheets("Data").Activate

    If bolPrivilegedUser = True Then
        SelectForm.Show
    Else
        GoTo REGULAR
    End If

    If strUserFormSelection = "Privileged" Then
        uf_CV_Tracker_Admin.Show vbModeless
    ElseIf strUserFormSelection = "CV Impact" Then
        uf_Impact_Meeting.Show vbModeless
    ElseIf strUserFormSelection = "Regular" Then

REGULAR:
        uf_CV_Tracker_Regular.Show vbModeless
        
        ' Do this to force the Dynamic Search to be active
        uf_CV_Tracker_Regular.cmb_DynamicSearch.Enabled = False
            uf_CV_Tracker_Regular.cmb_DynamicSearch.Enabled = True
        
        uf_CV_Tracker_Regular.cmb_DynamicSearch.SetFocus
    End If

End Sub
Sub Open_CV_Impact_Meeting_UF()
Attribute Open_CV_Impact_Meeting_UF.VB_ProcData.VB_Invoke_Func = "I\n14"

' Purpose: To open the CMML CV Tracker
' Trigger: Called
' Updated: 2/4/2021

' Change Log:
'       2/4/2021: Added a new split for the CV Impact UserForm
'       2/4/2021: Reset back to the basic code

' ****************************************************************************

    ThisWorkbook.Sheets("Data").Activate
    
    uf_Impact_Meeting.Show vbModeless

End Sub
