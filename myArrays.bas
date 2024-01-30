Attribute VB_Name = "myArrays"
Dim LOB_Array

'Dim Area_Array_Personal(1 To 7) As String

Option Explicit
Public Function Get_LOB_Array() As Variant
    
    Dim ws_Lists As Worksheet
        Set ws_Lists = ThisWorkbook.Sheets("Lists")
    
    Dim int_LastRow As Integer
        int_LastRow = ws_Lists.Cells(Rows.Count, "A").End(xlUp).Row
        
        ' Pull all the values from A2 to the end of the list
        LOB_Array = ws_Lists.Range(ws_Lists.Cells(2, 1), ws_Lists.Cells(int_LastRow, 1))

    Get_LOB_Array = LOB_Array
    
End Function
