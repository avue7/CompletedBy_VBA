Attribute VB_Name = "RemoveAllCheckboxes"
Option Explicit
 
Sub DeleteAllControls()
    Dim obj As OLEObject
     'Selects all ActiveX controls on the active sheet and deletes them
    For Each obj In ActiveSheet.OLEObjects
        obj.Delete
    Next obj
End Sub
 
 'If you wish to delete the controls selectively
 'This example will allow you to delete only the ActiveX Comboboxes on the sheet
Sub DeleteComboboxesOnly()
    Dim obj As OLEObject
     'Checks each ActiveX control on the active sheet
     'If the control is a ComboBox delete it
     'go to the next control and do same.
    For Each obj In ActiveSheet.OLEObjects
        If obj.progID = "Forms.ComboBox.1" Then obj.Delete
    Next obj
End Sub
6 , 2 1 3 7 8 2 4 3 , 5 9 5 1 7 4 5 9 4 , 1 7 8 4 6 7 3 0 , 1 7 8 4 6 7 5 3 , 1 7 8 8 5 4 0 9 , 2 1 3 1 3 5 0 3 , 1 8 4 0 1 4 1 6 , 6 3 6 6 2 9 0 , 1 8 4 0 1 4 1 3 , 2 1 3 7 8 2 5 5 , 6 3 6 6 2 9 1 , 7 1 1 6 0 5 3 , 2 2 9 2 9 4 2 7 , 1 8 4 0 1 4 1 4 , 3 6 2 9 2 4 3 5 , 2 1 5 6 1 4 8 7 , 1 9 2 0 0 0 8 2 , 2 1 3 1 3 5 0 6 , 2 1 3 1 3 5 3 7 , 2 1 5 8 7 0 8 2 , 1 8 4 0 1 4 1 5 , 9 2 4 2 0 0 9 , 1 9 2 0 0 0 8 8 , 2 1 3 1 3 5 0 7 , 2 1 3 7 8 2 4 6 , 5 0 8 9 0 3 2 8 , 2 1 3 7 8 2 4 7 , 5 8 9 8 8 4 5 , 2 1 3 7 8 2 