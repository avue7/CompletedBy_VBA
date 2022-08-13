Attribute VB_Name = "LinkAllCheckBoxesToCell"
Sub LinkAllCheckBoxes()
    Dim wrksht As Worksheet
    Dim chk As Object
    Dim rng As range
    
    Set wrksht = ThisWorkbook.Worksheets("Sheet1")
    
    For Each chk In wrksht.OLEObjects
        counter = counter + 1
        If TypeName(wrksht.OLEObjects(chk.Name).Object) = "CheckBox" Then
            With chk
                .LinkedCell = _
                .TopLeftCell.Offset(0, "1").Address
            End With
            chk.TopLeftCell.Offset(0, "1").EntireColumn.Hidden = True
            chk.Object.Value = False
        End If
    Next chk
End Sub

                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                