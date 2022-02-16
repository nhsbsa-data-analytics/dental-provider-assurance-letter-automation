'Adapted from https://answers.microsoft.com/en-us/msoffice/forum/all/dealing-with-merged-cells-in-vba/1fbafd59-967f-474c-9298-303ec48cefbc

Sub DeleteEmptyRowsSkipIfMerged(TargetDoc As Document)
    Dim oTbl As Table
    Dim oRow As Row
    Dim i As Long
    Dim j As Long
    Dim r As Long

    For Each oTbl In TargetDoc.Tables
        For i = 1 To oTbl.Rows.Count
            On Error Resume Next
            i = oTbl.Rows(i).Index
            If Err.Number <> 5991 Then
                With oTbl
                    For r = .Rows.Count To 1 Step -1
                        With .Rows(r)
                            If Len(.Range.Text) = .Cells.Count * 2 + 2 Then .Delete
                        End With
                    Next
                End With
            End If
        Next i
    Next oTbl
End Sub
