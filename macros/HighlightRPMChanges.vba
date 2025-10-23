Sub HighlightRPMChanges()
    Dim ws As Worksheet
    Dim sheetName As String
    
    ' Ïîëó÷àåì èìÿ àêòèâíîãî ëèñòà èëè çàïðàøèâàåì ó ïîëüçîâàòåëÿ
    On Error Resume Next
    sheetName = Application.InputBox("Ââåäèòå èìÿ ëèñòà äëÿ àíàëèçà:", "Àíàëèç RPM", ActiveSheet.Name, Type:=2)
    If sheetName = "False" Then Exit Sub ' Ïîëüçîâàòåëü íàæàë Cancel
    
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then
        MsgBox "Ëèñò '" & sheetName & "' íå íàéäåí!"
        Exit Sub
    End If
    
    'Find the last row with data
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    'Find the last column with data
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    'Each date occupies 9 columns (B-J, K-S, T-AB, etc.)
    Dim dayCount As Integer
    dayCount = (lastCol - 1) / 9 'Subtract column A, divide by 9 columns per date
    
    'Loop through all days starting from the second day
    For i = 2 To dayCount
        'Determine RPM columns for previous and current days
        Dim prevDayRPMCol As Integer
        prevDayRPMCol = 3 + (i - 2) * 9 'RPM is in the 3rd column of the group
        
        Dim currDayRPMCol As Integer
        currDayRPMCol = 3 + (i - 1) * 9 'RPM is in the 3rd column of the group
        
        'Loop through all data rows (starting from row 3)
        For j = 3 To lastRow
            'Skip rows with empty well names
            If ws.Cells(j, 1).Value <> "" Then
                'Get values from both days
                Dim prevValue As Variant
                Dim currValue As Variant
                prevValue = ws.Cells(j, prevDayRPMCol).Value
                currValue = ws.Cells(j, currDayRPMCol).Value
                
                'Only compare if both values are not empty
                If prevValue <> "" And currValue <> "" Then
                    'Compare RPM values
                    If prevValue <> currValue Then
                        'Highlight cell in orange if there's a change
                        ws.Cells(j, currDayRPMCol).Interior.Color = RGB(255, 192, 0)
                    Else
                        'Remove any existing color if values are the same
                        ws.Cells(j, currDayRPMCol).Interior.ColorIndex = xlNone
                    End If
                Else
                    'Remove any existing color if either value is empty
                    ws.Cells(j, currDayRPMCol).Interior.ColorIndex = xlNone
                End If
            End If
        Next j
    Next i
    
    MsgBox "Analysis complete. RPM changes highlighted in orange."
End Sub
