ub TransferData()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim sourceSheetName As String, destSheetName As String
    Dim lastRowSource As Long, lastColSource As Long
    Dim sourceDateRange As Range, destDateCell As Range
    Dim wellNames As Variant
    Dim i As Long, j As Long, k As Long
    Dim currentDate As Date
    Dim sourceData As Variant
    Dim colOffset As Integer
    Dim sourceWellCell As Range
    
    ' Çàïðàøèâàåì èìåíà ëèñòîâ ó ïîëüçîâàòåëÿ
    sourceSheetName = Application.InputBox("Ââåäèòå èìÿ èñõîäíîãî ëèñòà:", "Ïåðåíîñ äàííûõ", "october", Type:=2)
    If sourceSheetName = "False" Then Exit Sub
    
    destSheetName = Application.InputBox("Ââåäèòå èìÿ öåëåâîãî ëèñòà:", "Ïåðåíîñ äàííûõ", "horizon", Type:=2)
    If destSheetName = "False" Then Exit Sub
    
    ' Set worksheets
    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets(sourceSheetName)
    Set wsDest = ThisWorkbook.Sheets(destSheetName)
    
    If wsSource Is Nothing Then
        MsgBox "Èñõîäíûé ëèñò '" & sourceSheetName & "' íå íàéäåí!"
        Exit Sub
    End If
    If wsDest Is Nothing Then
        MsgBox "Öåëåâîé ëèñò '" & destSheetName & "' íå íàéäåí!"
        Exit Sub
    End If
    
    ' Çàïðàøèâàåì ñïèñîê ñêâàæèí ó ïîëüçîâàòåëÿ
    Dim wellInput As String
    wellInput = Application.InputBox("Ââåäèòå èìåíà ñêâàæèí ÷åðåç çàïÿòóþ:", "Ñïèñîê ñêâàæèí", "well_1,well_2,well_3", Type:=2)
    If wellInput = "False" Then Exit Sub
    
    ' Ðàçäåëÿåì ââåä¸ííûå èìåíà ñêâàæèí
    wellNames = Split(wellInput, ",")
    For i = 0 To UBound(wellNames)
        wellNames(i) = Trim(wellNames(i))
    Next i
    
    ' Determine the last column with data in the source
    lastColSource = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column
    
    ' Loop through all dates in the source (every 9 columns)
    For i = 1 To lastColSource Step 9
        ' Check if cell contains a date
        If IsDate(wsSource.Cells(1, i).Value) Then
            currentDate = wsSource.Cells(1, i).Value
            
            ' Find the date in the destination sheet
            Set destDateCell = wsDest.Columns("A:A").Find(currentDate, LookIn:=xlValues)
            If Not destDateCell Is Nothing Then
                ' For each well, transfer data
                For j = 0 To UBound(wellNames)
                    ' Find the well in the source
                    Set sourceWellCell = wsSource.Columns(i).Find(wellNames(j), LookIn:=xlValues)
                    If Not sourceWellCell Is Nothing Then
                        ' Determine column offset for data
                        colOffset = j * 6 ' Each well gets 6 columns (B, H, N, T, Z, AF, AL, AR, AX)
                        
                        ' Copy data (Oil, Fluid, Water, Gas)
                        For k = 0 To 3
                            sourceData = wsSource.Cells(sourceWellCell.Row, i + 3 + k).Value
                            wsDest.Cells(destDateCell.Row, 2 + colOffset + k).Value = sourceData
                        Next k
                    End If
                Next j
            End If
        End If
    Next i
    
    MsgBox "Data transfer completed successfully!"
End Sub
