Const MAX = 200
Dim vulnerabilityArr(MAX), severityArr(MAX), paragraphArr(MAX), magicString As String ' initializing the arrays
Dim counter As Long
Dim headers As Variant
Dim objTable As Table
Dim tableIndex As Integer

Sub Main()
    ' faster way to turn off screen updating
    Application.ScreenUpdating = False
    LanguageSelection
    tableIndex = GetIndex()
    ' create headers variant from cross reference
    headers = ActiveDocument.GetCrossReferenceItems(wdRefTypeHeading)
    counter = UBound(headers)
    ModHeaderArray
    FillFindingsTable
    DeleteEmptyRowsFromTable
    DeleteUnnecessaryRowsFromTable
    DeleteUnnecessaryExecutiveSummary
    ' when everything done, we need to show changes again
    Application.ScreenUpdating = True
End Sub

Private Function LanguageSelection()
With ActiveDocument
    With Selection.Find
        .Text = "Risk"
        .Execute
            If .Found = True Then
                    'MsgBox ("English language found")
                    magicString = "Risk: "
            Else
            .Text = "Kockázat"
            .Execute
                If .Found = True Then
                    'MsgBox ("Hungarian language found")
                    magicString = "Kockázat: "
                Else
                    MsgBox ("Wrong language found!")
                End If
            End If
    End With
End With
End Function

' Modify header array for table elements
Private Function ModHeaderArray()
    Dim intItem As Integer
    Dim curHeaderStr, tmpSeverityLevel As String
    For intItem = 1 To counter
        curHeaderStr = Trim(headers(intItem))
        If (IsNumeric(Left(curHeaderStr, 1))) Then
            tmpSeverityLevel = getSeverity(ByVal curHeaderStr)
            tmpHolder = Split(curHeaderStr, " ", 2)
            paragraphArr(intItem - 1) = tmpHolder(0)
            vulnerabilityArr(intItem - 1) = tmpHolder(1)
            severityArr(intItem - 1) = tmpSeverityLevel
        End If
    Next intItem
End Function

' Fill table with header array
Private Function FillFindingsTable()
    Dim RowIndex As Integer
    ' Select findings table
    Set objTable = ActiveDocument.Tables(tableIndex)
    'MsgBox ("Target table index number = '" & tableIndex & "'")
    RowIndex = 2
    For i = 0 To counter
		If (Len(paragraphArr(RowIndex - 2)) = 1) Then ' when main paragraph's chapter length equal with one character 
        'If (Len(paragraphArr(RowIndex - 2)) < 4) Then ' when main paragraph's chapter length less then 4 characters (e.g. x.x)
            objTable.Rows.Add
            objTable.Cell(RowIndex, 1).Range.Text = paragraphArr(RowIndex - 2)
            With objTable.Cell(RowIndex, 1).Range
                .Words.First.Font.AllCaps = True
            End With
            objTable.Cell(RowIndex, 2).Range.Text = vulnerabilityArr(RowIndex - 2)
            With objTable.Cell(RowIndex, 2).Range
                .Sentences.First.Font.AllCaps = True
            End With
            objTable.Cell(RowIndex, 3).Range.Text = "-"
            RowIndex = RowIndex + 1
        Else
            objTable.Rows.Add
            objTable.Cell(RowIndex, 1).Range.Text = paragraphArr(RowIndex - 2)
            objTable.Cell(RowIndex, 2).Range.Text = vulnerabilityArr(RowIndex - 2)
            With objTable.Cell(RowIndex, 2).Range
                .Sentences.First.Font.AllCaps = False
            End With
            objTable.Cell(RowIndex, 3).Range.Text = severityArr(RowIndex - 2)
            RowIndex = RowIndex + 1
        End If
    Next
    Set objTable = Nothing
End Function
    
' contains risks (paragraph with magic string)
Private Function getSeverity(ByVal searchPara As String) As String
    getSeverity = "x"
    For Each curParagraph In ActiveDocument.Paragraphs
        If (Len(curParagraph.Range.Text) > 3) Then 
            If InStr(curParagraph.Range.Text, Split(searchPara, " ", 2)(1)) > 0 Then ' found the main paragraph (header)
                If InStr(curParagraph.Next.Range.Text, magicString) > 0 Then
                    getSeverity = Split(Left(curParagraph.Next.Range.Text, Len(curParagraph.Next.Range.Text) - 1), " ")(1)
                    Exit For
                End If
            End If
        End If
    Next
End Function

' Get Table Index for Findings table
Private Function GetIndex() As Integer
    With ActiveDocument
        For i = 1 To .Tables.Count
            If .Tables(i).Title = "findingstable" Then
                GetIndex = i
                Exit For
            Else
                If i = .Tables.Count Then
                    MsgBox ("Findings table did not found! Set it: Table Properties > Alt Text > Title > findingstable")
                End If
            End If
        Next
    End With
End Function

' Delete empty rows from Findings table
Private Function DeleteEmptyRowsFromTable()
    Dim cel As Cell, i As Long, n As Long, fEmpty As Boolean
    On Error GoTo ErrHandler
    Set objTable = ActiveDocument.Tables(tableIndex) ' Change the 'tableIndex' to whatever table # you want to process
    n = objTable.Rows.Count
    For i = n To 1 Step -1
        fEmpty = True
        For Each cel In objTable.Rows(i).Cells
            If Len(cel.Range.Text) > 2 Then
                fEmpty = False
                Exit For
            End If
        Next cel
            If fEmpty Then objTable.Rows(i).Delete
    Next i
    Set objTable = Nothing
ExitHandler:
        Set cel = Nothing
        Set objTable = Nothing
        Exit Function
ErrHandler:
        MsgBox Err.Description, vbExclamation
        Resume ExitHandler
End Function

' Delete rows from Findings table that contains 'x' on the risk column
Private Function DeleteUnnecessaryRowsFromTable()
    Dim i As Long, n As Long, ColumnValue As String
    Set objTable = ActiveDocument.Tables(tableIndex) ' Change the 'tableIndex' to whatever table # you want to process
    n = objTable.Rows.Count
    For i = n To 1 Step -1
        ColumnValue = objTable.Cell(i, 3).Range.Text
            If ColumnValue Like "x*" Then
                objTable.Rows(i).Delete
            End If
    Next i
    Set objTable = Nothing
End Function

' Delete Executive Summary rows from Findings table that contains '1*' on the chapter column
Private Function DeleteUnnecessaryExecutiveSummary()
    Dim i As Long, n As Long, ColumnValue As String
    Set objTable = ActiveDocument.Tables(tableIndex) ' Change the 'tableIndex' to whatever table # you want to process
    n = objTable.Rows.Count
    For i = n To 1 Step -1
        ColumnValue = objTable.Cell(i, 1).Range.Text
            ' Check rows where chapter begans "1"
            If ColumnValue Like "1*" Then
                objTable.Rows(i).Delete
            End If
    Next i
    Set objTable = Nothing
End Function