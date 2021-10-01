Option Explicit

Public Function ArrayLen(arr As Variant) As Integer
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

Sub readFromPDF(name As String, outputSheet As String, baseExcelRow As Integer, baseExcelCol As Integer)
    Dim fpath As Variant
    Dim currentExcelRow As Integer, currentExcelCol As Integer
    
    Dim wApp As New Word.Application
    Dim wDoc As Word.Document
    Dim pg As Word.Paragraph
    Dim wLine As String
    
    Dim inTable As Boolean
    Dim currentTable As String, currentHeaderRow As Variant, currentRowLabels As Variant
    Dim tableRow As Integer
    Dim rowParts As Variant, rowParts_idx As Integer, rowParts_jdx As Integer, finalRowPart As String
    
    Dim headlineHeaderRow As Variant, headlineRowLabels As Variant
    Dim mainHeaderRow As Variant, mainRowLabels As Variant
    Dim underHeaderRow As Variant, underRowLabels As Variant
    
    '--- Set constants:
    headlineHeaderRow = Array("Name", "~", "Number")
    headlineRowLabels = Array("Dry Solids", "pH", "Density")
    mainHeaderRow = Array("~", "~", "Dry Basis", "Median", "Lower 90%", _
        "Upper 90%")
    mainRowLabels = Array("Hydrogen", "Helium", "Li", "Be", _
        "B", "C", "N", "O", "F", "Ne", "Sodium", "Magnesium", _
        "Aluminium", "Silicon", "Phosphorus", "Sulfur", "Chlorine", _
        "Argon", "K", "Ca", "Sc", "Ti", "V")
    underHeaderRow = Array("Calculations", "~", "~")
    underRowLabels = Array("Adjusted Crude Sugar", "Salt", "Flour", "Eggs")
    
    '--- Read PDF file into Word
    fpath = Application.GetOpenFilename(Title:="Open " & name & " file :^)", _
                                        FileFilter:="PDF Files (*.pdf),*.pdf")
    If fpath = False Then Exit Sub
    wApp.Visible = False
    Set wDoc = wApp.Documents.Open(filename:=fpath, Format:="PDF Files", _
                                   ConfirmConversions:=False, ReadOnly:=True)
    
    '--- Initialise important variables:
    Cells(baseExcelRow, baseExcelCol).Value = name
    Cells(baseExcelRow, baseExcelCol).Font.Bold = True
    Cells(baseExcelRow, baseExcelCol).Interior.ColorIndex = 34
    currentExcelRow = baseExcelRow + 1
    currentExcelCol = baseExcelCol
    
    currentTable = "Headline Table"
    currentHeaderRow = headlineHeaderRow
    currentRowLabels = headlineRowLabels
    inTable = False
    
    For Each pg In wDoc.Paragraphs
        wLine = WorksheetFunction.Trim(WorksheetFunction.Clean(pg.Range.Text))
        
        If Left(wLine, Len(currentRowLabels(0))) = currentRowLabels(0) Then
            'first row of main table detected
            currentExcelCol = baseExcelCol
            Cells(currentExcelRow, currentExcelCol).Value = currentTable
            Cells(currentExcelRow, currentExcelCol).Font.Bold = True
            currentExcelRow = currentExcelRow + 1
            For rowParts_idx = 0 To (ArrayLen(currentHeaderRow) - 1)
                Cells(currentExcelRow, currentExcelCol + rowParts_idx).Value = currentHeaderRow(rowParts_idx)
            Next rowParts_idx
            currentExcelRow = currentExcelRow + 1
            tableRow = 1
            inTable = True
        End If
        If inTable Then
            Debug.Print wLine
            If Left(wLine, Len(currentRowLabels(tableRow))) = currentRowLabels(tableRow) Then
                'on next line of main table -> move to new line in Excel
                tableRow = tableRow + 1
                currentExcelRow = currentExcelRow + 1
                currentExcelCol = baseExcelCol
            End If
            
            'first and last row need lots of special care >:|
            If tableRow = 1 And currentTable <> "Under Table" Then
                Cells(currentExcelRow, baseExcelCol).Value = currentRowLabels(tableRow - 1)
                rowParts = Split(Mid(wLine, Len(currentRowLabels(0)) + 1), " ")
                For rowParts_idx = 0 To (ArrayLen(rowParts) - 1)
                    Cells(currentExcelRow, baseExcelCol + 1 + rowParts_idx).Value = rowParts(rowParts_idx)
                Next rowParts_idx
            ElseIf tableRow = ArrayLen(currentRowLabels) Then
                Cells(currentExcelRow, baseExcelCol).Value = currentRowLabels(tableRow - 1)
                rowParts = Split(Mid(wLine, Len(currentRowLabels(tableRow - 1)) + 1), " ")
                currentExcelCol = baseExcelCol + 1
                For rowParts_idx = 0 To (ArrayLen(rowParts) - 2)
                    If rowParts(rowParts_idx) <> "\t" Then
                        Cells(currentExcelRow, currentExcelCol).Value = rowParts(rowParts_idx)
                        currentExcelCol = currentExcelCol + 1
                    End If
                Next rowParts_idx
                
                'need to insert spaces between numbers (uses different index var)
                finalRowPart = Replace(rowParts(ArrayLen(rowParts) - 1), "-", "")
                For rowParts_jdx = 0 To (Len(finalRowPart) / 4 - 1)
                    Cells(currentExcelRow, baseExcelCol + ArrayLen(rowParts) + rowParts_jdx).Value = Mid(finalRowPart, 4 * rowParts_jdx + 1, 4)
                Next rowParts_jdx
                
                'last row -> exit afterwards
                inTable = False
                currentExcelRow = currentExcelRow + 2
                If currentTable = "Headline Table" Then
                    currentTable = "Main Table"
                    currentHeaderRow = mainHeaderRow
                    currentRowLabels = mainRowLabels
                ElseIf currentTable = "Main Table" Then
                    currentTable = "Under Table"
                    currentHeaderRow = underHeaderRow
                    currentRowLabels = underRowLabels
                End If
            Else
                'all other rows are separated automatically
                If (currentExcelCol - baseExcelCol) > 2 Then
                    'first 2 columns are row labels, the rest should be numeric
                    wLine = Replace(wLine, "-", " ")
                End If
                Cells(currentExcelRow, currentExcelCol).Value = wLine
            End If
            
            currentExcelCol = currentExcelCol + 1
        End If
    Next pg
    
    'Autofit all used columns
    'With Worksheets("Sheet1")
    '    .Columns.AutoFit
    'End With
    
    wDoc.Close False
    wApp.Quit
End Sub
